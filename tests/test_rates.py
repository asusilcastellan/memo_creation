"""Index rate / effective rate rules (memo_app/rates.py + facts). Offline: HTTP is mocked."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from memo_app import rates  # noqa: E402
from memo_app.extract import Extracted  # noqa: E402
from memo_app.facts import UserInputs, build_facts  # noqa: E402
from memo_app.net import httpx  # noqa: E402
from memo_app.render import render, strip_markers  # noqa: E402
from memo_app.settings import Settings, load_yaml_resource  # noqa: E402
from test_offline import CANNED  # noqa: E402

SETTINGS = Settings(raw={"rules": {"prime_rate_pct": 7.50}})


def _client(routes: dict[str, object], status: int = 200):
    def handler(req):
        for frag, body in routes.items():
            if frag in str(req.url):
                return httpx.Response(status, json=body)
        return httpx.Response(404, json={})
    return httpx.Client(transport=httpx.MockTransport(handler))


NYFED = _client({"/sofr/last/1.json": {"refRates": [{"effectiveDate": "2026-10-05", "type": "SOFR", "percentRate": 4.31}]},
                 "/sofrai/last/1.json": {"refRates": [{"type": "SOFRAI", "average30day": 4.25}]}})


def _facts(override="", rate=None, quote=None, **ts_changes):
    ts = CANNED["TermSheet"].model_copy(update=ts_changes)
    x = Extracted(term_sheet=ts, model=CANNED["UWModel"])
    ui = UserInputs(deal_path=".", loan_type="acquisition", index_name_override=override, index_rate_pct=rate)
    return build_facts(x, ui, SETTINGS, index_quote=quote)


def _risk_text(f) -> str:
    bp = load_yaml_resource("templates", "boilerplate.yaml")
    tpl = next(r["text"] for r in bp["risks"] if r["key"] == "interest_rate")
    return strip_markers(render(tpl, f), f)


# --------------------------------------------------------------------------- classify / display
def test_classify_and_display():
    assert rates.classify("1-Month CME SOFR") == "term_sofr_1m"
    assert rates.classify("One Month Term SOFR") == "term_sofr_1m"
    assert rates.classify("3M Term SOFR") == "term_sofr_3m"
    assert rates.classify("30-Day Average SOFR") == "sofr_avg_30"
    assert rates.classify("Prime Rate as published in the WSJ") == "prime"
    assert rates.classify("AMERIBOR") == "ameribor"
    assert rates.classify("Fixed rate") == "fixed"
    assert rates.classify("SONIA") == ""
    assert rates.display("WSJ Prime") == "WSJ Prime Rate"
    assert rates.display("AMERIBOR") == "AMERIBOR index"
    assert rates.display("30-day Avg SOFR") == "30-day Average SOFR index"
    assert rates.display("Some Bank Base Rate") == "Some Bank Base Rate"


# --------------------------------------------------------------------------- fetchers
def test_fetchers_parse_nyfed_and_never_raise():
    assert rates.fetch_sofr(client=NYFED) == 4.31
    assert rates.fetch_sofr_30d_avg(client=NYFED) == 4.25
    broken = _client({"sofr": {"error": "down"}}, status=503)
    assert rates.fetch_sofr(client=broken) is None


def test_prime_from_settings_and_fred():
    v, src = rates.prime_rate(SETTINGS, fetch=False)
    assert v == 7.50 and "settings.yaml" in src
    fred = _client({"fred/series/observations": {"observations": [{"date": "2026-10-06", "value": "7.25"}]}})
    s = Settings(fred_api_key="k", raw=SETTINGS.raw)
    assert rates.prime_rate(s, client=fred) == (7.25, "FRED DPRIME")


# --------------------------------------------------------------------------- effective rate per index
def test_effective_rate_term_sofr_user_value():
    f = _facts(rate=4.31)                                   # term sheet: 1-Month CME SOFR + 5.81, floor 9.64
    assert round(f.effective_rate, 2) == 10.12 and f.index_rate_source == "entered by the analyst"
    assert "floating based on the 1-Month CME Term SOFR index plus a spread of 5.81%" in _risk_text(f)


def test_effective_rate_term_sofr_falls_back_to_overnight_sofr():
    q = rates.resolve("term_sofr_1m", None, SETTINGS, client=NYFED)
    assert q.rate_pct == 4.31 and "proxy" in q.source
    f = _facts(quote=q)
    assert round(f.effective_rate, 2) == 10.12
    assert any(fl.level == "check" and "overnight SOFR" in fl.message for fl in f.flags)


def test_effective_rate_30d_avg_sofr():
    q = rates.resolve("sofr_avg_30", None, SETTINGS, client=NYFED)
    assert q.rate_pct == 4.25 and not q.note


def test_effective_rate_prime():
    f = _facts(override="WSJ Prime", index_name="Prime Rate")
    assert f.index_rate == 7.50 and round(f.effective_rate, 2) == 13.31      # 7.50 + 5.81 > floor 9.64
    assert "floating based on the WSJ Prime Rate plus a spread" in _risk_text(f)
    assert not any("contradicts" in fl.message for fl in f.flags)


def test_effective_rate_ameribor():
    f = _facts(override="AMERIBOR", rate=3.50)               # 3.50 + 5.81 = 9.31 < floor -> floor
    assert f.effective_rate == 9.64
    assert "based on the AMERIBOR index" in _risk_text(f)
    assert any("contradicts the term sheet (1-Month CME SOFR)" in fl.message for fl in f.flags)
    f = _facts(override="AMERIBOR")                         # manual only: no value -> floor + flag
    assert f.effective_rate == 9.64 and f.index_rate is None
    assert any("AMERIBOR has no free data feed" in fl.message for fl in f.flags)


def test_fixed_rate_override_and_fixed_term_sheet():
    f = _facts(override="Fixed rate")
    assert f.is_floating is False and f.effective_rate is None and f.index_name is None
    assert any("contradicts" in fl.message for fl in f.flags)
    f = _facts(is_floating=False, index_name=None, spread_pct=None, fixed_rate_pct=11.5)
    assert f.is_floating is False and "fixed interest rate of 11.50%" in _risk_text(f)
    f = _facts(override="WSJ Prime", is_floating=False, index_name=None, spread_pct=None, fixed_rate_pct=11.5)
    assert f.is_floating is False                            # a fixed term sheet wins, but it is flagged
    assert any("contradicts the term sheet (fixed rate)" in fl.message for fl in f.flags)


def test_missing_index_value_uses_floor():
    f = _facts()                                            # no value typed, no fetch, model has none
    assert f.index_rate is None and f.effective_rate == 9.64
    assert any("set to the floor rate (no 1-Month CME SOFR value given)" in fl.message for fl in f.flags)
    assert rates.effective_rate(None, 5.81, None) is None
