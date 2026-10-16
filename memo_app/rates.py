"""Index rates for the "effective rate at closing" sentence.

    classify(name)            '1-Month CME SOFR' -> 'term_sofr_1m', 'Prime Rate' -> 'prime', ...
    display(name)             the phrase used after "floating based on the ..." in the memo
    fetch_sofr()              overnight SOFR from the NY Fed (public, no key)
    fetch_sofr_30d_avg()      30-day average SOFR from the NY Fed
    prime_rate(settings)      FRED DPRIME when FRED_API_KEY is set, else rules.prime_rate_pct
    resolve(...)              -> IndexQuote: which rate to use and where it came from
    effective_rate(...)       max(index + spread, floor)

Term SOFR (CME) and AMERIBOR have no free API, so they rely on the value the analyst types.
A fetch failure never raises: it returns None and the memo falls back to the floor rate (flagged).
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Callable, Optional

from .net import httpx

Log = Callable[[str], None]

NYFED_SOFR = "https://markets.newyorkfed.org/api/rates/secured/sofr/last/1.json"
NYFED_SOFR_AVG = "https://markets.newyorkfed.org/api/rates/secured/sofrai/last/1.json"
FRED_OBS = "https://api.stlouisfed.org/fred/series/observations"
TIMEOUT = 10.0

FROM_TERM_SHEET = "(from term sheet)"
FIXED = "Fixed rate"
# GUI choices -> canonical kind
CHOICES = {
    FROM_TERM_SHEET: "",
    "1M Term SOFR": "term_sofr_1m",
    "3M Term SOFR": "term_sofr_3m",
    "30-day Avg SOFR": "sofr_avg_30",
    "AMERIBOR": "ameribor",
    "WSJ Prime": "prime",
    FIXED: "fixed",
}
# kind -> phrase after "Castellan's interest rate is floating based on the ..."
DISPLAY = {
    "term_sofr_1m": "1-Month CME Term SOFR index",
    "term_sofr_3m": "3-Month CME Term SOFR index",
    "sofr_avg_30": "30-day Average SOFR index",
    "sofr": "Secured Overnight Financing Rate (SOFR)",
    "prime": "WSJ Prime Rate",
    "ameribor": "AMERIBOR index",
}
FETCHABLE = {"term_sofr_1m", "term_sofr_3m", "sofr_avg_30", "sofr", "prime"}


def classify(name: Optional[str]) -> str:
    """Canonical index kind for a GUI choice or a term-sheet phrase; '' when unknown."""
    if not name:
        return ""
    if name in CHOICES:
        return CHOICES[name]
    n = name.lower()
    if "fixed" in n:
        return "fixed"
    if "prime" in n:
        return "prime"
    if "ameribor" in n:
        return "ameribor"
    if "sofr" in n:
        if re.search(r"\b30[\s-]*day\b|average|avg", n):
            return "sofr_avg_30"
        if re.search(r"\b(3|three)[\s-]*(m|mo|month)\b|90[\s-]*day", n):
            return "term_sofr_3m"
        if "term" in n or re.search(r"\b(1|one)[\s-]*(m|mo|month)\b", n):
            return "term_sofr_1m"
        return "sofr"
    return ""


def display(name: Optional[str]) -> Optional[str]:
    """'WSJ Prime' -> 'WSJ Prime Rate'; unknown names keep their own wording plus ' index'."""
    if not name:
        return None
    kind = classify(name)
    if kind in DISPLAY:
        return DISPLAY[kind]
    return name if re.search(r"\b(index|rate)\s*$", name, re.I) else f"{name} index"


# --------------------------------------------------------------------------- fetchers
def _get_json(url: str, params: dict | None = None, client=None, timeout: float = TIMEOUT):
    if client is not None:
        r = client.get(url, params=params, timeout=timeout)
    else:
        r = httpx.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    return r.json()


def fetch_sofr(client=None, timeout: float = TIMEOUT, log: Log | None = None) -> float | None:
    """Latest overnight SOFR in percent (e.g. 4.31), or None if the NY Fed cannot be reached."""
    try:
        return float(_get_json(NYFED_SOFR, client=client, timeout=timeout)["refRates"][0]["percentRate"])
    except Exception as exc:
        if log:
            log(f"  (SOFR fetch failed: {exc})")
        return None


def fetch_sofr_30d_avg(client=None, timeout: float = TIMEOUT, log: Log | None = None) -> float | None:
    """Latest 30-day average SOFR in percent, or None."""
    try:
        return float(_get_json(NYFED_SOFR_AVG, client=client, timeout=timeout)["refRates"][0]["average30day"])
    except Exception as exc:
        if log:
            log(f"  (30-day average SOFR fetch failed: {exc})")
        return None


def fetch_prime_fred(api_key: str, client=None, timeout: float = TIMEOUT, log: Log | None = None) -> float | None:
    """WSJ Prime (FRED series DPRIME). Needs a free FRED API key."""
    if not api_key:
        return None
    try:
        data = _get_json(FRED_OBS, {"series_id": "DPRIME", "api_key": api_key, "file_type": "json",
                                    "sort_order": "desc", "limit": 5}, client=client, timeout=timeout)
        for obs in data.get("observations", []):
            if obs.get("value") not in (None, "", "."):
                return float(obs["value"])
    except Exception as exc:
        if log:
            log(f"  (FRED prime fetch failed: {exc})")
    return None


def prime_rate(settings, client=None, log: Log | None = None, fetch: bool = True) -> tuple[float | None, str]:
    """(rate, source). FRED when FRED_API_KEY is set, else rules.prime_rate_pct from settings.yaml."""
    key = getattr(settings, "fred_api_key", "") or os.getenv("FRED_API_KEY", "")
    if fetch and key:
        v = fetch_prime_fred(key, client=client, log=log)
        if v is not None:
            return v, "FRED DPRIME"
    v = settings.get("rules.prime_rate_pct") if settings is not None else None
    return (float(v), "settings.yaml rules.prime_rate_pct") if v not in (None, "") else (None, "")


# --------------------------------------------------------------------------- resolution
@dataclass
class IndexQuote:
    kind: str                     # canonical kind ('' = unknown)
    rate_pct: Optional[float]     # index value in percent, None = unknown
    source: str = ""              # where the value came from (shown in the log and the sources report)
    note: str = ""                # becomes a "check" flag when a fallback was used


def resolve(kind: str, user_rate: Optional[float], settings=None, model_rate: Optional[float] = None,
            fetch: bool = True, client=None, log: Log | None = None) -> IndexQuote:
    """Pick the index value: analyst's number > live / configured source > underwriting model > nothing."""
    if kind == "fixed":
        return IndexQuote(kind, None, "fixed rate - no index")
    if user_rate is not None:
        return IndexQuote(kind, float(user_rate), "entered by the analyst")

    if kind == "prime":
        v, src = prime_rate(settings, client=client, log=log, fetch=fetch)
        if v is not None:
            note = "" if src.startswith("FRED") else (
                f"Prime rate {v:.2f}% taken from settings.yaml (rules.prime_rate_pct) - confirm it is current.")
            return IndexQuote(kind, v, src, note)
    if kind == "sofr_avg_30" and fetch:
        v = fetch_sofr_30d_avg(client=client, log=log)
        if v is not None:
            return IndexQuote(kind, v, "NY Fed 30-day average SOFR")
    if kind == "sofr" and fetch:
        v = fetch_sofr(client=client, log=log)
        if v is not None:
            return IndexQuote(kind, v, "NY Fed SOFR")

    if model_rate is not None:
        return IndexQuote(kind, float(model_rate), "underwriting model",
                          f"Index rate {model_rate:.2f}% taken from the underwriting model - confirm it is current.")

    if kind in ("term_sofr_1m", "term_sofr_3m", "sofr_avg_30") and fetch:
        v = fetch_sofr(client=client, log=log)
        if v is not None:
            label = {"term_sofr_1m": "1M Term SOFR", "term_sofr_3m": "3M Term SOFR",
                     "sofr_avg_30": "30-day average SOFR"}[kind]
            return IndexQuote(kind, v, "NY Fed overnight SOFR (proxy)",
                              f"No {label} value entered; used overnight SOFR {v:.2f}% from the NY Fed as a proxy - "
                              f"enter the actual {label} rate to be exact.")
    if kind == "ameribor":
        return IndexQuote(kind, None, "", "AMERIBOR has no free data feed - enter the current AMERIBOR rate.")
    return IndexQuote(kind, None, "")


def effective_rate(index_rate: Optional[float], spread: Optional[float], floor: Optional[float]) -> Optional[float]:
    """max(index + spread, floor) when the index is known; the floor otherwise (None if neither)."""
    if index_rate is not None and spread:
        return max(index_rate + spread, floor or 0)
    return floor
