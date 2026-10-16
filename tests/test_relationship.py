"""Prior-relationship suggestion (memo_app/relationship.py) on a fake local portfolio."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from memo_app.extract import Extracted  # noqa: E402
from memo_app.facts import UserInputs, build_facts  # noqa: E402
from memo_app.relationship import find_prior_deals, names_from_files, surnames  # noqa: E402
from memo_app.settings import Settings, load_settings  # noqa: E402
from memo_app.storage import FileEntry, LocalStorage  # noqa: E402
from test_offline import CANNED, make_deal  # noqa: E402


def _touch(p: Path, text: str = "x"):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text)


def make_portfolio(root: Path) -> Path:
    port = root / "Current Portfolio"
    deal = make_deal(port)                                                   # 150 E Aurora St
    _touch(deal / "PFS & SREO" / "PFS Shaya Margareten.pdf")
    _touch(port / "12 Main St" / "PFS & SREO" / "PFS - S. MARGARETEN 2024.pdf")   # level 2, PFS folder
    _touch(port / "77 Oak Ave" / "Margareten Term Sheet signed.pdf")         # level 1 (deal root)
    _touch(port / "9 Elm St" / "PFS" / "PFS John Smith.pdf")                 # other sponsor
    _touch(port / "5 Pine Rd" / "Closing" / "Archive" / "Margareten.pdf")    # level 3 - never listed
    _touch(port / "3 Lee Ct" / "Resume" / "Bo Lee bio.pdf")                  # surname < 4 letters
    _touch(port / "notes.txt")                                               # a file, not a deal
    return deal


def test_names_from_files():
    e = lambda n: FileEntry(n, n, n, False, None)  # noqa: E731
    names = names_from_files([e("PFS Shaya Margareten.pdf"), e("Shaya bio.pdf"), e("Personal Financial Statement.pdf"),
                              e("Resume - JOHN SMITH (2025).docx")])
    assert names == ["Shaya Margareten", "Shaya", "JOHN SMITH"]
    assert surnames(names + ["Bo Lee"]) == ["margareten", "shaya", "smith"]


def test_find_prior_deals_two_deals_share_surname(tmp_path):
    deal = make_portfolio(tmp_path)
    hits = find_prior_deals(LocalStorage(), str(deal), ["Shaya Margareten"])
    assert hits == ["12 Main St (PFS & SREO/PFS - S. MARGARETEN 2024.pdf)",
                    "77 Oak Ave (Margareten Term Sheet signed.pdf)"]
    assert find_prior_deals(LocalStorage(), str(deal), ["Bo Lee"]) == []          # too short to match safely
    assert find_prior_deals(LocalStorage(), str(deal), ["John Smith"]) == ["9 Elm St (PFS/PFS John Smith.pdf)"]
    assert find_prior_deals(LocalStorage(), str(deal), ["Shaya Margareten"],
                            portfolio_root=str(tmp_path / "nowhere")) == []      # unreadable root -> no crash


def test_scan_suggests_and_facts_flag(tmp_path):
    import memo_app.pipeline as pipeline
    deal = make_portfolio(tmp_path)
    settings: Settings = load_settings()
    ui = UserInputs(deal_path=str(deal), loan_type="acquisition")
    prior = pipeline.scan(ui, settings).prior_deals
    assert [p.split(" (")[0] for p in prior] == ["12 Main St", "77 Oak Ave"]

    x = Extracted(term_sheet=CANNED["TermSheet"], pfs=CANNED["PFSResult"])
    f = build_facts(x, ui, settings, prior_matches=prior)
    assert f.prior_deals == prior and any("NOT worked with the sponsor" in fl.message for fl in f.flags)
    ui.prior_relationship = True
    f = build_facts(x, ui, settings, prior_matches=prior)
    assert not any("NOT worked with the sponsor" in fl.message for fl in f.flags)

    settings.raw.setdefault("memo", {})["prior_relationship_autodetect"] = False
    assert pipeline.scan(ui, settings).prior_deals == []
