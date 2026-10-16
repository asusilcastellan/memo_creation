"""Accuracy rules found in the 280 E Palmetto Park Rd draft (Part D). Offline, no LLM."""
from __future__ import annotations

import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from memo_app import schemas as S  # noqa: E402
from memo_app.extract import Extracted, dedupe_side_by_side  # noqa: E402
from memo_app.facts import UserInputs, build_facts, city_state_from, property_phrase  # noqa: E402
from memo_app.locator import locate_all  # noqa: E402
from memo_app.render import render, strip_markers, when  # noqa: E402
from memo_app.settings import Settings, load_yaml_resource  # noqa: E402
from memo_app.storage import FileEntry  # noqa: E402
from test_offline import CANNED  # noqa: E402

SETTINGS = Settings(raw={"rules": {"strong_credit_min": 700, "prime_rate_pct": 7.5}})
BP = load_yaml_resource("templates", "boilerplate.yaml")


def _x(**changes) -> Extracted:
    base = {k: CANNED[n] for k, n in (("term_sheet", "TermSheet"), ("psa", "PSA"), ("model", "UWModel"),
                                         ("pfs", "PFSResult"), ("background", "BackgroundResult"),
                                         ("resume", "ResumeFacts"), ("appraisal", "Appraisal"),
                                         ("environmental", "Environmental"), ("insurance", "Insurance"))}
    for key, upd in changes.items():
        base[key] = base[key].model_copy(update=upd) if isinstance(upd, dict) and base.get(key) is not None else upd
    return Extracted(**base)


def _facts(x: Extracted, loan_type="acquisition", **ui):
    return build_facts(x, UserInputs(deal_path=".", loan_type=loan_type, **ui), SETTINGS)


def _risk(f, key) -> str:
    return strip_markers(render(next(r["text"] for r in BP["risks"] if r["key"] == key), f), f)


def _highlights(f) -> list[str]:
    return [strip_markers(render(h["text"], f), f) for h in BP["highlights"] if when(h.get("when", "true"), f)]


def _flags(f, text) -> list[str]:
    return [fl.message for fl in f.flags if text in fl.message]


# --------------------------------------------------------------------------- D2 participations
def test_pari_passu_uses_share_and_total_loan():
    x = _x(term_sheet={"issuing_lender": "Maxim Capital", "structure": "pari_passu", "lead_lender": "Maxim Capital",
                       "total_loan_amount": 70_000_000, "castellan_share_amount": 25_000_000, "blended_loan_amount": None,
                       "senior_loan_amount": None, "junior_loan_amount": None, "junior_lender_name": None},
           model={"senior_loan": 70_000_000, "junior_loan": None},
           appraisal={"as_is_value": 100_000_000, "as_complete_value": None})
    f = _facts(x)
    assert (f.loan, f.total_loan, f.structure, f.lead_lender) == (25_000_000, 70_000_000, "pari_passu", "Maxim Capital")
    assert round(f.memo_ltv, 2) == 0.70 and round(f.ltv, 2) == 0.25 and round(f.castellan_share_pct, 1) == 35.7
    opener = strip_markers(render(BP["openers"]["structure"]["pari_passu"], f), f)
    assert opener == ("Castellan Capital will provide a $25,000,000 pari passu participation in a $70,000,000 "
                      "senior loan led by Maxim Capital to")
    ltv = _risk(f, "ltv")
    assert "The $70,000,000 senior loan led by Maxim Capital, in which Castellan holds a $25,000,000 pari passu position" in ltv
    assert "Loan-to-Value ratio of 70%" in ltv
    assert any("pari passu position in a $70,000,000 senior loan" in h and "70%" in h for h in _highlights(f))
    assert not _flags(f, "differs from model senior loan")


def test_term_sheet_from_another_lender_means_participation():
    x = _x(term_sheet={"issuing_lender": "Maxim Capital", "structure": None, "total_loan_amount": 70_000_000,
                       "castellan_share_amount": None, "junior_loan_amount": None, "junior_lender_name": None})
    f = _facts(x)
    assert f.is_participation and f.lead_lender == "Maxim Capital" and _flags(f, "share of the participated loan")


# --------------------------------------------------------------------------- D3 values + LTV sanity
def test_ltv_sanity_189_percent_is_not_printed():
    f = _facts(_x(appraisal={"as_complete_value": 1_185_000}))
    assert round(f.memo_ltv * 100) == 189 and f.ltv_suspect
    risk = _risk(f, "ltv")
    assert "REVIEW: LTV could not be confirmed" in risk and "189%" not in risk
    hl = " ".join(_highlights(f))
    assert "189%" not in hl and "REVIEW: LTV could not be confirmed" in hl
    assert _flags(f, "LTV of 189% looks wrong")


def test_value_choice_condo_sellout_and_renovation_purpose():
    x = _x(appraisal={"property_type": "condominium", "project_qualifier": "Luxury", "as_is_value": 9_000_000,
                      "as_complete_value": 30_000_000, "gross_sellout_value": 42_000_000},
           model={"property_type": None, "capex_budget": None})
    f = _facts(x, loan_type="construction")
    assert (f.value_label, f.value, f.value_label_lower) == ("Gross Sellout", 42_000_000, "gross sellout")
    assert f.property_phrase == "a luxury condominium development"
    # no capex in the model, but the term sheet says renovation -> as-complete value
    x = _x(appraisal={"as_is_value": 4_000_000, "as_complete_value": None, "as_stabilized_value": 6_000_000},
           model={"capex_budget": None, "appraised_value": None},
           term_sheet={"loan_purpose": "acquisition and renovation"})
    f = _facts(x)
    assert f.value_label == "As Stabilized"
    x = _x(appraisal={"as_is_value": 4_000_000, "as_complete_value": None}, model={"capex_budget": None},
           term_sheet={"loan_purpose": "acquisition"})
    assert _facts(x).value_label == "As Is"


# --------------------------------------------------------------------------- D4 credit + background people
def test_credit_wording_793_647_and_non_sponsor_people():
    people = [S.PersonBackground(full_name="Shaya Margareten", credit_score=793, bankruptcies=False, judgments_or_liens=False,
                                 criminal_records=False, pending_civil_cases=False, past_due_accounts=False, adverse_summary=None),
              S.PersonBackground(full_name="Chaim Margareten", credit_score=647, bankruptcies=False, judgments_or_liens=False,
                                 criminal_records=False, pending_civil_cases=False, past_due_accounts=False, adverse_summary=None),
              S.PersonBackground(full_name="John Doe", credit_score=550, bankruptcies=True, judgments_or_liens=False,
                                 criminal_records=False, pending_civil_cases=False, past_due_accounts=True, adverse_summary=None)]
    x = _x(background=S.BackgroundResult(provider="Eaglescreen", people=people),
           term_sheet={"key_principals": ["Shaya Margareten", "Chaim Margareten"]})
    f = _facts(x)
    text = strip_markers(render(BP["sponsorship"]["background"], f), f)
    assert "credit scores of +793 and +647" in text and "strong" not in text and "550" not in text
    assert "John Doe" not in text and "John Doe" not in (f.sponsor_names or "")
    assert f.background_clean is True                                   # adverse items belong to a non-sponsor
    assert any("John Doe" in n and "bankruptcies" in n for n in f.background_other_people)
    assert _flags(f, "below 700") and _flags(f, "John Doe, who is not a sponsor")
    f2 = _facts(_x())
    assert "has a strong credit score of +793" in strip_markers(render(BP["sponsorship"]["background"], f2), f2)


# --------------------------------------------------------------------------- D5 environmental
def test_phase_ii_no_further_assessment_variant():
    p1 = CANNED["Environmental"].model_copy(update={"recognized_environmental_conditions": False,
                                                    "conditions_identified": "a REC associated with a former dry cleaner"})
    p2 = CANNED["Environmental"].model_copy(update={"phase": "phase_ii", "report_type": "Limited Phase II ESA",
                                                    "no_further_assessment": True, "recognized_environmental_conditions": False})
    x = _x(environmental=p1)
    x.environmental_reports = [p1, p2]
    f = _facts(x)
    assert f.env_status == "phase2_clear"
    assert ("Castellan engaged CBRE to conduct a Phase I Environmental Site Assessment on the property, which identified "
            "a REC associated with a former dry cleaner; a subsequent Limited Phase II ESA concluded that no further "
            "assessment is warranted.") in _risk(f, "environmental")
    # a noted condition alone (no recommendation) is not an "issue"
    x = _x(environmental=p1)
    assert _facts(x).env_status == "clean"


# --------------------------------------------------------------------------- D6 consistency with the one pager
def test_one_pager_numbers_tables_and_conflicts():
    tables = [S.Table(sheet=None, title="Loan Metrics", columns=[], rows=[["Loan Amount:", "$2,240,000"]], bold_rows=[]),
              S.Table(sheet=None, title="Loan Overview", columns=[], rows=[["Address:", "77 Mott Street, New York"]], bold_rows=[]),
              S.Table(sheet=None, title="Sales Comparables", columns=["Address", "$/SF", "Address", "$/SF"],
                      rows=[["1 Main Street", "$90", "1 Main Street", "$90"], ["2 Oak Avenue", "$70", "2 Oak Avenue", "$70"]],
                      bold_rows=[])]
    op = S.OnePager(tables=tables, property_address="150 E Aurora St", structure="senior", lead_lender=None,
                    total_loan_amount=None, castellan_share_amount=None, castellan_share_pct=None, loan_psf=25.0,
                    comps_avg_psf=95.0, index_name="1M Term SOFR", spread_pct=5.81, floor_rate_pct=8.00)
    x = _x()
    x.one_pager = op
    f = _facts(x)
    assert [t.title for t in x.one_pager.tables] == ["Loan Metrics", "Sales Comparables"]   # 77 Mott table dropped
    assert _flags(f, "'Loan Overview' mentions 77 Mott Street")
    assert f.comps_psf == 95.0 and f.loan_psf == 25.0
    assert any("$25/SF" in h and "$95/SF" in h for h in _highlights(f))
    assert _flags(f, "comps average $95/SF differs from the appraisal's $70/SF")
    assert _flags(f, "Term sheet floor 9.64% vs one pager 8.00%")
    deduped = dedupe_side_by_side(tables[2])
    assert deduped.columns == ["Address", "$/SF"] and deduped.rows == [["1 Main Street", "$90"], ["2 Oak Avenue", "$70"]]


# --------------------------------------------------------------------------- D7 city/state + property type
def test_city_state_and_property_phrase():
    assert city_state_from("280 E Palmetto Park Rd, Boca Raton, Florida 33432") == "Boca Raton, FL"
    assert city_state_from("150 E Aurora St, Waterbury, CT 06708") == "Waterbury, CT"
    assert city_state_from("150 E Aurora St Waterbury, CT 06708") == "150 E Aurora St Waterbury, CT"
    assert city_state_from("1 Main St, New York, New York") == "New York, NY"
    assert property_phrase("industrial") == "an industrial property"
    assert property_phrase("condominium", "Luxury", development=True) == "a luxury condominium development"
    assert property_phrase("office", "Class A") == "a Class A office property"
    f = _facts(_x(), property_type="condo")
    assert f.property_type == "condominium"


# --------------------------------------------------------------------------- D8 model discovery
def _entries(paths: list[str]) -> list[FileEntry]:
    out, dirs = [], set()
    for p in paths:
        parts = p.split("/")
        for i in range(1, len(parts)):
            dirs.add("/".join(parts[:i]))
        out.append(FileEntry(p, p, parts[-1], False, datetime(2025, 1, 1, tzinfo=timezone.utc)))
    out += [FileEntry(d, d, d.split("/")[-1], True, datetime(2025, 1, 1, tzinfo=timezone.utc)) for d in dirs]
    return out


def test_model_discovery_order():
    rules = load_yaml_resource("config", "sources.yaml")
    pick = lambda paths: locate_all(_entries(paths), rules)["model"].files  # noqa: E731
    assert [e.rel for e in pick(["Underwriting/280 Palmetto UW.xlsx", "Memo/280 Palmetto one pager.xlsx"])] == \
        ["Underwriting/280 Palmetto UW.xlsx"]
    assert [e.rel for e in pick(["280 Palmetto Sizing.xlsx", "Underwriting/old model.xlsx"])] == ["280 Palmetto Sizing.xlsx"]
    assert [e.rel for e in pick(["Memo/280 Palmetto one pager.xlsx", "Memo/280 Palmetto Pro Forma.xlsm"])] == \
        ["Memo/280 Palmetto Pro Forma.xlsm"]
    assert pick(["Memo/280 Palmetto one pager.xlsx"]) == []
