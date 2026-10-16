"""Merge the extracted documents into one DealFacts object and apply Castellan's rules.

Source priority follows the walkthrough with Vasist:
  loan amounts / rate / recourse  -> signed term sheet, then model
  purchase price                  -> executed PSA, then model (mismatch is flagged)
  capex, sources & uses           -> model
  values, size, comps, market     -> appraisal
  sponsor financials              -> PFS ; background & credit -> Eaglescreen / credit report
Rules from the template comments are implemented in `build_facts` (search "RULE").
"""
from __future__ import annotations

import itertools
import re
import threading
from dataclasses import asdict, dataclass, field
from datetime import date
from typing import Any, Optional

from . import rates
from . import schemas as S
from .extract import Extracted

LOAN_TYPES = {
    "Acquisition": "acquisition",
    "Construction": "construction",
    "Refinance": "refinance",
    "Note-on-Note": "note_on_note",
}


@dataclass
class UserInputs:
    deal_path: str
    loan_type: str                      # acquisition | construction | refinance | note_on_note
    property_type: str = ""             # blank = detect
    exit_strategy: str = ""             # blank = let the writer infer from notes/term sheet
    prior_relationship: bool = False
    notes: str = ""                     # pasted broker / borrower emails, purpose of the loan
    index_name_override: str = ""       # GUI "Index" choice; "" = use the term sheet's index
    index_rate_pct: Optional[float] = None   # current value of that index, to compute the effective rate
    run_web_search: bool = False
    output_dir: str = ""
    read_graph: bool = False            # also read the Outlook thread through Microsoft Graph
    conversation_id: str = ""           # optional Outlook conversation ID (skips the Message-ID lookup)


# Memo sections used for flags (the review workbook's "Memo section" column; same names as sources.yaml memo_sections).
TO, BP, SUB = "Transaction Overview", "Business Plan", "Submarket Overview"
SP_BB, SP_FIN, SP_BG = "Sponsorship - Business background", "Sponsorship - Personal financials", "Sponsorship - Background & credit"
R_LTV, R_SP, R_CON = "Risks - LTV", "Risks - Sponsor", "Risks - Construction Completion"
R_ENV, R_IR, R_INS = "Risks - Environmental", "Risks - Interest Rate", "Risks - Insurance"
LF, HL, PICS, MAP, GEN = "Loan Financials", "Investment Highlights", "Pictures", "Map", "General"

REMARK_MAX_WORDS = 15
UID = "\u2063"                                 # invisible separator around a flag uid inside a red remark
_uids = itertools.count(1)
_flag_lock = threading.Lock()

# [MISSING] values in the templates: plain name, memo section and source document of each field.
FIELD_LABELS = {
    "loan": ("Castellan loan amount", "term_sheet"), "total_loan": ("total loan amount", "term_sheet"),
    "lead_lender": ("lead lender", "term_sheet"), "junior_lender": ("junior lender name", "term_sheet"),
    "junior_loan": ("junior loan amount", "term_sheet"), "value": ("appraised value", "appraisal"),
    "value_label_quoted": ("appraised value type", "appraisal"), "value_label_lower": ("appraised value type", "appraisal"),
    "appraiser": ("appraiser", "appraisal"), "ltv": ("LTV", "appraisal"), "memo_ltv": ("LTV", "appraisal"),
    "property_phrase": ("property type", "appraisal"), "city_state": ("city and state", "appraisal"),
    "address": ("property address", "appraisal"), "loan_psf": ("loan per SF", "appraisal"),
    "comps_psf": ("comps $/SF", "appraisal"), "sponsor_names": ("sponsor names", "pfs"),
    "sponsor_ref": ("sponsor name", "pfs"), "net_worth": ("sponsor net worth", "pfs"), "liquid": ("sponsor liquidity", "pfs"),
    "real_estate": ("sponsor real estate assets", "pfs"), "total_liquidity": ("sponsor liquidity", "pfs"),
    "equity": ("sponsor equity", "model"), "equity_pct": ("equity %", "model"), "fixed_rate": ("fixed rate", "term_sheet"),
    "index_display": ("interest rate index", "term_sheet"), "spread": ("spread", "term_sheet"),
    "effective_rate": ("effective rate at closing", "term_sheet"), "floor": ("floor rate", "term_sheet"),
    "env_provider": ("environmental consultant", "environmental"), "env_report_phrase": ("environmental report type", "environmental"),
    "env_conditions": ("Phase I conditions", "environmental"), "env_phase2_type": ("Phase II report type", "environmental"),
    "property_coverage": ("property insurance amount", "insurance"), "liability_occurrence": ("liability insurance amount", "insurance"),
    "pcr_firm": ("plan & cost review firm", ""), "construction_reserve": ("construction reserve", "term_sheet"),
    "background_provider": ("background check provider", "background"), "insurance_agent": ("insurance agent", ""),
}


def short_title(text: str, max_words: int = REMARK_MAX_WORDS) -> str:
    words = str(text or "").split()
    return " ".join(words[:max_words]).rstrip(",;:.") + ("\u2026" if len(words) > max_words else "")


@dataclass
class Flag:
    """One item for the analyst. The SAME objects feed the red remarks in the .docx and the review workbook."""
    level: str                                  # review | missing | check
    details: str                                # full text (workbook "Details")
    title: str = ""                             # short text shown in the memo remark and the workbook "Item"
    memo_section: str = ""
    source_key: str = ""                        # sources.yaml key, or ""
    source_files: list = field(default_factory=list)   # relative paths inside the deal folder
    key: str = ""                               # code the boilerplate links to with f.remark('<key>')
    id: str = ""                                # R1, R2, ... assigned in document order when the docx is built
    inline: bool = False                        # True when it appears as a red remark in the memo body
    uid: int = 0

    @property
    def message(self) -> str:                   # older code / tests
        return self.details

    @property
    def level_label(self) -> str:
        return self.level.title()

    def remark_text(self) -> str:
        return f"[{self.level.upper()}{' ' + self.id if self.id else ''}: {self.title}]"


@dataclass
class SponsorView:
    name: str
    ref: str
    net_worth: Optional[float]
    liquid: Optional[float]
    real_estate: Optional[float]
    mention_cash: bool


@dataclass
class DealFacts:
    loan_type: str
    memo_date: str
    address: Optional[str] = None
    city_state: Optional[str] = None
    property_type: Optional[str] = None
    property_phrase: Optional[str] = None
    property_summary: Optional[str] = None
    units: Optional[int] = None
    unit_mix: Optional[str] = None
    gross_sf: Optional[float] = None
    net_sf: Optional[float] = None
    sf: Optional[float] = None
    # loan
    loan: Optional[float] = None                # Castellan's own loan (its share in a participation)
    structure: Optional[str] = None             # senior | pari_passu | participation | mezz | note_on_note
    is_participation: bool = False
    lead_lender: Optional[str] = None
    total_loan: Optional[float] = None          # the whole participated senior loan
    castellan_share_pct: Optional[float] = None
    blended_loan: Optional[float] = None
    has_mezz: bool = False
    junior_lender: Optional[str] = None
    junior_loan: Optional[float] = None
    total_debt: Optional[float] = None
    loan_term_months: Optional[int] = None
    closing_date: Optional[str] = None
    security_summary: Optional[str] = None
    # project
    purchase_price: Optional[float] = None
    capex: Optional[float] = None
    closing_costs: Optional[float] = None
    total_cost: Optional[float] = None
    initial_funding: Optional[float] = None
    construction_reserve: Optional[float] = None
    construction_reserve_terms: Optional[str] = None
    pcr_firm: Optional[str] = None
    # value
    value: Optional[float] = None
    value_label: Optional[str] = None
    value_label_lower: Optional[str] = None
    value_label_quoted: Optional[str] = None
    all_values: dict = field(default_factory=dict)
    appraiser: Optional[str] = None
    ltv: Optional[float] = None                 # Castellan's loan / value
    ltv_total: Optional[float] = None           # total debt / value
    memo_ltv: Optional[float] = None            # the LTV the memo prints (total senior loan for a participation)
    ltv_suspect: bool = False                   # RULE: LTV > 100% or < 10% is not printed
    ltc: Optional[float] = None
    ltc_total: Optional[float] = None
    debt_to_price: Optional[float] = None
    loan_psf: Optional[float] = None
    comps_psf: Optional[float] = None
    # sponsors
    sponsors: list = field(default_factory=list)
    sponsor_names: Optional[str] = None
    sponsor_ref: Optional[str] = None
    plural: bool = False
    total_liquidity: Optional[float] = None
    liquidity_highlight: bool = False
    background_provider: str = "Eaglescreen"
    background_clean: Optional[bool] = None
    background_issues: Optional[str] = None
    credit_scores: list = field(default_factory=list)
    credit_strong: bool = False                 # RULE: every score >= rules.strong_credit_min
    background_other_people: list = field(default_factory=list)   # review notes only, never named in the memo
    no_past_due: bool = False
    online_search_clean: Optional[bool] = None
    online_search_summary: Optional[str] = None
    online_findings: list = field(default_factory=list)   # AdverseFinding dicts (review workbook "Online search")
    prior_relationship: bool = False
    prior_deals: list = field(default_factory=list)   # suggested by relationship.find_prior_deals
    sponsor_experience: str = ""
    personal_guaranty: bool = False
    completion_guaranty: bool = False
    carveout_guaranty: bool = False
    # equity
    equity: Optional[float] = None
    equity_pct: Optional[float] = None
    equity_highlight: bool = False
    # rate
    is_floating: Optional[bool] = None
    index_name: Optional[str] = None
    index_kind: str = ""                     # rates.classify(): term_sofr_1m, prime, ameribor, fixed ...
    index_display: Optional[str] = None      # "WSJ Prime Rate", "AMERIBOR index" (used in the memo sentence)
    index_rate: Optional[float] = None       # index value used for the effective rate
    index_rate_source: str = ""
    spread: Optional[float] = None
    floor: Optional[float] = None
    fixed_rate: Optional[float] = None
    rate_cap: Optional[str] = None
    effective_rate: Optional[float] = None
    interest_reserve: Optional[float] = None
    facility_leverage: Optional[float] = None
    facility_rate: Optional[float] = None
    # third-party
    env_status: str = "missing"
    env_provider: Optional[str] = None
    env_report_phrase: Optional[str] = None
    env_summary: Optional[str] = None
    env_conditions: Optional[str] = None        # what the Phase I identified (Phase II "no further assessment" variant)
    env_phase2_type: Optional[str] = None
    insurance_mode: str = "property"
    insurance_bound: bool = False
    property_coverage: Optional[float] = None
    liability_occurrence: Optional[float] = None
    umbrella: Optional[float] = None
    builders_risk: Optional[float] = None
    insurance_agent: str = "IMA Corp"
    # context for the writer
    exit_strategy: str = ""
    exit_strategy_source: str = ""      # analyst | email chain
    email_insights: dict = field(default_factory=dict)   # EmailInsights from the email thread (formatted for the writer)
    notes: str = ""
    flags: list = field(default_factory=list)

    # ------------------------------------------------------------------ analyst items
    def flag(self, level: str, details: str, title: str | None = None, section: str = "", source: str = "",
             key: str = "") -> Flag:
        fl = Flag(level, details, short_title(title or details), section or GEN, source, key=key)
        with _flag_lock:
            fl.uid = next(_uids)
            self.flags.append(fl)
        return fl

    def find(self, key: str) -> list[Flag]:
        return [fl for fl in self.flags if fl.key == key or fl.key.startswith(key + ":")]

    @staticmethod
    def marker(fl: Flag) -> str:
        """A red remark in memo text that points at `fl` (rendered as e.g. [REVIEW R3: title] in the .docx)."""
        return f"\u27ea{UID}{fl.uid}{UID}\u27eb"

    def remark(self, key: str, level: str = "review", title: str = "", section: str = "", source: str = "") -> str:
        """Boilerplate: {{ f.remark('background_adverse') }} -> red remark linked to the flag raised in build_facts."""
        found = self.find(key)
        fl = found[0] if found else self.flag(level, title or key.replace("_", " "), title or key.replace("_", " "),
                                               section, source, key=key)
        return self.marker(fl)

    def new_remark(self, level: str, title: str, details: str = "", section: str = "", source: str = "",
                   key: str = "") -> str:
        if key and self.find(key):
            return self.marker(self.find(key)[0])
        return self.marker(self.flag(level, details or title, title, section, source, key=key))

    def missing_remark(self, name: str, section: str = "") -> str:
        """A template value that is None -> [MISSING Rn: <plain name>] (one flag per field and section)."""
        label, source = FIELD_LABELS.get(name, (name.replace("_", " "), ""))
        key = f"missing:{name}:{section}"
        return self.new_remark("missing", label, f"{label[:1].upper() + label[1:]} could not be determined from the deal "
                               f"documents{' (' + section + ')' if section else ''}.", section, source, key=key)

    def for_prompt(self) -> dict[str, Any]:
        """Facts handed to the writer model (no internal bookkeeping)."""
        d = asdict(self)
        for k in ("flags", "notes", "sponsor_experience", "index_kind", "index_rate_source", "prior_deals",
                  "email_insights", "exit_strategy_source", "background_other_people", "ltv_suspect",
                  "online_findings"):
            d.pop(k, None)
        return {k: v for k, v in d.items() if v not in (None, "", [], {}, False) or k in ("has_mezz",)}


# --------------------------------------------------------------------------- helpers
def _first(*vals):
    for v in vals:
        if v not in (None, "", 0):
            return v
    return None


def _div(a, b):
    return a / b if a and b else None


STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR", "california": "CA", "colorado": "CO",
    "connecticut": "CT", "delaware": "DE", "district of columbia": "DC", "florida": "FL", "georgia": "GA",
    "hawaii": "HI", "idaho": "ID", "illinois": "IL", "indiana": "IN", "iowa": "IA", "kansas": "KS", "kentucky": "KY",
    "louisiana": "LA", "maine": "ME", "maryland": "MD", "massachusetts": "MA", "michigan": "MI", "minnesota": "MN",
    "mississippi": "MS", "missouri": "MO", "montana": "MT", "nebraska": "NE", "nevada": "NV", "new hampshire": "NH",
    "new jersey": "NJ", "new mexico": "NM", "new york": "NY", "north carolina": "NC", "north dakota": "ND",
    "ohio": "OH", "oklahoma": "OK", "oregon": "OR", "pennsylvania": "PA", "rhode island": "RI",
    "south carolina": "SC", "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT",
    "virginia": "VA", "washington": "WA", "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
    "puerto rico": "PR",
}
_STATE_CODES = set(STATES.values())


def city_state_from(address: Optional[str]) -> Optional[str]:
    """'280 E Palmetto Park Rd, Boca Raton, Florida 33432' -> 'Boca Raton, FL' (USPS state codes)."""
    if not address:
        return None
    parts = [p.strip() for p in address.split(",") if p.strip()]
    for i in range(len(parts) - 1, 0, -1):
        words = re.sub(r"\b\d{5}(?:-\d{4})?\b", "", parts[i]).replace(".", "").strip()
        code = words.upper() if words.upper() in _STATE_CODES else STATES.get(words.lower())
        if not code:   # "Waterbury CT 06708" in one part
            m = re.match(r"(.+?)\s+([A-Za-z][A-Za-z ]+)$", words)
            if m and (m.group(2).upper() in _STATE_CODES or m.group(2).lower() in STATES):
                return f"{m.group(1).strip()}, {m.group(2).upper() if m.group(2).upper() in _STATE_CODES else STATES[m.group(2).lower()]}"
            continue
        return f"{parts[i - 1]}, {code}" if i >= 1 else code
    return ", ".join(parts[-2:]) if len(parts) >= 2 else address


def article(word: str) -> str:
    return "an" if word[:1].lower() in "aeiou" else "a"


PROPERTY_SYNONYMS = {
    "condo": "condominium", "condos": "condominium", "condominiums": "condominium", "apartment": "multifamily",
    "apartments": "multifamily", "multi-family": "multifamily", "warehouse": "industrial", "flex": "industrial",
    "logistics": "industrial", "hotel": "hospitality", "motel": "hospitality", "single family": "single-family",
    "sfr": "single-family", "mixed use": "mixed-use", "self storage": "self-storage", "storage": "self-storage",
}
PROPERTY_TYPES = {"multifamily", "industrial", "retail", "office", "single-family", "condominium", "mixed-use", "land",
                  "hospitality", "self-storage", "other"}


def normalize_property_type(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    v = value.strip().lower()
    v = PROPERTY_SYNONYMS.get(v, v)
    return v if v in PROPERTY_TYPES else "other"


def property_phrase(ptype: Optional[str], qualifier: Optional[str] = None, development: bool = False) -> Optional[str]:
    """('condominium', 'luxury', True) -> 'a luxury condominium development'; ('industrial',) -> 'an industrial property'."""
    if not ptype or ptype == "other":
        return None
    q = (qualifier or "").strip()
    if q and not any(c.isupper() for c in q[1:]):
        q = q.lower()                              # 'Luxury' -> 'luxury', but keep 'Class A'
    if q.lower() in (ptype, ptype + "s"):
        q = ""
    noun = "parcel" if ptype == "land" else ("development" if development else "property")
    words = " ".join(w for w in (q, ptype, noun) if w)
    return f"{article(words)} {words}"


def _person_ref(full: str, honorific: Optional[str], last: Optional[str]) -> str:
    last = last or full.split()[-1]
    return f"{honorific} {last}" if honorific else full


def _join(names: list[str]) -> str:
    if len(names) <= 1:
        return "".join(names)
    return ", ".join(names[:-1]) + " and " + names[-1]


# --------------------------------------------------------------------------- main
def build_facts(x: Extracted, ui: UserInputs, settings, index_quote: Optional[rates.IndexQuote] = None,
                prior_matches: Optional[list[str]] = None) -> DealFacts:
    rules = settings.get("rules", {}) or {}
    memo_cfg = settings.get("memo", {}) or {}
    ts, psa, md, ap = x.term_sheet, x.psa, x.model, x.appraisal
    f = DealFacts(loan_type=ui.loan_type, memo_date=memo_date())
    f.notes, f.exit_strategy, f.prior_relationship = ui.notes, ui.exit_strategy, ui.prior_relationship
    f.exit_strategy_source = "analyst" if ui.exit_strategy else ""
    f.insurance_agent = memo_cfg.get("insurance_agent", "IMA Corp")
    f.background_provider = (x.background.provider if x.background and x.background.provider
                             else memo_cfg.get("background_provider", "Eaglescreen"))

    # ---- property ---------------------------------------------------------
    f.address = _first(ap and ap.property_address, md and md.property_address,
                       ts and ts.property_address, psa and psa.property_address)
    f.city_state = city_state_from(f.address)
    f.property_type = normalize_property_type(ui.property_type or _first(md and md.property_type, ap and ap.property_type))
    f.property_phrase = property_phrase(f.property_type, ap.project_qualifier if ap else None,
                                        development=ui.loan_type == "construction")
    f.property_summary = _first(ap and ap.property_description, ts and ts.property_description)
    f.units = _first(md and md.units, ap and ap.units)
    f.unit_mix = md.unit_mix if md else None
    f.gross_sf = _first(ap and ap.gross_building_sf, md and md.gross_sf)
    f.net_sf = _first(ap and ap.net_rentable_sf, md and md.net_rentable_sf)
    f.sf = _first(f.net_sf, f.gross_sf)

    # ---- loan amounts (RULE: signed term sheet; blended - junior = Castellan) ----
    if ts:
        f.blended_loan = ts.blended_loan_amount
        f.junior_loan = ts.junior_loan_amount
        f.junior_lender = ts.junior_lender_name
        f.loan = _first(ts.senior_loan_amount,
                        (ts.blended_loan_amount - (ts.junior_loan_amount or 0)) if ts.blended_loan_amount else None)
        f.loan_term_months = ts.loan_term_months
        f.closing_date = ts.estimated_closing_date
        f.security_summary = ts.security_summary
        f.personal_guaranty = bool(ts.personal_guaranty)
        f.completion_guaranty = bool(ts.completion_guaranty)
        f.carveout_guaranty = bool(ts.carveout_guaranty)
        f.construction_reserve = ts.construction_reserve
        f.construction_reserve_terms = ts.construction_reserve_terms
        f.interest_reserve = ts.interest_reserve
        if ts.borrower_signed is False:
            f.flag("review", f"Term sheet does not look signed by the borrower ({ts.signature_evidence or 'no signature seen'}). Confirm the final loan amount.",
               title="term sheet may not be signed by the borrower", section=TO, source="term_sheet")
    _apply_structure(f, x)
    if md:
        if not f.find("share_missing"):        # RULE: with the share unknown, never use the whole loan as Castellan's
            f.loan = _first(f.loan, md.senior_loan)
        f.junior_loan = _first(f.junior_loan, md.junior_loan)
        f.construction_reserve = _first(f.construction_reserve, md.construction_reserve)
        f.initial_funding = md.initial_funding
        if f.loan and md.senior_loan and abs(f.loan - md.senior_loan) > 1 and not f.is_participation:
            f.flag("check", f"Castellan loan per term sheet ${f.loan:,.0f} differs from model senior loan ${md.senior_loan:,.0f}.",
               title="term sheet loan differs from the model", section=TO, source="term_sheet")
    f.has_mezz = bool(f.junior_loan and f.junior_loan > 0)      # RULE: Mezz sentences only when a junior loan exists
    if f.has_mezz and not f.junior_lender:
        f.flag("missing", "Junior loan found but the junior lender's name is missing.",
               title="junior lender name", section=TO, source="term_sheet", key="junior_lender_missing")
    if not f.structure:
        f.structure = "note_on_note" if ui.loan_type == "note_on_note" else ("mezz" if f.has_mezz else "senior")
    senior = f.total_loan if f.is_participation and f.total_loan else f.loan
    f.total_debt = (senior or 0) + (f.junior_loan or 0) or None

    # ---- purchase price & budget (RULE: executed PSA is the final source) ----
    f.purchase_price = _first(psa and psa.purchase_price, md and md.purchase_price)
    if psa and md and psa.purchase_price and md.purchase_price:
        tol = float(rules.get("price_mismatch_tolerance_pct", 0.01))
        if abs(psa.purchase_price - md.purchase_price) / md.purchase_price > tol:
            f.flag("check", f"PSA price ${psa.purchase_price:,.0f} vs model ${md.purchase_price:,.0f} - find out why they differ.",
               title="PSA price differs from the model", section=TO, source="psa")
    if md:
        f.capex = md.capex_budget
        f.closing_costs = md.closing_costs
        f.total_cost = _first(md.total_uses, md.total_sources)
        f.facility_leverage = md.facility_leverage
        f.facility_rate = md.facility_rate_pct
    if not f.total_cost and f.purchase_price:
        f.total_cost = f.purchase_price + (f.capex or 0) + (f.closing_costs or 0)

    # ---- value (RULE: pick as-is / as-complete / as-stabilised by business plan) ----
    vals = {}
    if ap:
        vals = {"As Is": ap.as_is_value, "As Complete": ap.as_complete_value, "As Stabilized": ap.as_stabilized_value,
                "Gross Sellout": ap.gross_sellout_value}
        f.appraiser = ap.appraiser_firm
    elif md and md.appraised_value:
        vals = {(md.appraised_value_type or "As Is").title(): md.appraised_value}
    vals = {k: v for k, v in vals.items() if v}
    f.all_values = vals
    # RULE: "renovating" from the loan type, the appraisal's value types or the term sheet's purpose - not only capex
    purpose = (ts.loan_purpose if ts else "") or ""
    renovating = (ui.loan_type == "construction" or (f.capex or 0) > 0
                  or bool(ap and (ap.as_complete_value or ap.gross_sellout_value))
                  or bool(re.search(r"construct|renovat|rehab|reposition|redevelop|gut|build", purpose, re.I)))
    if ui.loan_type == "note_on_note":
        order = ["As Is", "As Complete", "As Stabilized", "Gross Sellout"]
    elif ui.loan_type == "construction" and f.property_type == "condominium":
        order = ["Gross Sellout", "As Complete", "As Stabilized", "As Is"]    # RULE: condo construction -> sellout
    elif renovating or ui.loan_type == "refinance":
        order = ["As Complete", "As Stabilized", "As Is", "Gross Sellout"]
    else:
        order = ["As Is", "As Stabilized", "As Complete", "Gross Sellout"]
    for label in order:
        if label in vals:
            f.value, f.value_label = vals[label], label
            break
    if f.value_label:
        f.value_label_lower = {"As Is": "as-is", "As Complete": "as completed", "As Stabilized": "as-stabilized",
                               "Gross Sellout": "gross sellout"}.get(f.value_label, f.value_label.lower())
        f.value_label_quoted = f"‘{f.value_label}’"
        if len(vals) > 1:
            f.flag("check", f"Appraisal has several values {', '.join(f'{k} ${v:,.0f}' for k, v in vals.items())}; "
                            f"memo uses {f.value_label}.",
               title="several appraised values - check the one used", section=R_LTV, source="appraisal")
    else:
        f.flag("missing", "No appraised value found (Appraisal folder / model).",
               title="appraised value", section=R_LTV, source="appraisal", key="value_missing")

    f.ltv = _div(f.loan, f.value)
    f.ltv_total = _div(f.total_debt, f.value)
    f.memo_ltv = _div(f.total_loan, f.value) if f.is_participation and f.total_loan else f.ltv
    lo, hi = float(rules.get("ltv_sanity_min", 0.10)), float(rules.get("ltv_sanity_max", 1.00))
    if f.memo_ltv is not None and not (lo <= f.memo_ltv <= hi):        # RULE: implausible LTV is never printed
        f.ltv_suspect = True
        f.flag("review", f"LTV of {f.memo_ltv * 100:.0f}% looks wrong (loan ${(f.total_loan if f.is_participation and f.total_loan else f.loan) or 0:,.0f} / "
                         f"{(f.value_label or 'appraised').lower()} value ${f.value or 0:,.0f}). Check the loan amount and the "
                         f"value used; the memo does not print this LTV.",
               title="LTV could not be confirmed", section=R_LTV, source="appraisal", key="ltv_suspect")
    f.ltc = _div(f.loan, f.total_cost)
    f.ltc_total = _div(f.total_debt, f.total_cost)
    f.debt_to_price = _div((f.total_debt or 0) - (f.construction_reserve or 0), f.purchase_price) if f.purchase_price else None
    f.loan_psf = _div(f.loan, f.sf)
    if ap:
        avg = ap.sales_comps_avg_adjusted_psf
        if not avg:
            ps = [c.adjusted_price_per_sf or c.price_per_sf for c in ap.sales_comps if (c.adjusted_price_per_sf or c.price_per_sf)]
            avg = sum(ps) / len(ps) if ps else None
        f.comps_psf = avg
    op = x.one_pager
    if op:   # RULE: when the one pager shows these numbers, the text uses them so it matches the tables
        if op.loan_psf:
            f.loan_psf = op.loan_psf
        if op.comps_avg_psf:
            ap_avg = f.comps_psf
            f.comps_psf = op.comps_avg_psf
            if ap_avg and abs(op.comps_avg_psf - ap_avg) / ap_avg > float(rules.get("comps_mismatch_pct", 0.15)):
                f.flag("check", f"One pager comps average ${op.comps_avg_psf:,.0f}/SF differs from the appraisal's "
                                f"${ap_avg:,.0f}/SF by more than 15%; the memo uses the one pager.",
               title="comps average differs from the appraisal", section=HL, source="one_pager")

    # ---- equity (RULE: highlight only if > 30% of total cost) ----
    f.equity = _first(md and md.sponsor_equity, (f.total_cost - f.total_debt) if f.total_cost and f.total_debt else None)
    f.equity_pct = _div(f.equity, f.total_cost)
    f.equity_highlight = bool(f.equity_pct and f.equity_pct > float(rules.get("equity_highlight_min_pct", 0.30)))

    # ---- sponsors ---------------------------------------------------------
    hon = {}
    if x.resume:
        hon = {n.lower(): h for n, h in zip(x.resume.people, x.resume.honorifics) if h}
    sponsors: list[SponsorView] = []
    cash_min = float(rules.get("pfs_cash_mention_min", 1_000_000))
    pfs_people = x.pfs.sponsors if x.pfs else []
    for sp in pfs_people:
        h = sp.honorific or hon.get(sp.full_name.lower())
        liquid = (sp.cash or 0) + (sp.marketable_securities or 0) or None
        sponsors.append(SponsorView(sp.full_name, _person_ref(sp.full_name, h, sp.last_name),
                                    sp.net_worth, liquid, sp.real_estate_assets,
                                    bool(liquid and liquid > cash_min)))     # RULE: mention cash only above $1MM
    names = [s.name for s in sponsors] or (ts.key_principals if ts else []) or \
            (x.resume.people if x.resume else []) or \
            ([p.full_name for p in x.background.people] if x.background else [])
    f.plural = len(names) > 1
    f.sponsor_names = _join(names) if names else None
    if len(sponsors) > 1:
        for s in sponsors:
            s.ref = s.name
    f.sponsors = sponsors
    f.sponsor_ref = "the sponsors" if f.plural else (sponsors[0].ref if sponsors else (names[0] if names else None))
    if not names:
        f.flag("missing", "No sponsor names found (PFS, term sheet, background report, resume).",
               title="sponsor names", section=SP_BB, source="pfs", key="sponsor_missing")
    liq = sum(s.liquid or 0 for s in sponsors)
    f.total_liquidity = liq or None
    f.liquidity_highlight = liq > float(rules.get("liquidity_highlight_min", 2_000_000))   # RULE: > $2MM
    if not sponsors:
        f.flag("missing", "Personal financial statement not found in the deal folder (PFS & SREO).",
               title="personal financial statement", section=SP_FIN, source="pfs", key="pfs_missing")
    if not (f.personal_guaranty or f.completion_guaranty or f.carveout_guaranty):
        f.flag("review", "Recourse / guaranty structure not found in the signed term sheet - confirm whether there is a "
                         "full personal guaranty or a bad-boy carve-out guaranty.",
               title="guaranty structure - confirm PG or carve-out", section=R_SP, source="term_sheet", key="guaranty_missing")

    # RULE: prior relationship is the analyst's checkbox; file-name matches in other deals are only a suggestion
    f.prior_deals = list(prior_matches or [])
    if f.prior_deals and not ui.prior_relationship:
        f.flag("check", "Other deals in the portfolio mention the sponsor (" + "; ".join(f.prior_deals[:5]) +
                        "). The memo says Castellan has NOT worked with the sponsor before - tick the box if it has.",
               title="possible earlier deal with this sponsor", section=R_SP, key="prior_deals")

    # ---- background & credit (RULE: any adverse finding -> analyst review) ----
    if x.background and x.background.people:
        # RULE: only sponsors / guarantors / key principals are named in the memo
        known = {_surname(n) for n in ([s.name for s in sponsors] + list(ts.key_principals if ts else []) +
                                       list(x.resume.people if x.resume else []))} - {""}
        issues, scores, people = [], [], []
        for p in x.background.people:
            bad = [k for k in ("bankruptcies", "judgments_or_liens", "criminal_records", "pending_civil_cases")
                   if getattr(p, k)]
            if known and _surname(p.full_name) not in known:
                note = f"{p.full_name} is in the background reports but is not a sponsor, guarantor or key principal - not named in the memo."
                if bad or p.adverse_summary:
                    note += f" Adverse items: {p.adverse_summary or ', '.join(bad)}."
                    f.flag("review", f"Background report shows adverse items for {p.full_name}, who is not a sponsor or guarantor.",
               title=f"adverse items for {p.full_name} (not a sponsor)", section=SP_BG, source="background")
                f.background_other_people.append(note)
                continue
            people.append(p)
            if bad or p.adverse_summary:
                issues.append(f"{p.full_name}: {p.adverse_summary or ', '.join(bad)}")
            if p.credit_score:
                scores.append(p.credit_score)
        f.credit_scores = [f"+{s}" for s in scores]
        strong_min = int(rules.get("strong_credit_min", 700))
        f.credit_strong = bool(scores) and all(s >= strong_min for s in scores)      # RULE: "strong" only if all >= 700
        if scores and not f.credit_strong:
            f.flag("check", f"Credit scores {', '.join(map(str, scores))} include a score below {strong_min}; "
                            f"the memo does not call them strong - review the credit sentence.",
               title=f"credit score below {strong_min} - check the wording", section=SP_BG, source="background")
        f.background_clean = not issues
        f.background_issues = "; ".join(issues) or None
        f.no_past_due = bool(people) and all(p.past_due_accounts is False for p in people)
        if issues:
            f.flag("review", "Background check shows adverse items: " + f.background_issues,
               title="adverse items in background reports", section=SP_BG, source="background", key="background_adverse")
    else:
        f.flag("missing", "Background / credit report not found.",
               title="background / credit report", section=SP_BG, source="background", key="background_missing")
    if x.adverse_search is not None:
        s_ = x.adverse_search
        f.online_search_clean = not s_.adverse_found
        f.online_search_summary = s_.overall_assessment
        f.online_findings = [fd.model_dump() for fd in s_.findings]
        # RULE: one Review flag per person with adverse findings; the findings themselves go to the workbook
        by_person: dict[str, list] = {}
        for fd in s_.findings:
            if fd.severity != "info":
                by_person.setdefault(fd.person or "the sponsor", []).append(fd)
        if s_.adverse_found and not by_person:
            by_person["the sponsor"] = []
        for person, fds in (by_person.items() if s_.adverse_found else []):
            details = "; ".join(f"{fd.category} ({fd.severity}): {fd.finding}" + (f" - {fd.source_url}" if fd.source_url else "")
                                for fd in fds) or (s_.overall_assessment or "see the Online search sheet")
            f.flag("review", f"Online search found possible adverse items for {person}: {details}",
                   title=f"online search: possible adverse items for {person}", section=SP_BG,
                   key=f"online_search:{person}")
    else:
        f.flag("check", "Online adverse-media search was not run - confirm the 'online searches' sentence manually.",
               title="online search not run - confirm the sentence", section=SP_BG, key="online_search_not_run")

    # ---- rate (RULE: GUI index choice > term sheet; fixed if either says fixed) ----
    if ts:
        f.spread, f.floor, f.fixed_rate, f.rate_cap = ts.spread_pct, ts.floor_rate_pct, ts.fixed_rate_pct, ts.rate_cap
    if md:
        f.spread = _first(f.spread, md.rate_spread_pct)
        f.floor = _first(f.floor, md.floor_rate_pct)
    f.is_floating, f.index_name = selected_index(ui, ts)
    f.index_kind = "fixed" if f.is_floating is False else rates.classify(f.index_name)
    f.index_display = rates.display(f.index_name)
    override = _index_override(ui)
    if override and ts and (ts.index_name or _ts_floating(ts) is False):
        ts_kind = "fixed" if _ts_floating(ts) is False else rates.classify(ts.index_name)
        if rates.classify(override) != ts_kind:
            used = "a fixed rate" if f.is_floating is False else f.index_name
            f.flag("check", f"Index chosen in the app ({override}) contradicts the term sheet "
                            f"({ts.index_name or 'fixed rate'}); the memo uses {used}.",
               title="index differs from the term sheet", section=R_IR, source="term_sheet")
    if op and ts:
        for what, a, b in (("floor", ts.floor_rate_pct, op.floor_rate_pct), ("spread", ts.spread_pct, op.spread_pct)):
            if a and b and abs(a - b) > 0.005:
                f.flag("check", f"Term sheet {what} {a:.2f}% vs one pager {b:.2f}% - the memo uses the term sheet.",
               title=f"{what} differs between term sheet and one pager", section=R_IR, source="one_pager")
        if ts.index_name and op.index_name and rates.classify(ts.index_name) != rates.classify(op.index_name):
            f.flag("check", f"Term sheet index '{ts.index_name}' vs one pager '{op.index_name}' - the memo uses the term sheet.",
               title="index differs between term sheet and one pager", section=R_IR, source="one_pager")
    if f.is_floating is not False:
        q = index_quote if index_quote is not None else index_quote_for(x, ui, settings, fetch=False)
        f.index_rate, f.index_rate_source = q.rate_pct, q.source
        if q.note:
            f.flag("check", q.note,
               title="index rate - check the value used", section=R_IR, source="term_sheet")
        f.effective_rate = rates.effective_rate(q.rate_pct, f.spread, f.floor)
        if f.floor and (q.rate_pct is None or not f.spread):
            why = f"no {f.index_name or 'index'} value given" if q.rate_pct is None else "no spread found"
            f.flag("check", f"Effective rate at closing set to the floor rate ({why}).",
               title="effective rate set to the floor rate", section=R_IR, source="term_sheet")

    # ---- environmental (RULE: standard language only if the report is clean) ----
    reports = list(x.environmental_reports or ([x.environmental] if x.environmental else []))
    env = x.environmental or (reports[0] if reports else None)
    phase2 = next((r for r in reports if r.phase == "phase_ii"), None)
    if env:
        f.env_provider = env.provider
        rt = env.report_type or "Phase I Environmental Site Assessment"
        f.env_report_phrase = f"{article(rt)} {rt}"
        if phase2 and phase2 is not env and phase2.status == "completed" and phase2.no_further_assessment:
            # RULE: Phase I found something, a Phase II cleared it -> dedicated standard sentence
            f.env_status = "phase2_clear"
            f.env_conditions = env.conditions_identified or env.findings_summary
            f.env_phase2_type = phase2.report_type or "Limited Phase II Environmental Site Assessment"
            if not f.env_conditions:
                f.flag("missing", "Phase I conditions not found - fill in what the Phase I identified.",
               title="Phase I conditions", section=R_ENV, source="environmental", key="env_conditions_missing")
        elif env.status != "completed":
            f.env_status = "pending"
        elif any(r.recognized_environmental_conditions for r in reports):
            f.env_status = "issues"
            bad = next(r for r in reports if r.recognized_environmental_conditions)
            f.env_summary = "; ".join(filter(None, [bad.conditions_identified, bad.findings_summary, bad.recommendations]))
            f.flag("review", "Environmental report recommends further investigation or remediation: " + (f.env_summary or ""),
               title="environmental conditions identified", section=R_ENV, source="environmental", key="env_issues")
        else:
            f.env_status = "clean"
            if env.recognized_environmental_conditions is None:
                f.flag("check", "Could not confirm 'no RECs' from the environmental report - verify.",
               title="confirm no RECs in the environmental report", section=R_ENV, source="environmental")
    else:
        f.flag("missing", "Environmental report not found.",
               title="environmental report", section=R_ENV, source="environmental", key="env_missing")

    # ---- insurance --------------------------------------------------------
    ins = x.insurance
    if ins:
        f.property_coverage, f.liability_occurrence = ins.property_coverage, ins.liability_per_occurrence
        f.umbrella, f.builders_risk, f.insurance_bound = ins.umbrella, ins.builders_risk, bool(ins.is_bound)
    else:
        f.flag("missing", "Insurance evidence not found.",
               title="insurance evidence", section=R_INS, source="insurance", key="insurance_missing")
    f.insurance_mode = "construction" if (ui.loan_type == "construction" or f.builders_risk) else "property"

    _filter_one_pager_tables(f, x)

    # ---- email chain (RULE: requests in emails never override the signed term sheet; conflicts -> check) ----
    _apply_email_insights(f, x, ui, rules)

    if ui.loan_type == "construction" and not f.pcr_firm:
        f.flag("missing", "Plan & cost review firm is not in the deal folder - fill in the Construction Completion Risk.",
               title="plan & cost review firm", section=R_CON, key="pcr_missing")
    for name, val, key, sec, src, unless in (
            ("Castellan loan amount", f.loan, "loan_missing", TO, "term_sheet", "share_missing"),
            ("property address", f.address, "address_missing", TO, "appraisal", ""),
            ("LTV", f.memo_ltv, "ltv_missing", R_LTV, "appraisal", "value_missing")):
        if not val and not (unless and f.find(unless)):
            f.flag("missing", f"{name[:1].upper() + name[1:]} could not be determined.", title=name, section=sec,
                   source=src, key=key)
    return f


def _surname(name: str) -> str:
    words = re.findall(r"[A-Za-z]+", name or "")
    words = [w for w in words if w.lower() not in ("jr", "sr", "ii", "iii", "mr", "ms", "mrs", "dr")]
    return words[-1].lower() if words else ""


def _apply_structure(f: DealFacts, x: Extracted) -> None:
    """RULE (participations): a term sheet issued by another lender, or a pari passu / participation structure in the
    term sheet or one pager, means Castellan holds a share: f.loan = Castellan's share, f.total_loan = whole loan."""
    ts, op = x.term_sheet, x.one_pager
    f.structure = _first(ts and ts.structure, op and op.structure)
    issuer = ts.issuing_lender if ts else None
    lead = _first(ts and ts.lead_lender, op and op.lead_lender)
    if issuer and "castellan" not in issuer.lower():
        lead = lead or issuer
        if f.structure not in ("pari_passu", "participation"):
            f.structure = "participation"
    f.is_participation = f.structure in ("pari_passu", "participation")
    if not f.is_participation:
        return
    total = _first(ts and ts.total_loan_amount, op and op.total_loan_amount,
                   ts and (ts.senior_loan_amount or ts.blended_loan_amount))
    share = _first(ts and ts.castellan_share_amount, op and op.castellan_share_amount)
    f.lead_lender, f.total_loan = lead, total
    if share:
        f.loan = share
    else:
        f.loan = None          # RULE: never present the whole loan as Castellan's position
        f.flag("missing", "Castellan's share of the participated loan was not found - the loan amount may be the whole loan.",
               title="Castellan's share", section=TO, source="term_sheet", key="share_missing")
    f.castellan_share_pct = _first(ts and ts.castellan_share_pct, op and op.castellan_share_pct,
                                   (share / total * 100) if share and total else None)
    if not lead:
        f.flag("missing", "Lead lender of the participated loan was not found.",
               title="lead lender", section=TO, source="term_sheet", key="lead_missing")
    if not total:
        f.flag("missing", "Total amount of the participated loan was not found.",
               title="total participated loan amount", section=TO, source="term_sheet", key="total_loan_missing")


def _filter_one_pager_tables(f: DealFacts, x: Extracted) -> None:
    """RULE (one pager): drop tables that describe another property (a different street address, not comps)."""
    from .sources_plan import other_addresses, subject_street
    op = x.one_pager
    if not op or not op.tables or not f.address:
        return
    number, core = subject_street(f.address.split(",")[0])
    if not number:
        return
    kept = []
    for t in op.tables:
        text = " ".join([t.title] + list(t.columns) + [c for r in t.rows for c in r])
        others = [] if re.search(r"comp", t.title, re.I) else other_addresses(text, number, core)
        if others and not re.search(rf"\b{re.escape(number)}\b", text):
            f.flag("review", f"One pager table '{t.title}' mentions {others[0]}, not the subject - it was left out of the memo.",
               title=f"one pager table '{t.title}' is about another property", section=LF, source="one_pager")
            continue
        kept.append(t)
    if len(kept) != len(op.tables):
        x.one_pager = op.model_copy(update={"tables": kept})


_RATIO = re.compile(r"(\d+(?:\.\d+)?)\s*%?\s*(LTV|LTC|loan[- ]to[- ](?:value|cost))", re.I)


def requested_ratios(text: Optional[str]) -> dict[str, float]:
    """'65% LTV / 80% LTC' -> {'LTV': 65.0, 'LTC': 80.0} (percent numbers)."""
    out: dict[str, float] = {}
    for num, kind in _RATIO.findall(text or ""):
        k = "LTV" if kind.upper() == "LTV" or "value" in kind.lower() else "LTC"
        out.setdefault(k, float(num))
    return out


def _apply_email_insights(f: DealFacts, x: Extracted, ui: UserInputs, rules: dict) -> None:
    ei = x.email_insights
    if not ei:
        return
    d = ei.model_dump()
    if ei.requested_loan_amount:
        d["requested_loan_amount"] = f"${ei.requested_loan_amount:,.0f}"
    f.email_insights = {k: v for k, v in d.items() if v not in (None, "", [])}
    if not f.exit_strategy and ei.exit_strategy:          # GUI "(infer from notes)"
        f.exit_strategy, f.exit_strategy_source = ei.exit_strategy, "email chain"

    tol = float(rules.get("email_amount_tolerance_pct", 0.01))
    ts = x.term_sheet
    req = ei.requested_loan_amount
    if req and ts and f.loan:
        candidates = [v for v in (f.loan, f.blended_loan) if v]
        if all(abs(req - v) / v > tol for v in candidates):
            blended = f" (blended ${f.blended_loan:,.0f})" if f.blended_loan and f.blended_loan != f.loan else ""
            f.flag("check", f"Email chain requested loan ${req:,.0f} differs from the term sheet loan ${f.loan:,.0f}"
                            f"{blended}. The memo uses the term sheet - confirm the final amount.",
               title="requested loan differs from the term sheet", section=TO, source="emails")
    pts = float(rules.get("email_ratio_tolerance_pts", 1.0))
    for kind, wanted in requested_ratios(ei.requested_ltv_or_ltc).items():
        actual = f.ltv if kind == "LTV" else f.ltc
        if actual is not None and abs(wanted - actual * 100) > pts:
            f.flag("check", f"Email chain requested {wanted:.0f}% {kind}; the computed {kind} is {actual * 100:.0f}%.",
               title=f"requested {kind} differs from the computed {kind}", section=TO, source="emails")


def _index_override(ui: UserInputs) -> str:
    o = (ui.index_name_override or "").strip()
    return "" if o == rates.FROM_TERM_SHEET else o


def _ts_floating(ts: S.TermSheet) -> Optional[bool]:
    if ts.is_floating is not None:
        return ts.is_floating
    return False if ts.fixed_rate_pct and not ts.spread_pct else None


def selected_index(ui: UserInputs, ts: Optional[S.TermSheet]) -> tuple[Optional[bool], Optional[str]]:
    """(is_floating, index name). A fixed rate chosen in the app or stated in the term sheet means no index."""
    override = _index_override(ui)
    floating = _ts_floating(ts) if ts else None
    if rates.classify(override) == "fixed" or floating is False:
        return False, None
    if override:
        return True, override
    name = (ts.index_name if ts else None) or ("1-Month CME SOFR" if floating else None)
    return floating, name


def index_quote_for(x: Extracted, ui: UserInputs, settings, fetch: bool = True, log=None) -> rates.IndexQuote:
    """Index value for the effective rate. `fetch=True` may call the NY Fed / FRED (never raises)."""
    floating, name = selected_index(ui, x.term_sheet)
    kind = "fixed" if floating is False else rates.classify(name)
    model_rate = x.model.index_rate_pct if x.model else None
    return rates.resolve(kind, ui.index_rate_pct, settings, model_rate=model_rate, fetch=fetch, log=log)


def memo_date(d: date | None = None) -> str:
    """'October 7, 2026' on every OS (strftime's %-d is not available on Windows)."""
    d = d or date.today()
    return f"{d.strftime('%B')} {d.day}, {d.year}"


# --------------------------------------------------------------------------- loan financial tables
def _m(v):
    return f"${v:,.0f}" if v else ""


def _p(v, dp=0):
    return f"{v * 100:.{dp}f}%" if v else ""


def build_financial_tables(f: DealFacts, x: Extracted) -> list[S.Table]:
    """Use the one pager when it exists (the team's official tables); otherwise rebuild them from the facts."""
    if x.one_pager and x.one_pager.tables:
        return x.one_pager.tables
    tables = []
    rows = [["Loan Amount:", _m(f.loan)],
            [f"Appraised Value ({(f.value_label or '').lower()}):", _m(f.value)],
            ["Estimated Loan to Value:", _p(f.ltv)], ["", ""],
            ["Building Area (SF):", f"{f.sf:,.0f}" if f.sf else ""],
            ["Loan Per Square Foot:", _m(f.loan_psf)]]
    bold = [2, 5]
    if f.facility_leverage:
        annual = f.facility_leverage * (f.facility_rate or 0) / 100
        rows += [["", ""], ["Est. Leverage from Facility:", _m(f.facility_leverage)],
                 ["Facility Interest Rate:", f"{f.facility_rate:.2f}%" if f.facility_rate else ""],
                 ["Est. Annual Facility Interest:", _m(annual)],
                 ["Facility Estimated LTV:", _p(_div(f.facility_leverage, f.value))]]
        bold.append(len(rows) - 1)
    tables.append(S.Table(sheet=None, title="Loan Metrics", columns=[], rows=rows, bold_rows=bold))

    rate = (f"{f.index_name} + {f.spread:.2f}%" if f.is_floating is not False and f.spread
            else (f"{f.fixed_rate:.2f}% fixed" if f.fixed_rate else ""))
    tables.append(S.Table(sheet=None, title="Loan Terms", columns=[], rows=[
        ["Loan Amount", _m(f.loan)], ["Rate to Borrower", rate],
        ["Floor Rate", f"{f.floor:.2f}%" if f.floor else ""],
        ["Loan Term", f"{f.loan_term_months} Months" if f.loan_term_months else ""],
        ["Estimated Closing Date", f.closing_date or ""], ["Security", f.security_summary or ""],
        [f"Appraised Value ({(f.value_label or '').lower()})", _m(f.value)]], bold_rows=[6]))

    md = x.model
    if md and (md.sources or md.uses):
        n = max(len(md.sources), len(md.uses))
        su_rows = []
        for i in range(n):
            s = md.sources[i] if i < len(md.sources) else None
            u = md.uses[i] if i < len(md.uses) else None
            su_rows.append([s.label if s else "", _m(s.amount) if s else "",
                            (f"{s.pct:.0f}%" if s and s.pct else (_p(_div(s.amount, md.total_sources)) if s else "")),
                            u.label if u else "", _m(u.amount) if u else ""])
        su_rows.append(["Total Sources", _m(md.total_sources), "100%", "Total Uses", _m(md.total_uses)])
        tables.append(S.Table(sheet=None, title="Estimated Sources & Uses:", columns=["Sources", "$", "%", "Uses", "$"],
                              rows=su_rows, bold_rows=[len(su_rows) - 1]))

    ap = x.appraisal
    if ap and ap.sales_comps:
        crow = [[c.address, c.sale_date or "", _m(c.sale_price), f"{c.size_sf:,.0f}" if c.size_sf else "",
                 _m(c.adjusted_price_per_sf or c.price_per_sf)] for c in ap.sales_comps]
        crow.append(["Average", "", "", "", _m(f.comps_psf)])
        crow.append([f.address or "Subject Property", "", _m(f.value), f"{f.sf:,.0f}" if f.sf else "",
                     _m(_div(f.value, f.sf))])
        tables.append(S.Table(sheet=None, title="Sales Comparables", columns=["Address", "Sale Date", "Sale Price", "SF", "$/SF"],
                              rows=crow, bold_rows=[len(crow) - 2]))
    return tables
