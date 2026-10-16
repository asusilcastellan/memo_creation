"""Structured-output schemas. Each source document is read into one of these.

Rules for OpenAI structured outputs: every field is required but may be null, and
no field has a default value. Money is in US dollars as a plain number; rates and
percentages are percent numbers (9.64 means 9.64%).
"""
from __future__ import annotations

from typing import Literal, Optional

from pydantic import BaseModel, Field

Money = Optional[float]
Pct = Optional[float]
PropertyType = Literal["multifamily", "industrial", "retail", "office", "single-family", "condominium", "mixed-use",
                       "land", "hospitality", "self-storage", "other"]
Structure = Literal["senior", "pari_passu", "participation", "mezz", "note_on_note"]
PROPERTY_TYPE_HELP = ("One of the listed types. Condo units / condo development = condominium; apartments = multifamily; "
                      "warehouse / flex / logistics = industrial; hotel = hospitality.")
QUALIFIER_HELP = ("One or two words that describe the project as the document does, e.g. 'luxury', 'Class A', "
                  "'garden-style', 'ground-up'. Null if none. Never the property type itself.")


# --------------------------------------------------------------------------- term sheet
class TermSheet(BaseModel):
    property_address: Optional[str] = Field(description="Full property address as written in the term sheet.")
    borrower_signed: Optional[bool] = Field(description="True if the borrower/guarantor signature block is actually signed (handwritten, e-signature or /s/). Null if unclear.")
    signature_evidence: Optional[str] = Field(description="Short description of what shows it is or is not signed.")
    blended_loan_amount: Money = Field(description="'Blended Loan Amount' / total loan amount including any junior participation.")
    senior_loan_amount: Money = Field(description="Castellan's own senior loan amount if stated separately; else null.")
    junior_lender_name: Optional[str] = Field(description="Name of the junior / mezzanine lender or junior loan participant, if any.")
    junior_loan_amount: Money = Field(description="Junior / mezz loan amount, if any.")
    loan_purpose: Optional[str] = Field(description="Purpose of the loan as described (acquisition, renovation, construction, refinance...).")
    property_description: Optional[str] = Field(description="Property description line (type, size).")
    issuing_lender: Optional[str] = Field(description="The lender that issued this term sheet (letterhead / 'Lender:'), e.g. 'Castellan Capital' or 'Maxim Capital'.")
    structure: Optional[Structure] = Field(description="senior = Castellan alone; pari_passu / participation = Castellan holds a share of a larger senior loan led by another lender; mezz = there is a junior/mezzanine loan; note_on_note = loan secured by another loan. Null if unclear.")
    lead_lender: Optional[str] = Field(description="Lead / agent lender of a participated or pari passu loan, if not Castellan.")
    total_loan_amount: Money = Field(description="Total amount of the (participated) senior loan, if Castellan only holds a share.")
    castellan_share_amount: Money = Field(description="Castellan's own share of a participated / pari passu loan.")
    castellan_share_pct: Pct = Field(description="Castellan's share as a percent of the total loan, if stated.")
    is_floating: Optional[bool]
    index_name: Optional[str] = Field(description="e.g. '1-Month CME Term SOFR'.")
    spread_pct: Pct
    floor_rate_pct: Pct
    fixed_rate_pct: Pct
    rate_cap: Optional[str] = Field(description="Rate cap terms if any; null if not capped.")
    loan_term_months: Optional[int]
    extension_options: Optional[str]
    origination_fee: Optional[str]
    interest_reserve: Money
    construction_reserve: Money
    construction_reserve_terms: Optional[str] = Field(description="When / how the construction reserve is released.")
    max_ltv_pct: Pct
    personal_guaranty: Optional[bool] = Field(description="Full personal guaranty of payment (recourse) by key principals.")
    completion_guaranty: Optional[bool]
    carveout_guaranty: Optional[bool] = Field(description="Non-recourse carve-out ('bad boy') guaranty.")
    recourse_text: Optional[str] = Field(description="Short paraphrase of the Recourse / Guaranty clause.")
    key_principals: list[str] = Field(description="Names of key principals / guarantors.")
    security_summary: Optional[str]
    estimated_closing_date: Optional[str]


# --------------------------------------------------------------------------- PSA
class PSA(BaseModel):
    purchase_price: Money
    deposit: Money
    seller: Optional[str]
    buyer: Optional[str]
    agreement_date: Optional[str]
    closing_terms: Optional[str]
    property_address: Optional[str]
    is_executed: Optional[bool] = Field(description="True if signed by both parties.")


# --------------------------------------------------------------------------- underwriting model / one pager
class LineItem(BaseModel):
    label: str
    amount: Money
    pct: Pct


class UWModel(BaseModel):
    property_address: Optional[str]
    property_type: Optional[PropertyType] = Field(description=PROPERTY_TYPE_HELP)
    property_features: list[str] = Field(description="Notable features from the Loan Overview block (loading docks, ceiling height, amenities...).")
    year_built: Optional[str]
    units: Optional[int]
    unit_mix: Optional[str]
    gross_sf: Optional[float]
    net_rentable_sf: Optional[float]
    land_acres: Optional[float]
    purchase_price: Money
    capex_budget: Money
    closing_costs: Money
    senior_loan: Money
    junior_loan: Money
    sponsor_equity: Money
    total_sources: Money
    total_uses: Money
    sources: list[LineItem]
    uses: list[LineItem]
    appraised_value: Money
    appraised_value_type: Optional[str]
    initial_funding: Money
    construction_reserve: Money
    facility_leverage: Money
    facility_rate_pct: Pct
    rate_spread_pct: Pct
    floor_rate_pct: Pct
    index_rate_pct: Pct = Field(description="Index (e.g. SOFR) value used in the model, if shown.")
    notes: Optional[str]


class Table(BaseModel):
    sheet: Optional[str] = Field(description="Name of the workbook sheet the table is on (the '### Sheet:' line); null if not from a workbook.")
    title: str
    columns: list[str] = Field(description="Column headers; empty list for 2-column label/value tables.")
    rows: list[list[str]] = Field(description="Rows of already-formatted cell text, e.g. ['Loan Amount:', '$2,240,000'].")
    bold_rows: list[int] = Field(description="0-based indexes of rows shown in bold (totals, LTV lines).")


class OnePager(BaseModel):
    tables: list[Table] = Field(description="Every table on the one pager in reading order: Loan Metrics, Loan Terms, Estimated Sources & Uses, Sales Comparables, etc. Each logical table ONCE.")
    property_address: Optional[str] = Field(description="Subject property address shown on the one pager.")
    structure: Optional[Structure] = Field(description="pari_passu if it shows e.g. 'Castellan pari passu loan'; participation for a participation; else senior / mezz.")
    lead_lender: Optional[str]
    total_loan_amount: Money = Field(description="Total senior loan / 'Total Debt' when Castellan only holds a share.")
    castellan_share_amount: Money = Field(description="Castellan's loan / pari passu share.")
    castellan_share_pct: Pct
    loan_psf: Money = Field(description="Loan per square foot as shown.")
    comps_avg_psf: Money = Field(description="Average (adjusted) $/SF of the sales comparables as shown.")
    index_name: Optional[str]
    spread_pct: Pct
    floor_rate_pct: Pct


# --------------------------------------------------------------------------- sponsor documents
class SponsorPFS(BaseModel):
    full_name: str
    honorific: Optional[Literal["Mr.", "Ms.", "Mrs.", "Dr."]]
    last_name: Optional[str]
    as_of_date: Optional[str]
    net_worth: Money
    cash: Money
    marketable_securities: Money
    real_estate_assets: Money
    total_liabilities: Money


class PFSResult(BaseModel):
    sponsors: list[SponsorPFS]


class PersonBackground(BaseModel):
    full_name: str
    credit_score: Optional[int]
    bankruptcies: Optional[bool]
    judgments_or_liens: Optional[bool]
    criminal_records: Optional[bool]
    pending_civil_cases: Optional[bool]
    past_due_accounts: Optional[bool]
    adverse_summary: Optional[str] = Field(description="Plain-English summary of anything adverse (fraud, foreclosure, litigation, liens, delinquency). Null if clean.")


class BackgroundResult(BaseModel):
    provider: Optional[str]
    people: list[PersonBackground]


class ResumeFacts(BaseModel):
    people: list[str]
    honorifics: list[str] = Field(description="'Mr.'/'Ms.' for each person in the same order, '' if unknown.")
    company: Optional[str]
    general_contractor: Optional[str]


# --------------------------------------------------------------------------- appraisal
class SaleComp(BaseModel):
    address: str
    sale_date: Optional[str]
    sale_price: Money
    size_sf: Optional[float]
    price_per_sf: Money
    adjusted_price_per_sf: Money


class Appraisal(BaseModel):
    appraiser_firm: Optional[str]
    effective_date: Optional[str]
    property_address: Optional[str]
    property_type: Optional[PropertyType] = Field(description=PROPERTY_TYPE_HELP)
    project_qualifier: Optional[str] = Field(description=QUALIFIER_HELP)
    property_description: Optional[str] = Field(description="One or two sentences: building/land, size, condition, configuration.")
    gross_building_sf: Optional[float]
    net_rentable_sf: Optional[float]
    units: Optional[int]
    site_acres: Optional[float]
    year_built: Optional[str]
    as_is_value: Money
    as_complete_value: Money
    as_stabilized_value: Money
    gross_sellout_value: Money = Field(description="Condo 'gross retail sellout' / 'aggregate retail value' of all units, if concluded.")
    sales_comps: list[SaleComp]
    sales_comps_avg_adjusted_psf: Money
    location: Optional[str] = Field(description="Where the property is: corner/street, neighbourhood, city, county, state; character of the immediate area.")
    access_and_demand_drivers: Optional[str] = Field(description="Highways, transit, employers, hospitals, universities, amenities.")
    market_statistics: list[str] = Field(description="Each submarket statistic as a sentence with its number, period and source (vacancy, availability, rents by type, under construction, absorption, rent growth, sales pricing, occupancy/ADR).")
    demographics: list[str] = Field(description="Population, growth, median household income / AMI with year and geography.")
    highest_best_use_or_plan: Optional[str]


# --------------------------------------------------------------------------- third-party reports
class Environmental(BaseModel):
    provider: Optional[str]
    report_type: Optional[str] = Field(description="e.g. 'Phase I Environmental Site Assessment', 'Limited Phase II Environmental Site Assessment', 'Historical Records + Database Review (HRDR)'.")
    phase: Optional[Literal["phase_i", "phase_ii", "database_review", "other"]]
    report_date: Optional[str]
    status: Literal["completed", "draft", "pending"]
    recognized_environmental_conditions: Optional[bool] = Field(description="True ONLY when the report recommends further investigation or remediation. Conditions noted without such a recommendation -> false.")
    conditions_identified: Optional[str] = Field(description="Short noun phrase for the conditions found, e.g. 'a REC associated with a former dry cleaner'. Null if none.")
    no_further_assessment: Optional[bool] = Field(description="True if the report concludes that no further assessment / action is warranted.")
    findings_summary: Optional[str]
    recommendations: Optional[str]


class Insurance(BaseModel):
    insurer_or_broker: Optional[str]
    document_type: Optional[str] = Field(description="evidence of property insurance, certificate of liability, binder, quote...")
    is_bound: Optional[bool]
    property_coverage: Money
    liability_per_occurrence: Money
    liability_aggregate: Money
    umbrella: Money
    builders_risk: Money


# --------------------------------------------------------------------------- email chain
class EmailInsights(BaseModel):
    loan_purpose: Optional[str] = Field(description="Why the borrower needs the loan, as described in the emails.")
    business_plan: Optional[str] = Field(description="What the borrower will do with the property (renovation, lease-up, construction...), with numbers and timing as stated.")
    exit_strategy: Optional[str] = Field(description="How the loan will be repaid (refinance, sale, sellout...). Null if not discussed.")
    requested_loan_amount: Money = Field(description="Loan amount requested by the broker/borrower (a request, not final terms).")
    requested_ltv_or_ltc: Optional[str] = Field(description="Requested leverage as written, e.g. '65% LTV' or '75% LTC'. Null if not stated.")
    requested_terms: Optional[str] = Field(description="Other requested terms (rate, term, IO, reserves, fees) as written.")
    property_facts: list[str] = Field(description="Facts about the property (size, units, condition, occupancy, rents, zoning).")
    sponsor_facts: list[str] = Field(description="Facts about the sponsor (experience, track record, other properties, liquidity).")
    timeline: list[str] = Field(description="Dated events: 'YYYY-MM-DD - what happened or is planned'.")
    open_issues: list[str] = Field(description="Open questions, missing items, conditions or concerns raised in the thread.")
    key_quotes: list[str] = Field(description="Verbatim excerpts of at most 25 words, each ending with '- <sender>, <date>'.")
    participants: list[str] = Field(description="'Name (role/company, email)' for each person in the thread.")


# --------------------------------------------------------------------------- writer: plan step
class PlannedSentence(BaseModel):
    purpose: str = Field(description="What this sentence does, e.g. 'state the loan amount and purpose'.")
    facts_used: list[str] = Field(description="Each fact this sentence uses, copied exactly as written in the DEAL FACTS or source excerpts (numbers unchanged).")


class SectionPlan(BaseModel):
    sentences: list[PlannedSentence] = Field(description="The sentences of the section, in order, following the playbook's required elements.")


class AdverseFinding(BaseModel):
    person: str = Field(description="The sponsor / guarantor the finding is about, as named in the request.")
    finding: str = Field(description="One sentence: what the source says (one item per finding).")
    category: Literal["litigation", "criminal", "bankruptcy", "foreclosure", "lien", "regulatory", "press", "other"]
    severity: Literal["high", "medium", "low", "info"] = Field(description="info = neutral / not adverse (e.g. a news mention).")
    source_url: Optional[str] = Field(description="URL of the page the finding comes from.")
    source_title: Optional[str] = Field(description="Title of that page / article.")


class AdverseSearch(BaseModel):
    findings: list[AdverseFinding] = Field(description="One entry per item found, each with its URL. Empty if nothing relevant.")
    overall_assessment: Optional[str] = Field(description="One or two sentences summarising the search.")
    adverse_found: bool = Field(description="True only if at least one credible adverse finding clearly matches the person.")
