"""Write the free-text sections in Castellan's voice: plan -> write -> check, per section.

System prompt = style_guide.md + resources/style/playbooks/<section>.yaml (purpose, required elements, guidance
for this loan type / structure, allowed attributions, banned phrases, target length, notes).
User prompt   = top-k tagged past examples (memo_app/examples.py) + the deal facts (already computed and
                formatted - no arithmetic left to the model) + source excerpts.

  1. plan   LLM.extract(SectionPlan): one entry per sentence with the facts it uses. Code checks every number
            in those facts against the facts/sources; sentences with unknown numbers are dropped (review note).
  2. write  the writer model writes the paragraph from the approved plan.
  3. check  code only (memo_app/language.py): banned phrases, length, unknown numbers, markdown, "we".
            On failure ONE repair call listing the exact problems; if it still fails the text is kept with a
            short red remark linked to a review item (the review workbook lists the problems).
Every prompt and output is written to the run's debug.log, never to the screen.

Sources are described to the model as background facts it must not cite (it would otherwise echo labels such
as "according to the email chain"). The analyst's Notes box is the highest-priority narrative source.
"""
from __future__ import annotations

import json
import re
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Callable

import yaml

from . import paths
from . import schemas as S
from .examples import deal_tags, load_examples, select
from .extract import Extracted
from .facts import DealFacts
from .language import allowed_numbers, check_text, verify_plan
from .llm import LLM
from .render import ai, render, short_remark, strip_markers

Log = Callable[[str], None]

SECTIONS = ["transaction_overview", "business_background", "submarket_overview", "business_plan", "sponsor_experience"]
LOAN_TYPE_NAMES = {"acquisition": "Acquisition", "construction": "Construction",
                   "refinance": "Refinance", "note_on_note": "Note-on-Note"}
SECTION_TITLES = {"transaction_overview": "Transaction Overview", "business_background": "Business background",
                  "submarket_overview": "Submarket Overview", "business_plan": "Business Plan",
                  "sponsor_experience": "Sponsor Risk (experience sentence)"}
# memo section of each AI section (review workbook "Memo section")
MEMO_SECTION = {"transaction_overview": "Transaction Overview", "business_background": "Sponsorship - Business background",
                "submarket_overview": "Submarket Overview", "business_plan": "Business Plan",
                "sponsor_experience": "Risks - Sponsor"}
SECTION_SOURCE = {"transaction_overview": "term_sheet", "business_background": "resume", "submarket_overview": "appraisal",
                  "business_plan": "model", "sponsor_experience": "resume"}

MONEY_KEYS = {"loan", "blended_loan", "junior_loan", "total_debt", "purchase_price", "capex", "closing_costs",
              "total_cost", "initial_funding", "construction_reserve", "value", "loan_psf", "comps_psf",
              "total_liquidity", "equity", "interest_reserve", "facility_leverage", "property_coverage",
              "liability_occurrence", "umbrella", "builders_risk", "total_loan"}
RATIO_KEYS = {"ltv", "ltc", "ltc_total", "debt_to_price", "equity_pct", "ltv_total", "memo_ltv"}
RATE_KEYS = {"spread", "floor", "fixed_rate", "effective_rate", "facility_rate", "index_rate", "castellan_share_pct"}

PLAN_TASK = (
    "\n\nTASK NOW: plan the section before writing it. Return one entry per sentence, in order, following the "
    "required elements. For each sentence copy the facts it will use EXACTLY as they appear in the DEAL FACTS or "
    "source excerpts (numbers unchanged). Use no fact that is not supplied.")


def _fmt_facts(f: DealFacts) -> dict:
    out = {}
    for k, v in f.for_prompt().items():
        if k in MONEY_KEYS and isinstance(v, (int, float)):
            out[k] = f"${v:,.0f}"
        elif k in RATIO_KEYS and isinstance(v, (int, float)):
            out[k] = f"{v * 100:.0f}%"
        elif k in RATE_KEYS and isinstance(v, (int, float)):
            out[k] = f"{v:.2f}%"
        elif k == "all_values":
            out[k] = {kk: f"${vv:,.0f}" for kk, vv in v.items()}
        else:
            out[k] = v
    out["loan_type"] = LOAN_TYPE_NAMES.get(f.loan_type, f.loan_type)
    return out


class StylePack:
    def __init__(self, root: Path | None = None):
        self.root = root or paths.resource("style")
        self.guide = (self.root / "style_guide.md").read_text(encoding="utf-8")
        self._examples: dict[str, list] = {}

    def playbook(self, section: str) -> dict:
        p = self.root / "playbooks" / f"{section}.yaml"
        return (yaml.safe_load(p.read_text(encoding="utf-8")) or {}) if p.exists() else {}

    def variant_keys(self, f: DealFacts, section: str) -> list[str]:
        tags = deal_tags(f)
        keys = [f.loan_type, f.structure or "", f.property_type or "", "default"]
        if tags.get("exit"):
            keys.append(f"{tags['exit']}_exit")
        if section == "business_background":
            keys.append("several_sponsors" if f.plural else "single_sponsor")
        return [k for k in keys if k]

    def system(self, section: str, f: DealFacts) -> str:
        pb = self.playbook(section)
        lines = [f"SECTION: {pb.get('section', SECTION_TITLES.get(section, section))}",
                 f"Purpose: {pb.get('purpose', '').strip()}"]
        req = pb.get("required_elements") or []
        if req:
            lines.append("Required elements, in this order:")
            lines += [f"{i}. {r}" for i, r in enumerate(req, 1)]
        variants = pb.get("variants") or {}
        guidance = [f"- {variants[k].strip()}" for k in self.variant_keys(f, section) if k in variants]
        if guidance:
            lines.append("Guidance for this deal:")
            lines += guidance
        if pb.get("allowed_attributions"):
            lines.append("The ONLY allowed ways to cite a source: " + "; ".join(f'"{a}"' for a in pb["allowed_attributions"])
                         + ". Never mention emails, correspondence, the term sheet, the underwriting model or the notes.")
        if pb.get("banned_phrases"):
            lines.append("Never use these words or phrases: " + ", ".join(f'"{b}"' for b in pb["banned_phrases"]) + ".")
        tw = pb.get("target_words") or {}
        if tw:
            lines.append(f"Length: {tw.get('min')}-{tw.get('max')} words.")
        if pb.get("notes"):
            lines.append(f"Notes: {str(pb['notes']).strip()}")
        return self.guide + "\n\n" + "\n".join(lines)

    def examples(self, section: str, f: DealFacts, k: int) -> list[tuple[str, str]]:
        if section not in self._examples:
            self._examples[section] = load_examples(self.root / "examples", section)
        return [(ex.label, ex.text) for ex in select(self._examples[section], deal_tags(f), k)]


class Writer:
    def __init__(self, llm: LLM, settings, boilerplate: dict, log: Log = print):
        self.llm = llm
        self.s = settings
        self.bp = boilerplate
        self.log = log
        self.pack = StylePack()
        self.k = int(settings.get("memo.few_shot_examples_per_section", 4))

    # ------------------------------------------------------------------ prompts
    def _sources_for(self, section: str, f: DealFacts, x: Extracted) -> str:
        parts = []
        narrative = section in ("transaction_overview", "business_plan")
        if f.notes.strip() and narrative:
            parts.append("ANALYST NOTES (highest priority - they override every other narrative fact; background "
                         "facts, do not cite or mention them):\n" + f.notes.strip())
        if f.email_insights and narrative:
            parts.append("BACKGROUND FACTS ON PURPOSE, PLAN AND EXIT (do not cite or mention where they come from; "
                         "requested amounts are not the loan terms - numbers come from DEAL FACTS):\n" +
                         json.dumps(f.email_insights, indent=1, ensure_ascii=False))
        if f.exit_strategy and section in ("transaction_overview", "business_plan", "submarket_overview"):
            label = "chosen by the analyst" if f.exit_strategy_source == "analyst" else "background fact"
            parts.append(f"EXIT STRATEGY ({label}): {f.exit_strategy}")
        ts = x.term_sheet
        if ts and narrative:
            parts.append("LOAN TERMS (background facts - do not cite):\n" + json.dumps({
                "loan_purpose": ts.loan_purpose, "property_description": ts.property_description,
                "construction_reserve_terms": ts.construction_reserve_terms,
                "junior_loan_participant": ts.junior_lender_name}, indent=1))
        if x.model and narrative:
            parts.append("PROJECT BUDGET AND PROPERTY FACTS (background facts - do not cite):\n" +
                         x.model.model_dump_json(indent=1, exclude_none=True))
        if x.appraisal and narrative:
            parts.append("APPRAISAL - PROPERTY AND PLAN (may be cited as the appraisal report):\n" + json.dumps({
                "property_description": x.appraisal.property_description,
                "highest_best_use_or_plan": x.appraisal.highest_best_use_or_plan}, indent=1))
        if section == "submarket_overview":
            if x.appraisal:
                a = x.appraisal
                parts.append("APPRAISAL - MARKET FACTS (may be cited as the appraisal report / its data source):\n" + json.dumps({
                    "location": a.location, "access_and_demand_drivers": a.access_and_demand_drivers,
                    "market_statistics": a.market_statistics, "demographics": a.demographics,
                    "effective_date": a.effective_date}, indent=1))
            if x.appraisal_text:
                parts.append("APPRAISAL - RELEVANT PAGES (for wording and extra detail):\n" + x.appraisal_text[:60000])
        if section in ("business_background", "sponsor_experience"):
            if x.resume_text:
                parts.append("SPONSOR RESUME / BIO (primary source; may be cited as the resume provided to Castellan):\n"
                             + x.resume_text[:60000])
            if x.resume and x.resume.general_contractor and f.loan_type == "construction":
                parts.append(f"GENERAL CONTRACTOR: {x.resume.general_contractor}")
        return "\n\n".join(parts) if parts else "(no additional source excerpts)"

    def _user_prompt(self, section: str, f: DealFacts, x: Extracted, opening: str | None) -> str:
        exs = self.pack.examples(section, f, self.k)
        ex_block = "\n\n".join(f"<example source=\"{lbl}\">\n{txt}\n</example>" for lbl, txt in exs) or "(none yet)"
        req = f"\nREQUIRED OPENING - begin the paragraph with exactly these words and continue the sentence:\n{opening}\n" if opening else ""
        return (
            f"Write the {section.replace('_', ' ').upper()} for a {LOAN_TYPE_NAMES.get(f.loan_type)} loan.\n\n"
            f"Past Castellan memos - match their wording, sentence structure and length:\n{ex_block}\n\n"
            f"DEAL FACTS (already verified and computed - use these numbers exactly):\n"
            f"{json.dumps(_fmt_facts(f), indent=1, default=str)}\n\n"
            f"SOURCE EXCERPTS:\n{self._sources_for(section, f, x)}\n{req}"
        )

    # ------------------------------------------------------------------ helpers
    @staticmethod
    def _clean(text: str) -> str:
        text = text.strip().strip('"').strip()
        text = re.sub(r"\*\*|__|^#+\s*", "", text, flags=re.M)
        text = re.sub(r"^(Business background|Transaction Overview|Business Plan|Submarket Overview)\s*:\s*", "", text, flags=re.I)
        return re.sub(r"[ \t]+", " ", text).strip()

    @staticmethod
    def _remarks(text: str, f: DealFacts, section: str) -> str:
        """[MISSING: ...] / [CHECK: ...] written by the model become red remarks linked to a review item."""
        def sub(m):
            level, _, what = m.group(0)[1:-1].partition(":")
            what = what.strip() or level
            return f.new_remark(level.strip().lower(), short_remark(what),
                                f"The AI draft of the {SECTION_TITLES[section]} marked this as {level.strip().lower()}: {what}",
                                MEMO_SECTION[section], SECTION_SOURCE[section])
        return re.sub(r"\[(?:MISSING|CHECK|REVIEW)\b[^\]]*\]", sub, text)

    def _debug(self, section: str, step: str, text: str) -> None:
        self.log(f"===== {section} | {step} =====\n{text}")

    # ------------------------------------------------------------------ plan -> write -> check
    def _plan(self, section: str, system: str, user: str, allowed: set[str], f: DealFacts) -> str:
        try:
            plan = self.llm.extract(S.SectionPlan, system + PLAN_TASK, user)
        except Exception as exc:
            self._debug(section, "plan failed", repr(exc))
            return ""
        self._debug(section, "plan", plan.model_dump_json(indent=1))
        approved, dropped = verify_plan(plan, allowed)
        for s, nums in dropped:
            f.flag("review", f"{SECTION_TITLES[section]}: a planned sentence ({s.purpose}) used numbers that are not in "
                             f"the sources ({', '.join(nums)}) and was left out.",
                   title="a planned sentence used numbers not in the sources", section=MEMO_SECTION[section],
                   source=SECTION_SOURCE[section])
        return "\n".join(f"{i}. {s.purpose} [facts: {'; '.join(s.facts_used)}]" for i, s in enumerate(approved, 1))

    def write_section(self, section: str, f: DealFacts, x: Extracted) -> str:
        opening_marked = None
        if section == "transaction_overview":
            openers = self.bp.get("openers", {})
            variant = (f.structure if f.loan else "participation_no_share") if f.is_participation else None
            tpl = openers.get("structure", {}).get(variant) if variant else None
            tpl = tpl or openers.get("transaction_overview", {}).get(f.loan_type)
            opening_marked = render(tpl, f, MEMO_SECTION[section]) if tpl else None
        opening_plain = strip_markers(opening_marked).rstrip() if opening_marked else None
        playbook = self.pack.playbook(section)
        system = self.pack.system(section, f)
        user = self._user_prompt(section, f, x, opening_plain)
        allowed = allowed_numbers(json.dumps(_fmt_facts(f), default=str), self._sources_for(section, f, x),
                                  f.notes, opening_plain or "")
        self._debug(section, "system prompt", system)
        self._debug(section, "user prompt", user)

        plan_text = self._plan(section, system, user, allowed, f)                  # 1. plan
        prompt = user + (f"\nAPPROVED PLAN - write the section from this plan, in this order, using only these "
                         f"facts:\n{plan_text}\n" if plan_text else "")
        text = self._clean(self.llm.write(system, prompt))                         # 2. write
        self._debug(section, "draft", text)
        problems = check_text(text, playbook, allowed, opening_plain)              # 3. check
        if problems:
            self._debug(section, "check failed", "\n".join(problems))
            repair = (prompt + "\n\nYOUR DRAFT:\n" + text + "\n\nREWRITE THE DRAFT TO FIX EXACTLY THESE PROBLEMS "
                      "(keep everything else; output only the corrected section):\n- " + "\n- ".join(problems))
            fixed = self._clean(self.llm.write(system, repair))
            self._debug(section, "repaired", fixed)
            still = check_text(fixed, playbook, allowed, opening_plain)
            if len(still) <= len(problems):
                text, problems = fixed, still
            if problems:
                self._debug(section, "check failed after repair", "\n".join(problems))
        text = self._remarks(text, f, section)
        tail = ""
        if problems:
            fl = f.flag("review", f"{SECTION_TITLES[section]} - the wording failed the style check after one repair: "
                                  + " ".join(problems), title="wording needs a look (style check)",
                        section=MEMO_SECTION[section], source=SECTION_SOURCE[section])
            tail = " " + f.marker(fl)
        # keep the fixed opener as template text
        if opening_plain and text.startswith(opening_plain):
            return opening_marked + ai(text[len(opening_plain):]) + tail
        paras = [ai(p.strip()) for p in text.split("\n\n") if p.strip()]
        return "\n\n".join(paras) + tail

    def write_all(self, f: DealFacts, x: Extracted) -> dict[str, str]:
        results: dict[str, str] = {}

        def job(sec):
            try:
                results[sec] = self.write_section(sec, f, x)
                self.log(f"  wrote {sec.replace('_', ' ')}")
            except Exception as exc:
                results[sec] = f.new_remark("missing", "section could not be written - write it by hand",
                                            f"The {SECTION_TITLES.get(sec, sec)} could not be written by the AI ({exc}).",
                                            MEMO_SECTION.get(sec, ""), SECTION_SOURCE.get(sec, ""))
                self.log(f"  ERROR writing {sec}: {exc!r}")

        with ThreadPoolExecutor(max_workers=int(self.s.get("openai.max_parallel_requests", 4))) as pool:
            list(pool.map(job, SECTIONS))
        return results
