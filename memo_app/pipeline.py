"""End-to-end flow used by both the GUI and the command line.

    scan()      list the deal folder ONCE, pick the source files, re-apply the analyst's saved choices,
                suggest prior deals with the sponsor -> SourcesPlan (no AI calls)
    generate()  plan -> extract -> facts -> write -> .docx -> review .xlsx  (uses the plan as given; never rescans)
                The output folder receives exactly two files: the .docx and "<memo> - Review.xlsx". debug.log (and,
                with debug.keep_run_files, facts.json + a sources report) go to cache/runs/<deal>/<timestamp>/.

Progress goes to a progress.Reporter: step()/done() for the analyst, detail() for the debug log.
"""
from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from . import cache, paths
from .docx_builder import build_memo
from .emails import collect as collect_emails
from .extract import Extractor
from .facts import UserInputs, build_facts, build_financial_tables, index_quote_for
from .llm import LLM
from .locator import Selection, locate_all, reorder
from .progress import (BUILDING, CHECK_FOLDER, DONE, EMAILS, GENERATE_STEPS, PRIOR, SCAN_STEPS, SPONSOR, WRITING,
                       Reporter, UserFacingError)
from .relationship import find_prior_deals, names_from_files
from .review import finish_flags, index_summary, report_markdown
from .review_workbook import build_review_workbook
from .settings import Settings, load_yaml_resource
from .sources_plan import SourcesPlan, build_plan, pick_sheets
from .storage import open_storage
from .text_extract import pdf_signature_hints
from .writer import Writer

Log = Callable[[str], None]


def _reporter(rep, steps) -> Reporter:
    if isinstance(rep, Reporter):
        return rep
    r = Reporter(steps=steps)
    if callable(rep):                      # a plain log function: send every detail line to it as well
        detail = r.detail
        r.detail = lambda m: (detail(m), rep(m))   # type: ignore[method-assign]
    return r


def deal_name_of(deal_path: str) -> str:
    return re.split(r"[/\\]", deal_path.rstrip("/\\"))[-1]


# --------------------------------------------------------------------------- scan
def scan(ui: UserInputs, settings: Settings, rep: Reporter | Log | None = None, finish: bool = True) -> SourcesPlan:
    rep = _reporter(rep, SCAN_STEPS)
    rep.step(CHECK_FOLDER)
    try:
        storage = open_storage(ui.deal_path, settings, rep)
        entries = storage.walk(ui.deal_path)
    except Exception as exc:
        rep.exception("listing the deal folder", exc)
        rep.done(CHECK_FOLDER, ok=False, note="Couldn't open the deal folder.")
        raise UserFacingError("Couldn't open the deal folder. Check the path (and the Dropbox connection) "
                              "and try again.") from exc
    rep.detail(f"deal folder ({storage.kind}) {ui.deal_path}: {sum(e.is_dir for e in entries)} folders, "
               f"{sum(not e.is_dir for e in entries)} files")
    rules = load_yaml_resource("config", "sources.yaml")
    deal_name = deal_name_of(ui.deal_path)
    selections = locate_all(entries, rules, deal_name=deal_name)
    _promote_signed_term_sheet(storage, selections, rules, rep)

    plan = build_plan(ui.deal_path, selections, entries, rules, storage)
    applied = plan.apply_saved()
    auto_one_pager_sheets(plan, rep)
    for sel in selections.values():
        rep.detail("  auto: " + sel.summary())
    for line in plan.summary_lines():
        rep.detail("  plan: " + line)
    found = sum(r.found for r in plan.rows.values())
    rep.done(CHECK_FOLDER, note=f"{found} of {len(plan.rows)} document types found"
                                + (f", {len(applied)} saved choice(s) re-applied" if applied else ""))

    rep.step(PRIOR)
    plan.prior_deals = suggest_prior_deals(storage, ui, selections, settings, rep)
    rep.done(PRIOR, note=(f"Possible earlier deals: {len(plan.prior_deals)}" if plan.prior_deals else "None found"))
    if finish:
        rep.done(DONE)
    return plan


def _promote_signed_term_sheet(storage, selections: dict[str, Selection], rules: dict, rep: Reporter) -> None:
    """Signed term sheet: prefer a PDF that carries signature evidence."""
    ts = selections.get("term_sheet")
    if not (ts and ts.found and rules["sources"]["term_sheet"].get("signature_check", True)):
        return
    pdfs = [c.entry for c in ts.candidates if c.entry.suffix == ".pdf"][:3]
    for e in pdfs:
        try:
            hints = pdf_signature_hints(storage.fetch(e))
        except Exception as exc:
            rep.detail(f"  signature check failed for {e.rel}: {exc}")
            continue
        if hints:
            reorder(ts, e, "signature evidence: " + ", ".join(hints))
            return
    if pdfs and ts.files[0].suffix != ".pdf":
        reorder(ts, pdfs[0], "PDF preferred over Word draft (signed copies are PDFs)")


def auto_one_pager_sheets(plan: SourcesPlan, rep: Reporter | None = None, notes: bool = True) -> None:
    """Pre-select the one-pager sheets about THIS property (unless the analyst already chose sheets) and remember the
    sheets that clearly belong to another deal (never read). The selection only decides which TABLES go in the
    memo; facts are read from every sheet that is not excluded."""
    row = plan.rows.get("one_pager")
    if not row or not row.files or row.files[0].suffix not in (".xlsx", ".xlsm"):
        return
    try:
        names, chosen, rejected = pick_sheets(plan.storage.fetch(row.files[0]), deal_name_of(plan.deal_path))
    except Exception as exc:
        if rep:
            rep.detail(f"  could not list one-pager sheets: {exc}")
        return
    row.excluded_sheets = [s for s in rejected if s not in row.sheets]
    if row.sheets:
        return
    row.sheets = chosen
    if row.auto:
        row.auto_sheets = list(chosen)
    if not notes:
        return
    for sheet, why in rejected.items():
        plan.notes.append(f"One pager: sheet '{sheet}' was not used ({why}).")
    if rep:
        rep.detail(f"  one pager sheets {names}: using {chosen}, rejected {rejected}")


def suggest_prior_deals(storage, ui: UserInputs, selections: dict[str, Selection], settings: Settings,
                        log: Log = print) -> list[str]:
    """Other deals in the portfolio whose files mention the sponsor's surname (names from file names only)."""
    if not settings.get("memo.prior_relationship_autodetect", True):
        return []
    files = [c.entry for k in ("pfs", "resume") if k in selections for c in selections[k].candidates]
    names = names_from_files(files)
    if not names:
        return []
    log(f"Looking for earlier Castellan deals with: {', '.join(names)}")
    prior = find_prior_deals(storage, ui.deal_path, names, log=log)
    log(f"  prior deals: {prior or 'none'}")
    return prior


@dataclass
class MemoOutput:
    memo: Path                     # the .docx
    review: Optional[Path]         # "<memo> - Review.xlsx", None if it could not be created


REVIEW_FAILED = "The memo was saved, but the review list could not be created."


def _safe(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]+', "", name).strip()


# --------------------------------------------------------------------------- generate
def generate(ui: UserInputs, settings: Settings, plan: Optional[SourcesPlan] = None,
             rep: Reporter | Log | None = None) -> MemoOutput:
    rep = _reporter(rep, GENERATE_STEPS)
    if settings.missing_secrets():
        raise UserFacingError("The OpenAI key is missing from the .env file - please ask IT to add it.")
    if plan is None:
        plan = scan(ui, settings, rep, finish=False)
    else:
        rep.done(CHECK_FOLDER, note="Using the sources you checked")
        rep.done(PRIOR, note=(f"Possible earlier deals: {len(plan.prior_deals)}" if plan.prior_deals else "None found"))
    storage = plan.storage or open_storage(plan.deal_path, settings, rep)
    if plan.storage is None:
        plan.storage = storage
    auto_one_pager_sheets(plan, rep, notes=False)       # the analyst may have picked another workbook since the check
    selections = plan.selections()
    llm = LLM(settings, rep)

    rep.step(EMAILS)
    chain = collect_emails(storage, selections.get("emails"), settings, rep,
                           conversation_id=ui.conversation_id, use_graph=ui.read_graph)
    if not chain.messages:
        rep.done(EMAILS, note="No emails found")

    ex = Extractor(storage, llm, settings, rep)
    x = ex.run(selections, emails=chain, rep=rep, defer={SPONSOR})
    x.plan_notes = list(plan.notes)

    quote = index_quote_for(x, ui, settings, fetch=bool(settings.get("rates.auto_fetch", True)), log=rep)
    f = build_facts(x, ui, settings, index_quote=quote, prior_matches=plan.prior_deals)
    if ui.run_web_search:
        ex.adverse_search([s.name for s in f.sponsors] or ([f.sponsor_names] if f.sponsor_names else []),
                          f.city_state or "", x)
        f = build_facts(x, ui, settings, index_quote=quote, prior_matches=plan.prior_deals)
    ok, note = ex.group_status(SPONSOR)
    rep.done(SPONSOR, ok=ok, note=note)
    rep.detail(index_summary(f))
    for line in plan.notes:
        rep.detail("  " + line)

    rep.step(WRITING)
    bp = load_yaml_resource("templates", "boilerplate.yaml")
    writer = Writer(llm, settings, bp, rep)
    sections = writer.write_all(f, x)
    f.sponsor_experience = sections.get("sponsor_experience", "")
    failed = [fl for fl in f.flags if fl.title.startswith("section could not be written")]
    rep.done(WRITING, ok=not failed,
             note="Some sections could not be written - they are marked for review." if failed else "")

    rep.step(BUILDING)
    tables = build_financial_tables(f, x)
    finish_flags(f, x, plan)
    out_dir = Path(ui.output_dir or paths.default_output_dir())
    out_dir.mkdir(parents=True, exist_ok=True)
    pattern = settings.get("memo.output_filename", "Memo - {address} (DRAFT {date}).docx")
    addr = (f.address or deal_name_of(ui.deal_path)).split(",")[0]
    out_path = out_dir / _safe(pattern.format(address=addr, date=datetime.now().strftime("%Y-%m-%d %H%M")))
    try:
        build_memo(str(out_path), f, sections, tables, x.photos, bp)       # assigns the review IDs R1, R2, ...
    except PermissionError as exc:
        rep.exception("saving the memo", exc)
        rep.done(BUILDING, ok=False, note="Couldn't save the memo - is a file with the same name open in Word?")
        raise UserFacingError("Couldn't save the memo. Close it in Word (or choose another folder) and try again.") from exc
    review_path: Optional[Path] = out_path.with_name(out_path.stem + " - Review.xlsx")
    try:
        build_review_workbook(review_path, f, f.flags, plan, x, llm.usage)
        rep.done(BUILDING)
    except Exception as exc:                  # the memo is kept even if the review list fails
        rep.exception("building the review list", exc)
        try:
            review_path.unlink(missing_ok=True)
        except OSError:
            pass
        review_path = None
        rep.done(BUILDING, ok=False, note=REVIEW_FAILED)

    rep.detail(f"saved memo {out_path}; review list {review_path}")
    rep.detail(f"OpenAI usage: {llm.usage}")
    # Debug files live in the run's cache folder, never next to the memo.
    run_dir = rep.log_path.parent if rep.log_path else cache.new_run_dir(ui.deal_path)
    if not rep.log_path:
        rep.attach_log(run_dir / "debug.log")
    run_files = write_run_files(run_dir, ui, plan, f, x, llm, settings)
    if not settings.get("debug.keep_run_files", False):
        for p in run_files:
            p.unlink(missing_ok=True)
    rep.done(DONE, note="Memo and review list saved" if review_path else REVIEW_FAILED)
    return MemoOutput(out_path, review_path)


def write_run_files(run_dir: Path, ui: UserInputs, plan: SourcesPlan, f, x, llm, settings) -> list[Path]:
    """facts.json + sources report in the run's cache folder (deleted after success unless debug.keep_run_files)."""
    rp = run_dir / "sources report.md"
    rp.write_text(report_markdown(f, x, plan, ui, settings), encoding="utf-8")
    chain = x.emails
    dump = {"inputs": asdict(ui), "facts": asdict(f),
            "extracted": {k: (v.model_dump() if hasattr(v, "model_dump") else None)
                          for k, v in x.__dict__.items()
                          if k not in ("photos", "resume_text", "appraisal_text", "provenance", "errors", "emails")},
            "email_chain": {"summary": chain.summary(), "omitted": chain.omitted, "errors": chain.errors,
                            "messages": [m.meta() for m in chain.messages]} if chain else None,
            "sources_plan": {k: {"files": [e.rel for e in r.files], "sheets": r.sheets, "auto": r.auto,
                                 "skipped": r.skipped} for k, r in plan.rows.items()},
            "provenance": x.provenance, "usage": llm.usage}
    fj = run_dir / "facts.json"
    fj.write_text(json.dumps(dump, indent=2, default=str), encoding="utf-8")
    return [rp, fj]
