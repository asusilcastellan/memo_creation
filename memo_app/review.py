"""Review items (facts.Flag) bookkeeping before the memo and the review workbook are built.

    finish_flags(f, x, plan)   extraction errors and sources-check notes become flags; every flag gets the
                               relative paths of its source document's files (from the Sources table)
    report_markdown(...)       debug copy of the review list for the run's cache folder
"""
from __future__ import annotations

from .facts import LF
from .extract import PLAIN_NAMES
from .render import strip_markers


def index_summary(f) -> str:
    if f.is_floating is False:
        return f"Interest rate: fixed {f.fixed_rate:.2f}%" if f.fixed_rate else "Interest rate: fixed (rate not found)"
    idx = f.index_name or "index not found"
    val = f"{f.index_rate:.2f}% ({f.index_rate_source})" if f.index_rate is not None else "no value"
    eff = f"{f.effective_rate:.2f}%" if f.effective_rate is not None else "unknown"
    return f"Index used: {idx} = {val}; effective rate at closing {eff}"


def finish_flags(f, x, plan) -> None:
    """Call once after build_facts (and the writer), before build_memo."""
    for key, err in getattr(x, "error_items", []):
        name = PLAIN_NAMES.get(key, key)
        row = plan.rows.get(key) if plan is not None else None
        section = row.memo_sections[0] if row and row.memo_sections else ""
        f.flag("review", f"Couldn't read the {name}: {err}. The parts of the memo that use it are marked for review.",
               title=f"couldn't read the {name}", section=section, source=key, key=f"read_error:{key}")
    for note in getattr(x, "plan_notes", []) or []:
        if not any(note in fl.details for fl in f.flags):
            f.flag("review", note, title="one pager sheet about another property was not used", section=LF,
                   source="one_pager")
    if plan is not None:
        for fl in f.flags:
            if fl.source_key and not fl.source_files and fl.source_key in plan.rows:
                fl.source_files = [e.rel for e in plan.rows[fl.source_key].files]


def report_markdown(f, x, plan, ui, settings) -> str:
    """Debug copy of the review list (kept in the run folder only with debug.keep_run_files)."""
    lines = [f"# Review items - {f.address or (ui.deal_path if ui else '')}", ""]
    for fl in sorted(f.flags, key=lambda z: (z.level, z.id)):
        lines.append(f"- {fl.id} {fl.level_label} [{fl.memo_section}] {fl.title} :: {strip_markers(fl.details)}")
    lines += ["", f"- {index_summary(f)}", "", "## Files used", ""]
    lines += [f"- {r.display_name}: {r.files_text()}" for r in (plan.rows.values() if plan is not None else [])]
    return "\n".join(lines)
