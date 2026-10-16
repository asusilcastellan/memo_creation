"""Pick the right file(s) for each memo source inside a deal folder.

Rules live in resources/config/sources.yaml. Every pick carries a human-readable
reason so the analyst can see *why* a file was chosen (shown in the app and in the
sources report saved next to the memo).
"""
from __future__ import annotations

import fnmatch
import re
from dataclasses import dataclass, field
from pathlib import PurePosixPath

from .storage import FileEntry


def norm(s: str) -> str:
    """'Background Authorisation & Credit Report' -> 'backgroundauthorisationcreditreport'"""
    return re.sub(r"[^a-z0-9]", "", s.lower())


@dataclass
class Candidate:
    entry: FileEntry
    score: int
    reasons: list[str]


@dataclass
class Selection:
    key: str
    label: str
    folder: str | None                      # matched folder (relative), None if not found
    files: list[FileEntry] = field(default_factory=list)
    candidates: list[Candidate] = field(default_factory=list)
    note: str = ""
    sheets: list[str] = field(default_factory=list)   # one pager: sheets whose TABLES go in the memo, [] = all
    exclude_sheets: list[str] = field(default_factory=list)   # one pager: sheets about another property (never read)

    @property
    def found(self) -> bool:
        return bool(self.files)

    def summary(self) -> str:
        if not self.folder:
            return f"[{self.label}] folder not found ({self.note})"
        if not self.files:
            return f"[{self.label}] no matching file in '{self.folder}' ({self.note})"
        if self.key == "emails" and len(self.files) > 3:
            tail = f"  ({self.note})" if self.note else ""
            return f"[{self.label}] {len(self.files)} email file(s) in '{self.folder}'{tail}: " + \
                ", ".join(f.name for f in self.files[:3]) + ", ..."
        picks = []
        for f in self.files:
            c = next((c for c in self.candidates if c.entry is f), None)
            why = "; ".join(c.reasons) if c else ""
            picks.append(f"{f.rel}  (modified {f.modified:%Y-%m-%d %H:%M}; {why})")
        others = len(self.candidates) - len(self.files)
        tail = f"  [+{others} other candidate(s) ignored]" if others > 0 else ""
        return f"[{self.label}] " + " | ".join(picks) + tail


def _match_any(name: str, patterns: list[str]) -> str | None:
    low = name.lower()
    for p in patterns or []:
        if fnmatch.fnmatch(low, p.lower()):
            return p
    return None


def find_folder(entries: list[FileEntry], aliases: list[str]) -> str | None:
    """Return the relative path of the best-matching folder ('.' = deal root, used only if no named folder exists)."""
    named = [a for a in aliases if a != "."]
    if not named and "." in aliases:
        return "."
    dirs = [e for e in entries if e.is_dir]
    wanted = [norm(a) for a in named]
    # 1) exact normalised name, shallowest first
    for w in wanted:
        hits = sorted((d for d in dirs if norm(d.name) == w), key=lambda d: d.rel.count("/"))
        if hits:
            return hits[0].rel
    # 2) folder name contains the alias (e.g. "Appraisal - Final" for "Appraisal")
    for w in wanted:
        hits = sorted((d for d in dirs if w and w in norm(d.name)), key=lambda d: d.rel.count("/"))
        if hits:
            return hits[0].rel
    return "." if "." in aliases else None


def select(key: str, rule: dict, entries: list[FileEntry], deal_name: str = "") -> Selection:
    label = rule.get("label", key)
    aliases = [a.replace("<deal>", deal_name) for a in rule.get("folders", []) if deal_name or "<deal>" not in a]
    if rule.get("ordered"):
        return _select_ordered(key, label, rule, entries, aliases)
    folder = find_folder(entries, aliases)
    fallback = bool(rule.get("fallback_any_folder"))
    if folder is None and not fallback:
        return Selection(key, label, None, note=f"looked for: {', '.join(aliases)}")
    cands = _candidates(rule, entries, folder) if folder is not None else []
    note = ""
    if not cands and fallback and folder != ".":
        # RULE (fallback_any_folder): try the other named folders in order, then any file anywhere in the deal
        tried = [folder] if folder else []
        for alias in aliases:
            other = find_folder(entries, [alias]) if alias != "." else None
            if other and other not in tried:
                tried.append(other)
                cands = _candidates(rule, entries, other)
                if cands:
                    folder = other
                    break
        if not cands:
            note = (f"none in {', '.join(repr(t) for t in tried)}; " if tried else "") + "searched the whole deal folder"
            folder = "."
            cands = _candidates({**rule, "recursive": True}, entries, ".")

    n = 1 if rule.get("pick", "best") == "best" else int(rule.get("max_files", 4))
    sel = Selection(key, label, folder, files=[c.entry for c in cands[:n]], candidates=cands, note=note)
    if not cands:
        sel.note = f"include={rule.get('include')} exclude={rule.get('exclude')}"
    return sel


def _select_ordered(key: str, label: str, rule: dict, entries: list[FileEntry], aliases: list[str]) -> Selection:
    """RULE (ordered): try the folders in the listed order ('.' = deal root) and use the first with a match."""
    tried = []
    for alias in aliases:
        folder = "." if alias == "." else find_folder(entries, [alias])
        if folder is None or folder in tried:
            continue
        tried.append(folder)
        cands = _candidates(rule, entries, folder)
        if cands:
            n = 1 if rule.get("pick", "best") == "best" else int(rule.get("max_files", 4))
            note = f"found in '{folder}' after {', '.join(tried[:-1])}" if len(tried) > 1 else ""
            return Selection(key, label, folder, files=[c.entry for c in cands[:n]], candidates=cands, note=note)
    if not tried:
        return Selection(key, label, None, note=f"looked for: {', '.join(aliases)}")
    return Selection(key, label, tried[0], note=f"no match in {', '.join(tried)}; include={rule.get('include')}")


def _candidates(rule: dict, entries: list[FileEntry], folder: str) -> list[Candidate]:
    recursive = rule.get("recursive", True)
    exclude_dirs = {norm(d) for d in rule.get("exclude_dirs", [])}
    prefix = "" if folder == "." else folder + "/"

    cands: list[Candidate] = []
    for e in entries:
        if e.is_dir or not e.rel.startswith(prefix):
            continue
        inner = e.rel[len(prefix):]
        parts = PurePosixPath(inner).parts
        if not recursive and len(parts) > 1:
            continue
        if any(norm(p) in exclude_dirs for p in parts[:-1]):
            continue
        if not _match_any(e.name, rule.get("include", ["*"])):
            continue
        if _match_any(e.name, rule.get("exclude", [])):
            continue

        score, reasons = 0, []
        prefer = rule.get("prefer", [])
        hit = _match_any(e.name, prefer)
        if hit:
            bonus = 10 + 2 * (len(prefer) - prefer.index(hit))
            score += bonus
            reasons.append(f"name matches '{hit}'")
        exts = rule.get("prefer_ext", [])
        if e.suffix in exts:
            score += len(exts) - exts.index(e.suffix)
        if len(parts) > 1:
            score -= 1  # slight preference for files directly in the folder
        cands.append(Candidate(e, score, reasons))

    # best score first, then newest "date modified"
    cands.sort(key=lambda c: (c.score, c.entry.modified), reverse=True)
    for i, c in enumerate(cands):
        if i == 0 or c.score == cands[0].score:
            c.reasons.append("newest of equal-ranked files" if len(cands) > 1 else "only candidate")
    return cands


def locate_all(entries: list[FileEntry], rules: dict, deal_name: str = "") -> dict[str, Selection]:
    return {key: select(key, rule, entries, deal_name) for key, rule in rules.get("sources", {}).items()}


def reorder(selection: Selection, first: FileEntry, reason: str) -> None:
    """Promote a specific candidate to be the chosen file (used after the signature check)."""
    for c in selection.candidates:
        if c.entry is first:
            c.reasons.insert(0, reason)
    selection.files = [first] + [f for f in selection.files if f is not first]
    if len(selection.files) > 1 and selection.key == "term_sheet":
        selection.files = selection.files[:1]
