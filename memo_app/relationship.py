"""Suggest whether Castellan has worked with the sponsor before (the analyst confirms in the app).

    names_from_files(files)                 'PFS Shaya Margareten.pdf' -> 'Shaya Margareten'   (no LLM)
    find_prior_deals(storage, deal_path, sponsor_names)
        -> ['12 Main St (PFS - Margareten 2024.pdf)', ...]

Only the first two levels of every OTHER deal folder in the portfolio are listed: the deal root plus
its PFS / Resume / Term Sheet / Background sub-folders (folder names from sources.yaml). The
portfolio is never walked recursively, which keeps it to a few Dropbox API calls per deal.
A match is a sponsor surname (>= 4 letters, case-insensitive) inside a file name.
"""
from __future__ import annotations

import re
from concurrent.futures import ThreadPoolExecutor
from pathlib import PurePosixPath, PureWindowsPath
from typing import Callable, Iterable, Optional

from .locator import norm
from .storage import FileEntry, Storage

Log = Callable[[str], None]

# Sub-folders of another deal that are worth listing (in addition to the deal root).
SPONSOR_SOURCES = ("pfs", "resume", "term_sheet", "background")
DEFAULT_SUBFOLDERS = ["PFS & SREO", "PFS", "Personal Financial Statement", "Resume", "Bio",
                      "Term Sheet", "Term Sheets", "LOI", "Background", "Credit Report"]

# Words that appear in PFS / resume file names but are not names.
STOPWORDS = {
    "pfs", "personal", "financial", "financials", "statement", "statements", "sreo", "reo", "schedule", "of", "real",
    "estate", "owned", "resume", "resumes", "bio", "bios", "biography", "cv", "sponsor", "sponsors", "profile",
    "final", "signed", "executed", "updated", "update", "copy", "draft", "and", "the", "for", "with", "llc", "inc",
    "corp", "company", "group", "holdings", "partners", "properties", "capital", "investments", "realty", "mr", "mrs",
    "ms", "dr", "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "sept", "oct", "nov", "dec",
    "january", "february", "march", "april", "june", "july", "august", "september", "october", "november",
    "december", "scan", "scanned", "doc", "page", "pages", "new", "old", "version", "rev", "pdf", "docx", "xlsx",
    "background", "credit", "report", "results", "result", "eaglescreen", "check", "authorization",
    "authorisation", "consent", "form", "disclosure", "due", "diligence", "summary", "info", "information",
}


def names_from_files(files: Iterable[FileEntry]) -> list[str]:
    """Candidate sponsor names from file names only, e.g. 'PFS Shaya Margareten.pdf' -> 'Shaya Margareten'."""
    out: list[str] = []
    for e in files:
        stem = PurePosixPath(e.name).stem
        words = [w for w in re.split(r"[^A-Za-z']+", stem) if len(w) >= 2 and w.lower() not in STOPWORDS]
        if words:
            name = " ".join(w if not w.islower() else w.capitalize() for w in words[:4])
            if name not in out:
                out.append(name)
    return out


def surnames(sponsor_names: Iterable[str], min_len: int = 4) -> list[str]:
    """Last word of each name, letters only, lower-case, at least `min_len` letters."""
    out: list[str] = []
    for n in sponsor_names:
        words = re.findall(r"[A-Za-z]+", n or "")
        if words and len(words[-1]) >= min_len and words[-1].lower() not in out:
            out.append(words[-1].lower())
    return out


def _subfolder_names() -> list[str]:
    try:
        from .settings import load_yaml_resource
        src = load_yaml_resource("config", "sources.yaml").get("sources", {})
        names = [f for k in SPONSOR_SOURCES for f in src.get(k, {}).get("folders", []) if f != "."]
        return names or DEFAULT_SUBFOLDERS
    except Exception:
        return DEFAULT_SUBFOLDERS


def _base(path: str) -> str:
    p = path.rstrip("/\\")
    return (PureWindowsPath(p) if "\\" in p else PurePosixPath(p)).name


def find_prior_deals(storage: Storage, deal_path: str, sponsor_names: Iterable[str],
                     portfolio_root: Optional[str] = None, log: Log | None = None,
                     max_workers: int = 8) -> list[str]:
    """Other deal folders whose PFS / resume / term sheet / root file names contain a sponsor surname.

    portfolio_root defaults to the parent of the deal folder (e.g. /Bridge Lending/Current Portfolio).
    Returns 'Deal folder name (matched file)', one entry per deal, sorted by deal name.
    """
    keys = surnames(sponsor_names)
    if not keys:
        return []
    root = portfolio_root or storage.parent(deal_path)
    this_deal = _base(deal_path).lower()
    try:
        deals = [e for e in storage.list_dir(root) if e.is_dir and e.name.lower() != this_deal]
    except Exception as exc:
        if log:
            log(f"  (prior-deal search skipped: cannot list {root}: {exc})")
        return []
    wanted = [norm(a) for a in _subfolder_names()]

    def match(files: list[FileEntry]) -> Optional[str]:
        for f in files:
            if f.is_dir:
                continue
            low = f.name.lower()
            if any(k in low for k in keys):
                return f.name
        return None

    def scan_deal(deal: FileEntry) -> Optional[str]:
        try:
            level1 = storage.list_dir(deal.path)
            hit = match(level1)
            if hit:
                return f"{deal.name} ({hit})"
            for sub in level1:
                n = norm(sub.name)
                if sub.is_dir and any(w and (n == w or w in n) for w in wanted):
                    hit = match(storage.list_dir(sub.path))
                    if hit:
                        return f"{deal.name} ({sub.name}/{hit})"
        except Exception as exc:
            if log:
                log(f"  (prior-deal search: could not list {deal.name}: {exc})")
        return None

    with ThreadPoolExecutor(max_workers=max(1, max_workers)) as pool:
        results = list(pool.map(scan_deal, deals))
    return sorted(r for r in results if r)
