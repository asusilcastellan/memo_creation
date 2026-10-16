"""The app's cache folder (paths.cache_dir()). Nothing here is needed by the analyst; it is safe to delete.

    downloads/                      Dropbox files fetched for reading
    runs/<deal-slug>/<timestamp>/   debug.log (+ facts.json and sources report when debug.keep_run_files)
    plans/<deal-slug>.json          the analyst's source choices for a deal (re-applied on the next check)
"""
from __future__ import annotations

import hashlib
import re
import shutil
import time
from datetime import datetime
from pathlib import Path

from . import paths


def _sub(name: str) -> Path:
    p = paths.cache_dir() / name
    p.mkdir(parents=True, exist_ok=True)
    return p


def downloads_dir() -> Path:
    return _sub("downloads")


def runs_dir() -> Path:
    return _sub("runs")


def plans_dir() -> Path:
    return _sub("plans")


def deal_slug(deal_path: str) -> str:
    """'/Bridge Lending/Current Portfolio/150 E Aurora St' -> '150-e-aurora-st-1a2b3c' (unique per full path)."""
    norm = deal_path.strip().replace("\\", "/").rstrip("/").lower()
    name = re.split(r"/", norm)[-1] if norm else "deal"
    slug = re.sub(r"[^a-z0-9]+", "-", name).strip("-")[:50] or "deal"
    return f"{slug}-{hashlib.sha1(norm.encode()).hexdigest()[:6]}"


def plan_file(deal_path: str) -> Path:
    return plans_dir() / f"{deal_slug(deal_path)}.json"


def new_run_dir(deal_path: str, kind: str = "memo") -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    d = runs_dir() / deal_slug(deal_path) / f"{stamp}-{kind}"
    n = 1
    while d.exists():
        n += 1
        d = runs_dir() / deal_slug(deal_path) / f"{stamp}-{kind}-{n}"
    d.mkdir(parents=True)
    return d


def _age_days(p: Path) -> float:
    try:
        return (time.time() - p.stat().st_mtime) / 86400
    except OSError:
        return 0.0


def cleanup(retention_days: float = 7) -> int:
    """Delete downloads and run folders older than `retention_days`. Returns the number of items removed."""
    removed = 0
    for f in downloads_dir().iterdir():
        if _age_days(f) > retention_days:
            try:
                shutil.rmtree(f) if f.is_dir() else f.unlink()
                removed += 1
            except OSError:
                pass
    for deal in runs_dir().iterdir():
        if not deal.is_dir():
            continue
        for run in deal.iterdir():
            items = [run] + (list(run.rglob("*")) if run.is_dir() else [])
            if min(_age_days(x) for x in items) > retention_days:      # even the newest file is old
                shutil.rmtree(run, ignore_errors=True) if run.is_dir() else run.unlink(missing_ok=True)
                removed += 1
        if not any(deal.iterdir()):
            deal.rmdir()
    return removed


def clear() -> None:
    """Help -> Clear cache: downloads and run logs. Saved source choices (plans/) are kept."""
    for d in (downloads_dir(), runs_dir()):
        shutil.rmtree(d, ignore_errors=True)
        d.mkdir(parents=True, exist_ok=True)
