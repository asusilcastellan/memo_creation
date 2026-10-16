"""Resolve where resources and the .env live, both from source and from a PyInstaller build.

Lookup order for the editable `resources/` folder:
  1. next to the executable / project root  (so the team can edit style + examples after a build)
  2. inside the PyInstaller bundle (sys._MEIPASS)
"""
from __future__ import annotations

import os
import sys
from pathlib import Path


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def bundle_dir() -> Path:
    return Path(getattr(sys, "_MEIPASS", app_dir()))


def resource_dir() -> Path:
    for base in (app_dir(), bundle_dir(), bundle_dir() / "_internal"):
        cand = base / "resources"
        if cand.is_dir():
            return cand
    raise FileNotFoundError("Could not find the 'resources' folder next to the app.")


def resource(*parts: str) -> Path:
    return resource_dir().joinpath(*parts)


def env_file() -> Path:
    return app_dir() / ".env"


def cache_dir() -> Path:
    """Downloads, run logs and saved source choices (never the memo itself).

    Windows: %LOCALAPPDATA%\\CastellanMemo\\cache ; elsewhere ~/.cache/castellan_memo.
    CASTELLAN_CACHE_DIR overrides it (used by the tests).
    """
    override = os.getenv("CASTELLAN_CACHE_DIR")
    if override:
        base = Path(override)
    elif sys.platform == "win32":
        base = Path(os.getenv("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "CastellanMemo" / "cache"
    else:
        base = Path.home() / ".cache" / "castellan_memo"
    base.mkdir(parents=True, exist_ok=True)
    return base


def default_output_dir() -> Path:
    out = Path.home() / "Documents" / "Castellan Memos"
    return out
