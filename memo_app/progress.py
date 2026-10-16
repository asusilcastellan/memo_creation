"""Progress reporting: plain-English steps for the analyst, technical detail for the debug log only.

    rep.step("Reading the term sheet")                  a step starts
    rep.done("Reading the term sheet", ok=False,
             note="Couldn't read the term sheet - ...")  a step ends (check mark or warning icon)
    rep.detail("downloading Term Sheet/x.pdf")           debug.log ONLY - never shown in the app
    rep("...")                                           same as rep.detail (so a Reporter can be passed
                                                         wherever a plain log function is expected)

The UI callback receives (event, payload) with event in {"step", "done"} and payload
{label, ok, note, progress}. Everything sent to the UI is passed through `sanitize()`, which removes
file paths, file names, model names, token counts and other technical text.
"""
from __future__ import annotations

import logging
import re
import threading
import traceback
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

CHECK_FOLDER = "Checking the deal folder"
TERM_SHEET = "Reading the term sheet"
APPRAISAL = "Reading the appraisal"
SPONSOR = "Reading sponsor documents"
THIRD_PARTY = "Reading third-party reports"
EMAILS = "Reading emails"
PRIOR = "Checking for earlier deals with this sponsor"
WRITING = "Writing the memo sections"
BUILDING = "Building the Word document"
DONE = "Done"

GENERATE_STEPS = [CHECK_FOLDER, TERM_SHEET, APPRAISAL, SPONSOR, THIRD_PARTY, EMAILS, PRIOR, WRITING, BUILDING, DONE]
SCAN_STEPS = [CHECK_FOLDER, PRIOR, DONE]

UiCallback = Callable[[str, dict], None]


class UserFacingError(RuntimeError):
    """An error whose message is written for the analyst and may be shown as-is."""


# --------------------------------------------------------------------------- sanitising
_EXT = r"(?:pdf|docx?|xlsx|xlsm|xls|eml|msg|json|md|log|txt|png|jpe?g|csv|zip)"
_END = r"(?:\." + _EXT + r"\b|(?=\s[-(]|\s*$|[,;)\n]))"     # a path ends at a file extension or before ' - '
_SANITIZE = [
    (re.compile(r"[A-Za-z]:[\\/](?:[^\\/\n]+[\\/])*[^\\/\n]*?" + _END, re.I), ""),     # C:\... paths (spaces too)
    (re.compile(r"(?<![\w.])/(?:[^/\n]+/)+[^/\n]*?" + _END, re.I), ""),                  # /Bridge Lending/... paths
    (re.compile(r"[^\s,;()'\"]+\." + _EXT + r"\b", re.I), ""),            # file names
    (re.compile(r"\b(?:gpt|o\d|claude)[\w.\-]*\b", re.I), ""),            # model names
    (re.compile(r"\b\d[\d,]*\s*(?:input|output)?\s*tokens?\b", re.I), ""),
    (re.compile(r"\b(?:include|exclude)=\[.*?\]", re.I), ""),
    (re.compile(r"Traceback \(most recent call last\).*", re.S), ""),
]


def sanitize(text: str) -> str:
    t = str(text or "")
    for rx, repl in _SANITIZE:
        t = rx.sub(repl, t)
    t = re.sub(r"\(\s*\)|\[\s*\]", "", t)            # brackets emptied by the rules above
    t = re.sub(r"\s{2,}", " ", t)
    t = re.sub(r"\s+([.,;:])", r"\1", t)
    return t.strip(" -:;,")


# --------------------------------------------------------------------------- reporter
class Reporter:
    def __init__(self, ui: Optional[UiCallback] = None, steps: list[str] | None = None,
                 log_path: Optional[Path] = None):
        self.ui = ui
        self.steps = list(steps or GENERATE_STEPS)
        self.state: dict[str, str] = {}            # label -> running | ok | warn
        self.notes: dict[str, str] = {}
        self._lock = threading.Lock()
        self._buffer: list[str] = []
        self._fh = None
        self.log_path: Optional[Path] = None
        if log_path:
            self.attach_log(log_path)

    # ------------------------------------------------------------------ debug log
    def attach_log(self, path: Path) -> None:
        """Start writing detail() lines to `path` (lines logged before are written first)."""
        path.parent.mkdir(parents=True, exist_ok=True)
        with self._lock:
            self._fh = open(path, "a", encoding="utf-8")
            self.log_path = path
            for line in self._buffer:
                self._fh.write(line + "\n")
            self._buffer.clear()
            self._fh.flush()

    def close(self) -> None:
        with self._lock:
            if self._fh:
                self._fh.close()
                self._fh = None

    def detail(self, msg: str) -> None:
        line = f"{datetime.now():%H:%M:%S} {msg}"
        with self._lock:
            if self._fh:
                self._fh.write(line + "\n")
                self._fh.flush()
            else:
                self._buffer.append(line)

    __call__ = detail

    def exception(self, context: str, exc: BaseException) -> None:
        self.detail(f"ERROR {context}: {exc}\n" + "".join(traceback.format_exception(exc)))

    # ------------------------------------------------------------------ steps
    @property
    def progress(self) -> float:
        finished = sum(1 for s in self.steps if self.state.get(s) in ("ok", "warn"))
        return finished / max(1, len(self.steps))

    def _emit(self, event: str, label: str, ok: bool = True, note: str = "") -> None:
        if self.ui:
            self.ui(event, {"label": sanitize(label), "ok": ok, "note": sanitize(note), "progress": self.progress})

    def step(self, label: str) -> None:
        with self._lock:
            if label not in self.steps:
                self.steps.insert(max(0, len(self.steps) - 1), label)
            self.state[label] = "running"
        self.detail(f"--- {label}")
        self._emit("step", label)

    def done(self, label: str, ok: bool = True, note: str = "") -> None:
        with self._lock:
            if label not in self.steps:
                self.steps.insert(max(0, len(self.steps) - 1), label)
            if self.state.get(label) == "warn":
                ok = False                      # a warning is never downgraded to a check mark
            self.state[label] = "ok" if ok else "warn"
            if note:
                self.notes[label] = note
        self.detail(f"--- {label}: {'ok' if ok else 'warning'} {note}")
        self._emit("done", label, ok, note)


def console_ui(event: str, p: dict) -> None:
    """CLI: the same friendly steps, ASCII only (Windows consoles)."""
    if event == "step":
        print(f"[..] {p['label']}")
    else:
        print(f"[{'ok' if p['ok'] else '!!'}] {p['label']}" + (f" - {p['note']}" if p["note"] else ""))
