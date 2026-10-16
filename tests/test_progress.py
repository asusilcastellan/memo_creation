"""The analyst sees plain steps only; technical detail goes to debug.log."""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from memo_app.facts import UserInputs  # noqa: E402
from memo_app.progress import GENERATE_STEPS, Reporter, sanitize  # noqa: E402
from memo_app.settings import load_settings  # noqa: E402
from test_offline import StubLLM, make_deal, offline_rates  # noqa: E402

TECHNICAL = re.compile(r"[A-Za-z]:[\\/]|\.(pdf|docx|xlsx|eml|json|md|log)\b|\bgpt|token|Traceback|include=|/\w+/", re.I)


def test_sanitize():
    assert sanitize(r"Couldn't read C:\deal\Appraisal\final.pdf - see log") == "Couldn't read - see log"
    assert sanitize("downloading /Bridge Lending/Current Portfolio/150 E Aurora St/x.pdf") == "downloading"
    assert sanitize("read Final Appraisal.pdf with gpt-5.4-mini (1,234 tokens)") == "read Final with"
    assert sanitize("include=['*.pdf'] exclude=['x']") == ""
    assert sanitize("Couldn't read the appraisal - the memo will mark that section for review.") == \
        "Couldn't read the appraisal - the memo will mark that section for review."


def test_reporter_never_shows_paths_and_reaches_100(tmp_path, monkeypatch):
    import memo_app.pipeline as pipeline
    monkeypatch.setattr(pipeline, "LLM", StubLLM)
    deal = make_deal(tmp_path)
    (deal / "Appraisal" / "2025-3654 - 150 East Aurora Street_Final Appraisal_11.20.25.pdf").write_bytes(b"not a pdf")
    events: list[tuple[str, dict]] = []
    log = tmp_path / "run" / "debug.log"
    rep = Reporter(lambda e, p: events.append((e, p)), steps=GENERATE_STEPS, log_path=log)
    settings = load_settings()
    settings.openai_api_key = "stub"
    ui = UserInputs(deal_path=str(deal), loan_type="acquisition", output_dir=str(tmp_path / "out"))
    with offline_rates():
        pipeline.generate(ui, settings, rep=rep)
    rep.close()

    shown = [p["label"] + " " + p["note"] for _, p in events]
    assert {p["label"] for _, p in events} <= set(GENERATE_STEPS)
    assert not [s for s in shown if TECHNICAL.search(s)], [s for s in shown if TECHNICAL.search(s)]
    assert events[-1][1]["label"] == "Done" and events[-1][1]["progress"] == 1.0
    warn = [p for e, p in events if e == "done" and p["label"] == "Reading the appraisal"][0]
    assert not warn["ok"] and warn["note"] == "Couldn't read the appraisal - the memo will mark that section for review."
    text = log.read_text(encoding="utf-8")                       # the details are in the debug log
    assert "Final Appraisal_11.20.25.pdf" in text and "Traceback" in text
