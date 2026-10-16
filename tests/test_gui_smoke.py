"""The window builds with the Castellan theme and the inputs of each card line up (skipped without a display)."""
from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


@pytest.fixture(scope="module")
def app(tmp_path_factory):
    """ONE window for the whole module: creating and destroying several Tk interpreters in one process is flaky
    on Windows. The test cache folder is set here because the per-test cache fixture is function-scoped."""
    mp = pytest.MonkeyPatch()
    mp.setenv("CASTELLAN_CACHE_DIR", str(tmp_path_factory.mktemp("gui_cache")))
    from memo_app.gui import App
    try:
        a = App()
    except tk.TclError as exc:
        mp.undo()
        pytest.skip(f"no display available ({exc})")
    a.update()
    yield a
    a.destroy()
    mp.undo()


def test_theme_styles_exist(app):
    from memo_app import ui_theme as T
    style = app.style
    assert style.theme_use() == "clam"
    for name in ("Primary.TButton", "Secondary.TButton", "Ghost.TButton", "Card.TFrame", "Section.TLabel",
                 "Muted.TLabel", "Status.TLabel", "Orange.Horizontal.TProgressbar"):
        assert style.layout(name) or style.configure(name), name
    assert style.lookup("Primary.TButton", "background").lower() == T.ORANGE.lower()
    assert style.lookup("Secondary.TButton", "background").lower() == T.BLUE.lower()
    assert str(app.gen_btn.cget("style")) == "Primary.TButton" and str(app.scan_btn.cget("style")) == "Secondary.TButton"


def test_inputs_line_up_per_card(app):
    for card in (app.deal_card, app.rates_card):
        fields = [w for w in card.winfo_children() if w.winfo_class() in ("TEntry", "TCombobox")]
        assert len(fields) >= 3
        right_edges = [w.winfo_rootx() + w.winfo_width() for w in fields]
        left_edges = [w.winfo_rootx() for w in fields]
        assert max(right_edges) - min(right_edges) <= 2, right_edges
        assert max(left_edges) - min(left_edges) <= 2, left_edges
    browse = [w for w in app.deal_card.winfo_children() if w.winfo_class() == "TButton"]
    assert len(browse) == 2 and abs(browse[0].winfo_rootx() - browse[1].winfo_rootx()) <= 2


def test_notes_placeholder_and_shortcuts(app):
    assert app._notes_text() == ""                              # the hint text is never sent to the pipeline
    app.notes.focus_set()
    app._notes_focus(True)
    app.notes.insert("1.0", "Borrower will lease up and refinance.")
    assert app._notes_text() == "Borrower will lease up and refinance."
    assert app.gen_btn.instate(["disabled"])
    assert app._generate_shortcut() == "break"                  # Ctrl+G does nothing until sources are checked
    assert app.empty_label.winfo_ismapped()
