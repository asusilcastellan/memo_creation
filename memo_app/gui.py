"""Tkinter front end for analysts.

Flow: fill in the deal folder -> "1. Check sources" shows the Sources table (one row per document type,
double-click to change the file) -> "2. Generate memo". Progress is shown as plain-English steps;
technical detail goes to the run's debug.log in the cache folder (Help -> Open diagnostics folder).
Long work runs in a background thread; events come back through a queue.

Look and feel (colours, fonts, ttk styles) lives in ui_theme.py; this module only lays widgets out (grid only).
"""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from tkinter import font as tkfont

from . import __version__, cache, email_lookup, paths, rates
from . import ui_theme as T
from .facts import LOAN_TYPES, UserInputs
from .progress import GENERATE_STEPS, SCAN_STEPS, Reporter, UserFacingError
from .settings import load_settings
from .sources_plan import ROOT_LABEL, SourcesPlan, pick_sheets

PROPERTY_TYPES = ["(detect)", "Multifamily", "Industrial", "Retail", "Office", "Single-family", "Condominium",
                  "Mixed-use", "Land", "Hospitality", "Self-storage", "Other"]
EXITS = ["(infer from notes)", "Permanent refinance", "Sale of the property", "Condo / unit sellout",
         "Lease-up then refinance", "Sale of entitled land", "Refinance with construction loan"]
GENERIC_ERROR = ("Something went wrong and the memo was not created.\n\n"
                 "Details were saved for IT (Help → Open diagnostics folder).")
NOTES_HINT = ("e.g. The borrower is buying the property to renovate and lease it up; "
              "the exit is a refinance within 24 months.")
EMPTY_TABLE = "Enter the deal folder and click 1. Check sources"
LABEL_COL, BUTTON_COL = 150, 96          # card grid: label column / button column minimum widths (px)
ROW_PAD = 6
TIP_DELAY_MS = 400


def open_path(p: str | Path) -> None:
    if sys.platform == "win32":
        os.startfile(str(p))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(p)])
    else:
        subprocess.Popen(["xdg-open", str(p)])


# --------------------------------------------------------------------------- tooltips
class _Tip:
    """A small white tooltip window shown after TIP_DELAY_MS."""

    def __init__(self, widget: tk.Misc):
        self.widget, self.win, self.job = widget, None, None

    def schedule(self, text: str, x: int, y: int):
        self.cancel()
        if text:
            self.job = self.widget.after(TIP_DELAY_MS, lambda: self._show(text, x, y))

    def _show(self, text: str, x: int, y: int):
        self.job = None
        self.win = tk.Toplevel(self.widget)
        self.win.wm_overrideredirect(True)
        self.win.wm_geometry(f"+{x + 16}+{y + 14}")
        frame = tk.Frame(self.win, background="white", highlightthickness=1, highlightbackground=T.BORDER)
        frame.grid(row=0, column=0)
        tk.Label(frame, text=text, justify="left", background="white", foreground=T.TEXT, font=T.font("small"),
                 padx=8, pady=5, wraplength=520).grid(row=0, column=0)

    def cancel(self):
        if self.job:
            self.widget.after_cancel(self.job)
            self.job = None
        if self.win:
            self.win.destroy()
            self.win = None


class _ButtonTip:
    """Plain hover tooltip for a widget."""

    def __init__(self, widget: tk.Misc, text: str):
        self.tip = _Tip(widget)
        widget.bind("<Enter>", lambda e: self.tip.schedule(text, e.x_root, e.y_root), add="+")
        widget.bind("<Leave>", lambda e: self.tip.cancel(), add="+")
        widget.bind("<ButtonPress>", lambda e: self.tip.cancel(), add="+")


class TreeTooltip:
    """Hover help for a cell of the Sources table: text_for(row_id, column_id) -> text or ""."""

    def __init__(self, tree: ttk.Treeview, text_for):
        self.tree, self.text_for, self.cell = tree, text_for, None
        self.tip = _Tip(tree)
        tree.bind("<Motion>", self._move, add="+")
        tree.bind("<Leave>", lambda e: self._hide(), add="+")

    def _move(self, e):
        cell = (self.tree.identify_row(e.y), self.tree.identify_column(e.x))
        if cell == self.cell:
            return
        self._hide()
        self.cell = cell
        if cell[0]:
            self.tip.schedule(self.text_for(*cell), e.x_root, e.y_root)

    def _hide(self):
        self.tip.cancel()
        self.cell = None


# --------------------------------------------------------------------------- choose source dialog
class ChooseSourceDialog(tk.Toplevel):
    """Pick the folder + file(s) (and one-pager sheets) for one row. Sets self.result:
    ("use", files, sheets) | ("skip",) | ("reset",) | None (cancel)."""

    def __init__(self, parent, plan: SourcesPlan, key: str):
        super().__init__(parent)
        self.plan, self.key, self.row = plan, key, plan.rows[key]
        self.result = None
        self._sheet_names: list[str] = []
        self.title(f"Choose source - {self.row.display_name}")
        self.configure(background=T.BG)
        self.transient(parent)
        self.geometry("860x620")
        self.minsize(760, 520)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        outer, frm = T.card(self, f"{self.row.display_name}")
        outer.grid(row=0, column=0, sticky="nsew", padx=16, pady=(16, 8))
        frm.columnconfigure(0, minsize=LABEL_COL)
        frm.columnconfigure(1, weight=1)
        ttk.Label(frm, text=f"Used for {self.row.used_for}" + (f" · feeds {self.row.sections_text}" if self.row.sections_text else ""),
                  style="Muted.TLabel", wraplength=640, justify="left").grid(row=1, column=0, columnspan=3, sticky="w",
                                                                             pady=(0, 10))
        ttk.Label(frm, text="Folder").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=ROW_PAD)
        self.folder_var = tk.StringVar()
        folders = plan.folders()
        cur = self.row.folder if self.row.folder not in (None, ".", "") else ROOT_LABEL
        self.folder_var.set(cur if cur in folders else ROOT_LABEL)
        cb = ttk.Combobox(frm, textvariable=self.folder_var, values=folders, state="readonly")
        cb.grid(row=2, column=1, columnspan=2, sticky="ew", pady=ROW_PAD)
        cb.bind("<<ComboboxSelected>>", lambda e: self._fill_files())

        ttk.Label(frm, text="Files" if self.row.multi else "File").grid(row=3, column=0, sticky="nw", padx=(0, 12),
                                                                         pady=ROW_PAD)
        self.files = tk.Listbox(frm, selectmode="extended" if self.row.multi else "browse", exportselection=False,
                                height=10)
        T.style_listbox(self.files)
        self.files.grid(row=3, column=1, sticky="nsew", pady=ROW_PAD)
        fscroll = ttk.Scrollbar(frm, orient="vertical", command=self.files.yview)
        fscroll.grid(row=3, column=2, sticky="ns", pady=ROW_PAD)
        self.files.configure(yscrollcommand=fscroll.set)
        self.files.bind("<<ListboxSelect>>", lambda e: self._file_selected())
        hint = "Newest first. Ctrl/Shift-click to select several." if self.row.multi else "Newest first."
        ttk.Label(frm, text=hint, style="Muted.TLabel").grid(row=4, column=1, sticky="w")
        self.show_all = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="Show all file types", variable=self.show_all,
                        command=self._fill_files).grid(row=5, column=1, sticky="w", pady=ROW_PAD)
        frm.rowconfigure(3, weight=3)

        self.sheets = None
        if key == "one_pager":
            ttk.Label(frm, text="Sheet(s)").grid(row=6, column=0, sticky="nw", padx=(0, 12), pady=ROW_PAD)
            self.sheets = tk.Listbox(frm, selectmode="extended", exportselection=False, height=5)
            T.style_listbox(self.sheets)
            self.sheets.grid(row=6, column=1, sticky="nsew", pady=ROW_PAD)
            ttk.Label(frm, text="Pre-selected by the deal's street address. The sheets decide which tables appear in "
                                "Loan Financials.", style="Muted.TLabel", wraplength=560, justify="left"
                      ).grid(row=7, column=1, sticky="w")
            frm.rowconfigure(6, weight=1)

        bar = ttk.Frame(self, style="App.TFrame")
        bar.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 16))
        bar.columnconfigure(2, weight=1)
        ttk.Button(bar, text="None (skip this document)", style="Ghost.TButton",
                   command=lambda: self._finish(("skip",))).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(bar, text="Reset to automatic", style="Secondary.TButton",
                   command=lambda: self._finish(("reset",))).grid(row=0, column=1)
        ttk.Button(bar, text="Cancel", style="Ghost.TButton", command=lambda: self._finish(None)
                   ).grid(row=0, column=3, padx=(0, 8))
        ttk.Button(bar, text="Use selected", style="Primary.TButton", command=self._use).grid(row=0, column=4)

        self._entries = []
        self._fill_files(preselect=True)
        T.center_over(self, parent)
        T.color_title_bar(self)
        self.grab_set()
        self.bind("<Escape>", lambda e: self._finish(None))
        self.protocol("WM_DELETE_WINDOW", lambda: self._finish(None))
        self.wait_window(self)

    def _fill_files(self, preselect: bool = False):
        self._entries = self.plan.files_in(self.folder_var.get(), self.key, self.show_all.get())
        self.files.delete(0, "end")
        for e in self._entries:
            self.files.insert("end", f"{e.name}      ({e.modified:%Y-%m-%d %H:%M})")
        if preselect:
            current = {e.rel for e in self.row.files}
            for i, e in enumerate(self._entries):
                if e.rel in current:
                    self.files.selection_set(i)
            self._file_selected(keep=self.row.sheets)

    def _file_selected(self, keep: list[str] | None = None):
        if self.sheets is None:
            return
        self.sheets.delete(0, "end")
        sel = [self._entries[i] for i in self.files.curselection()]
        if not sel or sel[0].suffix not in (".xlsx", ".xlsm"):
            return
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            names, chosen, rejected = pick_sheets(self.plan.storage.fetch(sel[0]), Path(self.plan.deal_path).name)
        except Exception:
            names, chosen, rejected = [], [], {}
        finally:
            self.configure(cursor="")
        chosen = [s for s in (keep or []) if s in names] or chosen
        for i, n in enumerate(names):
            self.sheets.insert("end", n + ("   (other property)" if n in rejected else ""))
            if n in chosen:
                self.sheets.selection_set(i)
        self._sheet_names = names

    def _use(self):
        files = [self._entries[i] for i in self.files.curselection()]
        if not files:
            messagebox.showinfo("Choose source", "Select a file first (or choose 'None').", parent=self)
            return
        sheets = []
        if self.sheets is not None and self._sheet_names:
            sheets = [self._sheet_names[i] for i in self.sheets.curselection()]
        self._finish(("use", files, sheets))

    def _finish(self, result):
        self.result = result
        self.grab_release()
        self.destroy()


# --------------------------------------------------------------------------- email lookup dialog
class EmailLookupDialog(tk.Toplevel):
    """Pick the board row for this deal so its Outlook conversation ID can be read.
    Sets self.result to the chosen email_lookup.LookupRow, or None on cancel."""

    def __init__(self, parent, rows, query, cfg):
        super().__init__(parent)
        self.rows, self.cfg = rows, cfg
        self.result = None
        self.title("Find the deal's email thread")
        self.configure(background=T.BG)
        self.transient(parent)
        self.geometry("920x560")
        self.minsize(760, 440)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        outer, frm = T.card(self, "Which deal is this?")
        outer.grid(row=0, column=0, sticky="nsew", padx=16, pady=(16, 8))
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(3, weight=1)
        ttk.Label(frm, text="Matched on the deal folder name. Pick the right row - its email thread will be read "
                            "from Outlook. Type a different address and Search to refine.",
                  style="Muted.TLabel", wraplength=820, justify="left").grid(row=0, column=0, columnspan=3,
                                                                             sticky="w", pady=(0, 10))

        ttk.Label(frm, text="Search").grid(row=1, column=0, sticky="w", padx=(0, 12), pady=ROW_PAD)
        self.search_var = tk.StringVar(value=query or "")
        search = ttk.Entry(frm, textvariable=self.search_var)
        search.grid(row=1, column=1, sticky="ew", pady=ROW_PAD)
        ttk.Button(frm, text="Search", style="Secondary.TButton", command=self._search).grid(row=1, column=2,
                                                                                             sticky="ew", padx=(8, 0))
        search.bind("<Return>", lambda e: self._search())

        self.note = tk.StringVar()
        ttk.Label(frm, textvariable=self.note, style="Muted.TLabel").grid(row=2, column=0, columnspan=3, sticky="w")

        cols = ["score"] + list(cfg.display_columns)
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse", height=12)
        headings = {"score": "Match"}
        widths = {"score": 70, "Name": 250, "Broker": 150, "Loan Request": 120, "Loan Type": 120, "Asset Type": 120}
        for c in cols:
            self.tree.heading(c, text=headings.get(c, c))
            self.tree.column(c, width=widths.get(c, 130), anchor="center" if c == "score" else "w",
                             stretch=(c == "Name"))
        self.tree.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=ROW_PAD)
        tscroll = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)
        tscroll.grid(row=3, column=2, sticky="ns", pady=ROW_PAD)
        self.tree.configure(yscrollcommand=tscroll.set)
        self.tree.bind("<Double-1>", lambda e: self._use())

        bar = ttk.Frame(self, style="App.TFrame")
        bar.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 16))
        bar.columnconfigure(0, weight=1)
        ttk.Button(bar, text="Cancel", style="Ghost.TButton", command=lambda: self._finish(None)
                   ).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(bar, text="Use selected", style="Primary.TButton", command=self._use).grid(row=0, column=2)

        self._matches: list = []
        self._search()
        T.center_over(self, parent)
        T.color_title_bar(self)
        self.grab_set()
        self.bind("<Escape>", lambda e: self._finish(None))
        self.protocol("WM_DELETE_WINDOW", lambda: self._finish(None))
        search.focus_set()
        self.wait_window(self)

    def _search(self):
        self._matches = email_lookup.match(self.rows, self.search_var.get(), self.cfg)
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self._matches):
            vals = [f"{r.score}%"] + [r.values.get(c, "") for c in self.cfg.display_columns]
            self.tree.insert("", "end", iid=str(i), values=vals)
        if not self.rows:
            self.note.set("No deals with a conversation ID were found in the board workbook.")
        elif not self._matches:
            self.note.set("No rows to show. Type an address and Search.")
        elif self._matches[0].score < self.cfg.threshold:
            self.note.set(f"No match at or above {self.cfg.threshold}% - showing the closest rows. Refine your search.")
        else:
            self.note.set(f"Showing matches at or above {self.cfg.threshold}%. Pick the correct deal.")

    def _use(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Email lookup", "Select the deal's row first (or Cancel).", parent=self)
            return
        self._finish(self._matches[int(sel[0])])

    def _finish(self, result):
        self.result = result
        self.grab_release()
        self.destroy()


# --------------------------------------------------------------------------- done dialog
class DoneDialog(tk.Toplevel):
    """'Memo and review list saved.' with Open memo | Open review list | Open folder | Close."""

    def __init__(self, parent, out):
        super().__init__(parent)
        self.title("Done")
        self.configure(background=T.BG)
        self.transient(parent)
        self.resizable(False, False)
        frm = ttk.Frame(self, style="App.TFrame", padding=20)
        frm.grid(row=0, column=0, sticky="nsew")
        ok = bool(out.review)
        msg = "Memo and review list saved." if ok else "The memo was saved, but the review list could not be created."
        ttk.Label(frm, text=("✓ " if ok else "") + msg, style="DialogTitle.TLabel" if ok else "StatusError.TLabel"
                  ).grid(row=0, column=0, sticky="w")
        ttk.Label(frm, text=Path(out.memo).name + (f"\n{Path(out.review).name}" if out.review else ""),
                  style="DialogMuted.TLabel", justify="left").grid(row=1, column=0, sticky="w", pady=(6, 16))
        btns = ttk.Frame(frm, style="App.TFrame")
        btns.grid(row=2, column=0, sticky="e")
        ttk.Button(btns, text="Open memo", style="Primary.TButton",
                   command=lambda: self._open(out.memo)).grid(row=0, column=0, padx=(0, 8))
        review = ttk.Button(btns, text="Open review list", style="Secondary.TButton",
                            command=lambda: self._open(out.review))
        review.grid(row=0, column=1, padx=(0, 8))
        if not ok:
            review.state(["disabled"])
        ttk.Button(btns, text="Open folder", style="Ghost.TButton",
                   command=lambda: self._open(Path(out.memo).parent)).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(btns, text="Close", style="Ghost.TButton", command=self.destroy).grid(row=0, column=3)
        self.bind("<Escape>", lambda e: self.destroy())
        T.center_over(self, parent)
        T.color_title_bar(self)
        self.grab_set()

    def _open(self, p):
        if p:
            open_path(p)
        self.destroy()


# --------------------------------------------------------------------------- main window
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Castellan Memo Builder")
        self.style = T.apply_theme(self)
        self.scale = T.scale(self)
        self._size_window()
        self.q: queue.Queue = queue.Queue()
        self.busy = False
        self.last_output: Path | None = None
        self.plan: SourcesPlan | None = None
        self._notes_placeholder = False
        try:
            self.settings = load_settings()
        except Exception as exc:
            messagebox.showerror("Settings", f"Could not load settings:\n{exc}")
            raise
        self._build()
        T.color_title_bar(self)
        self.after(100, self._poll)
        self._startup()

    # ------------------------------------------------------------------ layout helpers
    def px(self, n: float) -> int:
        return int(round(n * self.scale))

    def _size_window(self):
        """1180x980 (scaled for high-DPI screens), never larger than the screen; maximised when the screen is small."""
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(self.px(1180), sw - 40), min(self.px(980), sh - 80)
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")
        self.minsize(min(self.px(1000), sw - 40), min(self.px(820), sh - 80))
        if (w < self.px(1180) or h < self.px(980)) and sys.platform == "win32":
            try:
                self.state("zoomed")
            except tk.TclError:
                pass

    def _fit_notes(self):
        """Small screens: give the Sources table room by showing 3 rows of notes instead of 5."""
        try:
            if self.tree.winfo_height() < self.px(200) and int(self.notes.cget("height")) > 3:
                self.notes.configure(height=3)
        except tk.TclError:
            pass

    @staticmethod
    def _label(parent, row: int, text: str) -> ttk.Label:
        lbl = ttk.Label(parent, text=text)
        lbl.grid(row=row, column=0, sticky="w", padx=(0, 12), pady=ROW_PAD)
        return lbl

    @staticmethod
    def _grid_input(widget, row: int, columnspan: int = 1):
        widget.grid(row=row, column=1, columnspan=columnspan, sticky="ew", pady=ROW_PAD)
        return widget

    @staticmethod
    def _field_hint(entry: ttk.Entry, var: tk.StringVar, text: str):
        """A small muted hint ("optional") inside the right end of a field; hidden while the field has text."""
        hint = ttk.Label(entry.master, text=text, style="FieldHint.TLabel")
        info = entry.grid_info()
        hint.grid(row=info["row"], column=info["column"], sticky="e", padx=(0, 10))
        hint.bind("<Button-1>", lambda e: entry.focus_set())

        def update(*_):
            (hint.grid_remove if var.get() or str(entry.cget("state")) == "disabled" and text == "optional"
             else hint.grid)()
        var.trace_add("write", update)
        update()
        return hint

    def _card_columns(self, frm):
        frm.columnconfigure(0, minsize=self.px(LABEL_COL))
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(2, minsize=self.px(BUTTON_COL))

    # ------------------------------------------------------------------ layout
    def _build(self):
        menubar = tk.Menu(self)
        helpm = tk.Menu(menubar, tearoff=False)
        helpm.add_command(label="Open diagnostics folder", command=lambda: open_path(paths.cache_dir()))
        helpm.add_command(label="Clear cache", command=self._clear_cache)
        helpm.add_separator()
        helpm.add_command(label=f"Version {__version__}", state="disabled")
        menubar.add_cascade(label="Help", menu=helpm)
        self.config(menu=menubar)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        root = ttk.Frame(self, style="App.TFrame", padding=16)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(4, weight=1)                      # the Sources card takes the remaining height

        self._build_header(root)
        cards = ttk.Frame(root, style="App.TFrame")
        cards.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        cards.columnconfigure(0, weight=1, uniform="cards")
        cards.columnconfigure(1, weight=1, uniform="cards")
        cards.rowconfigure(0, weight=1)
        self._build_deal_card(cards)
        self._build_rates_card(cards)
        self._build_notes(root)
        self._build_actions(root)
        self._build_sources(root)
        self._build_footer(root)

        self.bind_all("<Control-g>", self._generate_shortcut)
        self.bind_all("<Control-G>", self._generate_shortcut)
        self.after(300, self._fit_notes)

    def _build_header(self, root):
        head = tk.Frame(root, background=T.NAVY, height=64)
        head.grid(row=0, column=0, sticky="ew")
        head.columnconfigure(0, weight=1)
        tk.Label(head, text="Castellan Memo Builder", background=T.NAVY, foreground="white", font=T.font("title")
                 ).grid(row=0, column=0, sticky="w", padx=18, pady=(8, 0))
        tk.Label(head, text="Loan Investment Memorandum drafts", background=T.NAVY, foreground=T.HEADER_SUB,
                 font=T.font("subtitle")).grid(row=1, column=0, sticky="w", padx=18, pady=(0, 8))
        chips = tk.Frame(head, background=T.NAVY)
        chips.grid(row=0, column=1, rowspan=2, sticky="e", padx=18)
        self.chips = []
        for i, text in enumerate(("①  Check sources", "②  Generate memo")):
            chip = tk.Label(chips, text=text, font=T.font("bold"), padx=14, pady=6)
            chip.grid(row=0, column=i, padx=(8 if i else 0, 0))
            self.chips.append(chip)
        self._set_step(1)

    def _build_deal_card(self, cards):
        outer, frm = T.card(cards, "Deal")
        outer.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.deal_card = frm
        self._card_columns(frm)
        self._label(frm, 1, "Deal folder")
        self.path_var = tk.StringVar()
        self.path_entry = self._grid_input(ttk.Entry(frm, textvariable=self.path_var), 1)
        self.path_entry.bind("<Return>", lambda e: self._scan())
        ttk.Button(frm, text="Browse…", style="Ghost.TButton", command=self._browse_deal
                   ).grid(row=1, column=2, sticky="ew", padx=(8, 0), pady=ROW_PAD)
        self.path_var.trace_add("write", lambda *a: self._clear_plan())

        self._label(frm, 2, "Loan type")
        self.loan_var = tk.StringVar(value="Acquisition")
        self._grid_input(ttk.Combobox(frm, textvariable=self.loan_var, values=list(LOAN_TYPES), state="readonly"), 2)
        self._label(frm, 3, "Property type")
        self.ptype_var = tk.StringVar(value=PROPERTY_TYPES[0])
        self._grid_input(ttk.Combobox(frm, textvariable=self.ptype_var, values=PROPERTY_TYPES, state="readonly"), 3)
        self._label(frm, 4, "Exit strategy")
        self.exit_var = tk.StringVar(value=EXITS[0])
        self._grid_input(ttk.Combobox(frm, textvariable=self.exit_var, values=EXITS, state="readonly"), 4)
        self._label(frm, 5, "Save memo to")
        self.out_var = tk.StringVar(value=str(paths.default_output_dir()))
        self._grid_input(ttk.Entry(frm, textvariable=self.out_var), 5)
        ttk.Button(frm, text="Browse…", style="Ghost.TButton", command=self._browse_out
                   ).grid(row=5, column=2, sticky="ew", padx=(8, 0), pady=ROW_PAD)
        frm.rowconfigure(6, weight=1)                        # spare height at the bottom keeps the rows on top

    def _build_rates_card(self, cards):
        outer, frm = T.card(cards, "Rates & sources")
        outer.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self.rates_card = frm
        self._card_columns(frm)
        # Index for the "effective rate at closing". "1. Check sources" does not read the term sheet,
        # so this box is never changed automatically.
        self._label(frm, 1, "Index")
        self.index_var = tk.StringVar(value=rates.FROM_TERM_SHEET)
        self._grid_input(ttk.Combobox(frm, textvariable=self.index_var, values=list(rates.CHOICES), state="readonly"), 1)
        self._label(frm, 2, "Current index %")
        self.rate_var = tk.StringVar()
        rate_entry = self._grid_input(ttk.Entry(frm, textvariable=self.rate_var), 2)
        self._field_hint(rate_entry, self.rate_var, "optional")
        self.fetch_btn = ttk.Button(frm, text="Fetch", style="Ghost.TButton", command=self._fetch_rate)
        self.fetch_btn.grid(row=2, column=2, sticky="ew", padx=(8, 0), pady=ROW_PAD)
        _ButtonTip(self.fetch_btn, "Fills the current SOFR (NY Fed) or WSJ Prime rate")
        self.rate_info = tk.StringVar()               # result of Fetch; the row only shows when there is a message
        self.rate_info_label = ttk.Label(frm, textvariable=self.rate_info, style="Muted.TLabel", wraplength=440,
                                         justify="left")
        self.rate_info_label.grid(row=3, column=1, columnspan=2, sticky="w")
        self.rate_info.trace_add("write", lambda *a: (self.rate_info_label.grid if self.rate_info.get()
                                                      else self.rate_info_label.grid_remove)())
        self.rate_info_label.grid_remove()
        self._proxy_rate: tuple[str, str] | None = None   # (index, text) when Fetch filled overnight SOFR for Term SOFR

        has_graph = self.settings.has_graph
        self._conv_id = ""          # Outlook conversation ID chosen from the board workbook (email lookup)

        ttk.Separator(frm, orient="horizontal").grid(row=5, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        self.prior_var = tk.BooleanVar(value=False)
        self.web_var = tk.BooleanVar(value=True)
        self.graph_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="Castellan has done a loan with this sponsor before", variable=self.prior_var
                        ).grid(row=6, column=0, columnspan=3, sticky="w")
        self.prior_info = tk.StringVar()
        self.prior_info_label = ttk.Label(frm, textvariable=self.prior_info, style="Muted.TLabel",
                                          wraplength=self.px(520), justify="left")
        self.prior_info_label.grid(row=7, column=0, columnspan=3, sticky="w", padx=(28, 0))
        self._prior_autoticked: str | None = None
        ttk.Checkbutton(frm, text="Run online adverse-media search on the sponsor(s)", variable=self.web_var
                        ).grid(row=8, column=0, columnspan=3, sticky="w")
        self.graph_check = ttk.Checkbutton(frm, text="Read email chain from Outlook", variable=self.graph_var,
                                           command=self._on_graph_toggle,
                                           state="normal" if has_graph else "disabled")
        self.graph_check.grid(row=9, column=0, columnspan=3, sticky="w")
        self.graph_info = tk.StringVar(value="" if has_graph else "Outlook not set up")
        self.graph_info_label = ttk.Label(frm, textvariable=self.graph_info, style="Muted.TLabel",
                                          wraplength=self.px(520), justify="left")
        self.graph_info_label.grid(row=10, column=0, columnspan=3, sticky="w", padx=(28, 0))
        frm.rowconfigure(11, weight=1)

    def _build_notes(self, root):
        outer, frm = T.card(root, padding=12)
        outer.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        frm.columnconfigure(0, weight=1)
        ttk.Label(frm, text="Notes for the memo", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(frm, text="Loan purpose, business plan, exit - this wins over everything else", style="Muted.TLabel"
                  ).grid(row=0, column=1, sticky="e")
        self.notes = scrolledtext.ScrolledText(frm, height=5, wrap="word", font=T.font("base"), background="white",
                                               foreground=T.TEXT, insertbackground=T.NAVY, relief="flat",
                                               borderwidth=0, highlightthickness=1, highlightbackground=T.BORDER,
                                               highlightcolor=T.BLUE, padx=8, pady=6)
        self.notes.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self.notes.bind("<FocusIn>", lambda e: self._notes_focus(True))
        self.notes.bind("<FocusOut>", lambda e: self._notes_focus(False))
        self._notes_focus(False)

    def _build_actions(self, root):
        bar = ttk.Frame(root, style="App.TFrame")
        bar.grid(row=3, column=0, sticky="ew", pady=10)
        bar.columnconfigure(0, weight=1)
        self.open_btn = ttk.Button(bar, text="Open output folder", style="Ghost.TButton", command=self._open_out)
        self.open_btn.grid(row=0, column=1, padx=(0, 8))
        self.scan_btn = ttk.Button(bar, text="1. Check sources", style="Secondary.TButton", command=self._scan)
        self.scan_btn.grid(row=0, column=2, padx=(0, 8))
        self.gen_btn = ttk.Button(bar, text="2. Generate memo", style="Primary.TButton", command=self._generate,
                                  state="disabled")
        self.gen_btn.grid(row=0, column=3)

    def _build_sources(self, root):
        outer, frm = T.card(root, padding=14)
        outer.grid(row=4, column=0, sticky="nsew")
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(1, weight=1)
        ttk.Label(frm, text="Sources", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(frm, text="Double-click a row to change the file", style="Muted.TLabel"
                  ).grid(row=0, column=1, sticky="w", padx=(12, 0))
        self.change_btn = ttk.Button(frm, text="Change…", style="Secondary.TButton", command=self._change_selected,
                                     state="disabled")
        self.change_btn.grid(row=0, column=2, sticky="e")

        table = ttk.Frame(frm, style="Card.TFrame")
        table.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=("doc", "sections", "files"), show="headings", height=8,
                                 selectmode="browse")
        for col, title, width in (("doc", "Document", 240), ("sections", "Used in memo section(s)", 330),
                                  ("files", "File(s) used", 530)):          # about 22% / 30% / 48%
            self.tree.heading(col, text=title, anchor="w")
            self.tree.column(col, width=width, minwidth=100, stretch=True)
        self.tree.tag_configure("odd", background=T.ZEBRA)
        self.tree.tag_configure("missing", foreground=T.RED)
        self.tree.tag_configure("skipped", foreground=T.MUTED)
        yscroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.empty_label = ttk.Label(table, text=EMPTY_TABLE, style="Empty.TLabel")
        self.empty_label.grid(row=0, column=0)           # same cell as the table, shown on top until sources are checked
        self.tree.bind("<Double-1>", lambda e: self._change_selected())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.change_btn.configure(
            state="normal" if self.tree.selection() and not self.busy else "disabled"))
        TreeTooltip(self.tree, self._row_tooltip)

    def _build_footer(self, root):
        foot = ttk.Frame(root, style="App.TFrame")
        foot.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        foot.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(foot, mode="determinate", maximum=100, style="Orange.Horizontal.TProgressbar")
        self.progress.grid(row=0, column=0, sticky="ew")
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(foot, textvariable=self.status_var, style="Status.TLabel")
        self.status_label.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.warn_var = tk.StringVar()                  # step warnings; the line is hidden while empty
        warn = ttk.Label(foot, textvariable=self.warn_var, style="StatusError.TLabel", wraplength=self.px(1100),
                         justify="left")
        warn.grid(row=2, column=0, sticky="w")
        warn.grid_remove()
        self.warn_var.trace_add("write", lambda *a: (warn.grid if self.warn_var.get() else warn.grid_remove)())
        self._steps: dict[str, tuple[str, str]] = {}

    # ------------------------------------------------------------------ startup / cache
    def _startup(self):
        days = float(self.settings.get("debug.cache_retention_days", 7))
        threading.Thread(target=lambda: _quiet(cache.cleanup, days), daemon=True).start()
        if self.settings.missing_secrets():
            self._set_status_text("The OpenAI key is missing - ask IT to set up the .env file before generating memos.",
                                  "error")
        else:
            self._set_status_text("Enter the deal folder and click “1. Check sources”.", "muted")

    def _clear_cache(self):
        if messagebox.askyesno("Clear cache", "Delete downloaded files and diagnostic logs?\n"
                                              "(Your saved source choices are kept.)"):
            cache.clear()
            messagebox.showinfo("Clear cache", "Cache cleared.")

    # ------------------------------------------------------------------ helpers
    def _set_step(self, n: int):
        for i, chip in enumerate(self.chips, start=1):
            if i == n:
                chip.configure(background=T.ORANGE, foreground="white")
            else:
                chip.configure(background=T.LIGHT_BLUE, foreground=T.NAVY)

    def _notes_focus(self, focus: bool):
        if focus and self._notes_placeholder:
            self.notes.delete("1.0", "end")
            self.notes.configure(foreground=T.TEXT)
            self._notes_placeholder = False
        elif not focus and not self.notes.get("1.0", "end").strip():
            self.notes.delete("1.0", "end")
            self.notes.insert("1.0", NOTES_HINT)
            self.notes.configure(foreground=T.MUTED)
            self._notes_placeholder = True

    def _notes_text(self) -> str:
        """The analyst's notes - never the placeholder hint."""
        return "" if self._notes_placeholder else self.notes.get("1.0", "end").strip()

    def _browse_deal(self):
        d = filedialog.askdirectory(title="Select the deal folder")
        if d:
            self.path_var.set(d)

    def _browse_out(self):
        d = filedialog.askdirectory(title="Where should the memo be saved?")
        if d:
            self.out_var.set(d)

    def _open_out(self):
        target = self.last_output.parent if self.last_output else Path(self.out_var.get())
        target.mkdir(parents=True, exist_ok=True)
        open_path(target)

    def _fetch_rate(self):
        choice = self.index_var.get()
        kind = rates.classify(choice)
        if kind not in rates.FETCHABLE:
            why = {"": "Choose the index first (the term sheet has not been read yet).",
                   "ameribor": "AMERIBOR has no free data feed - type the current rate.",
                   "fixed": "A fixed-rate loan needs no index rate."}.get(kind, "This index cannot be fetched.")
            messagebox.showinfo("Index rate", why)
            return
        self.fetch_btn.configure(state="disabled")

        def worker():
            if kind == "prime":
                v, src = rates.prime_rate(self.settings)
            elif kind == "sofr_avg_30":
                v, src = rates.fetch_sofr_30d_avg(), "NY Fed"
            else:
                v, src = rates.fetch_sofr(), "NY Fed"
            self.q.put(("rate", (choice, kind, v, src)))

        threading.Thread(target=worker, daemon=True).start()

    def _rate_fetched(self, choice: str, kind: str, v: float | None, src: str):
        self.fetch_btn.configure(state="normal")
        if v is None:
            self.rate_info.set("Couldn't fetch the rate - type it, or leave it empty to use the floor rate.")
            return
        text = f"{v:.2f}"
        self.rate_var.set(text)
        self._proxy_rate = None
        if kind in ("term_sofr_1m", "term_sofr_3m"):
            self._proxy_rate = (choice, text)
            self.rate_info.set("Overnight SOFR shown - type the CME Term SOFR to be exact (otherwise flagged for review).")
        else:
            self.rate_info.set(f"From {'FRED' if src.startswith('FRED') else ('settings' if 'settings' in src else 'the NY Fed')}.")

    def _inputs(self) -> UserInputs | None:
        path = self.path_var.get().strip().strip('"')
        if not path:
            messagebox.showwarning("Deal folder", "Enter the deal folder path.")
            return None
        index_rate = None
        rate_text = self.rate_var.get().strip()
        if rate_text:
            try:
                index_rate = float(rate_text.rstrip("%"))
            except ValueError:
                messagebox.showwarning("Index rate", "The index rate must be a number like 4.31")
                return None
        if self._proxy_rate == (self.index_var.get(), rate_text):
            index_rate = None   # let the pipeline use overnight SOFR and flag it as a proxy
        index_choice = self.index_var.get()
        ptype = self.ptype_var.get()
        exit_ = self.exit_var.get()
        return UserInputs(
            deal_path=path,
            loan_type=LOAN_TYPES[self.loan_var.get()],
            property_type="" if ptype.startswith("(") else ptype.lower(),
            exit_strategy="" if exit_.startswith("(") else exit_,
            prior_relationship=self.prior_var.get(),
            notes=self._notes_text(),
            index_name_override="" if index_choice == rates.FROM_TERM_SHEET else index_choice,
            index_rate_pct=index_rate,
            run_web_search=self.web_var.get(),
            output_dir=self.out_var.get().strip(),
            read_graph=bool(self.graph_var.get() and self.settings.has_graph),
            conversation_id=self._conv_id if self.settings.has_graph else "",
        )

    # ------------------------------------------------------------------ email lookup (conversation ID)
    def _set_email_thread(self, name: str, keep: bool = False):
        """Show the chosen Outlook thread under the checkbox; keep=False clears the conversation ID."""
        if not keep:
            self._conv_id = ""
        if name:
            self.graph_info.set(f"Email thread: {name}")
        else:
            self.graph_info.set("" if self.settings.has_graph else "Outlook not set up")

    def _on_graph_toggle(self):
        if not self.graph_var.get():                       # unticked -> forget the thread
            self._set_email_thread("")
            return
        if not self.settings.has_graph:
            self.graph_var.set(False)
            return
        if getattr(self, "_lookup_busy", False):
            return
        self._lookup_busy = True
        self.graph_check.configure(state="disabled")
        self.graph_info.set("Looking up the deal in the board…")
        query = email_lookup.deal_name_from_path(self.path_var.get())

        def worker():
            try:
                rows = email_lookup.load_rows(self.settings)
                self.q.put(("lookup", (query, rows, None)))
            except Exception as exc:                        # plain-English message from load_rows
                self.q.put(("lookup", (query, None, str(exc))))

        threading.Thread(target=worker, daemon=True).start()

    def _on_lookup_loaded(self, query, rows, error):
        self._lookup_busy = False
        if self.settings.has_graph:
            self.graph_check.configure(state="normal")
        if error is not None:
            self.graph_var.set(False)
            self._set_email_thread("")
            messagebox.showerror("Email lookup", error, parent=self)
            return
        cfg = email_lookup.LookupConfig.from_settings(self.settings)
        dlg = EmailLookupDialog(self, rows, query, cfg)
        if dlg.result is None:                             # cancelled -> leave the box unticked
            self.graph_var.set(False)
            self._set_email_thread("")
            return
        row = dlg.result
        self._conv_id = row.conversation_id
        name = row.values.get(cfg.display_columns[0], "") if cfg.display_columns else ""
        self._set_email_thread(name or "selected deal", keep=True)

    # ------------------------------------------------------------------ sources table
    def _clear_plan(self):
        self.plan = None
        if getattr(self, "graph_var", None) is not None and self.graph_var.get():
            self.graph_var.set(False)                       # a different deal needs its own thread
            self._set_email_thread("")
        if hasattr(self, "tree"):
            self.tree.delete(*self.tree.get_children())
            self.gen_btn.configure(state="disabled")
            self.change_btn.configure(state="disabled")
            self.prior_info.set("")
            self.empty_label.configure(text=EMPTY_TABLE)
            self.empty_label.grid()
            self._set_step(1)

    def _show_plan(self, plan: SourcesPlan):
        if plan.deal_path != self.path_var.get().strip().strip('"'):
            return                                   # the folder changed while checking
        self.plan = plan
        self._refresh_tree()
        self._show_prior(plan.deal_path, plan.prior_deals)
        self.gen_btn.configure(state="normal")
        self._set_step(2)

    def _refresh_tree(self):
        sel = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        for i, (key, r) in enumerate(self.plan.rows.items()):
            tags = ["odd"] if i % 2 else []
            tag = "skipped" if r.skipped else ("missing" if not r.found else "")
            if tag:
                tags.append(tag)
            self.tree.insert("", "end", iid=key, values=(r.display_name, r.sections_text, r.files_text()),
                             tags=tuple(tags))
        if sel and sel[0] in self.plan.rows:
            self.tree.selection_set(sel[0])
        self.empty_label.grid_remove()

    def _row_tooltip(self, key: str, col: str) -> str:
        """Document cell: what it is used for. Other cells: the full text when the column cuts it off."""
        if not self.plan or key not in self.plan.rows:
            return ""
        r = self.plan.rows[key]
        if col == "#1":
            return f"Used for: {r.used_for}" if r.used_for else ""
        if col == "#2":
            text, cid = r.sections_text, "sections"
        elif col == "#3":
            if len(r.files) > 1 or r.sheets:
                return r.tooltip()
            text, cid = r.files_text(), "files"
        else:
            return ""
        try:
            width = int(self.tree.column(cid, "width"))
            truncated = tkfont.nametofont("TkDefaultFont").measure(text) > width - 12
        except Exception:
            truncated = True
        return text if truncated else ""

    def _change_selected(self):
        if self.busy or not self.plan or not self.tree.selection():
            return
        key = self.tree.selection()[0]
        d = ChooseSourceDialog(self, self.plan, key)
        if not d.result:
            return
        if d.result[0] == "use":
            self.plan.set_files(key, d.result[1], d.result[2])
        elif d.result[0] == "skip":
            self.plan.skip(key)
        else:
            self.plan.reset(key)
        try:
            self.plan.save()
        except OSError:
            pass
        self._refresh_tree()

    def _show_prior(self, deal_path: str, prior: list[str]):
        if not prior:
            self.prior_info.set("")
            return
        shown = "; ".join(p.split(" (")[0] for p in prior[:4]) + (f"; +{len(prior) - 4} more" if len(prior) > 4 else "")
        self.prior_info.set("Found in other deals: " + shown + " - untick if it is a different person")
        if self._prior_autoticked != deal_path:
            self.prior_var.set(True)
            self._prior_autoticked = deal_path

    # ------------------------------------------------------------------ status area
    def _set_status_text(self, text: str, kind: str = "normal"):
        styles = {"normal": "Status.TLabel", "ok": "StatusOk.TLabel", "error": "StatusError.TLabel",
                  "muted": "StatusMuted.TLabel"}
        self.status_label.configure(style=styles.get(kind, "Status.TLabel"))
        self.status_var.set(text)

    def _on_progress(self, event: str, p: dict):
        state = "running" if event == "step" else ("ok" if p["ok"] else "warn")
        old = self._steps.get(p["label"], ("", ""))
        if not (old[0] == "warn" and state == "ok"):
            self._steps[p["label"]] = (state, p["note"] or old[1])
        self.progress["value"] = round(p["progress"] * 100)
        running = [label for label, (st, _) in self._steps.items() if st == "running"]
        current = running[-1] if running else p["label"]
        if current != "Done":
            self._set_status_text(f"{current}…  ({round(p['progress'] * 100)}%)")
        warnings = [f"⚠ {label}: {note}" for label, (st, note) in self._steps.items() if st == "warn" and note]
        self.warn_var.set("\n".join(warnings))

    # ------------------------------------------------------------------ worker plumbing
    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "progress":
                    self._on_progress(*payload)
                elif kind == "plan":
                    self._show_plan(payload)
                elif kind == "rate":
                    self._rate_fetched(*payload)
                elif kind == "lookup":
                    self._on_lookup_loaded(*payload)
                elif kind == "done":
                    self._set_busy(False)
                    if payload is not None and hasattr(payload, "memo"):
                        self.last_output = Path(payload.memo)
                        if payload.review:
                            self._set_status_text("✓ Memo and review list saved", "ok")
                        else:
                            self._set_status_text("The memo was saved, but the review list could not be created.",
                                                  "error")
                        DoneDialog(self, payload)
                    elif self.plan is not None:
                        self._set_status_text("✓ Sources checked - review the table, then click “2. Generate memo”.",
                                              "ok")
                elif kind == "error":
                    self._set_busy(False)
                    self._set_status_text(payload.splitlines()[0], "error")
                    messagebox.showerror("Memo Builder", payload)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _set_busy(self, busy: bool):
        self.busy = busy
        self.scan_btn.configure(state="disabled" if busy else "normal")
        self.gen_btn.configure(state="disabled" if busy or not self.plan else "normal")
        self.change_btn.configure(state="disabled" if busy or not self.tree.selection() else "normal")

    def _start(self, kind: str, steps: list[str], job):
        if self.busy:
            return
        ui = self._inputs()
        if not ui:
            return
        self._set_busy(True)
        self._steps = {}
        self.progress["value"] = 0
        self.warn_var.set("")
        self._set_status_text("Starting…")
        rep = Reporter(lambda e, p: self.q.put(("progress", (e, p))), steps=steps,
                       log_path=cache.new_run_dir(ui.deal_path, kind) / "debug.log")

        def worker():
            try:
                result = job(ui, rep)
                self.q.put(("done", result))
            except UserFacingError as exc:
                rep.exception(kind, exc)
                self.q.put(("error", str(exc)))
            except Exception as exc:
                rep.exception(kind, exc)
                self.q.put(("error", GENERIC_ERROR))
            finally:
                rep.close()

        threading.Thread(target=worker, daemon=True).start()

    def _scan(self):
        from .pipeline import scan

        def job(ui, rep):
            self.q.put(("plan", scan(ui, self.settings, rep)))
        self._set_step(1)
        if self.plan is None:
            self.empty_label.configure(text="Checking the deal folder…")
        self._start("check", SCAN_STEPS, job)

    def _generate(self):
        from .pipeline import generate

        if not self.plan:
            messagebox.showinfo("Generate memo", "Click “1. Check sources” first.")
            return
        plan = self.plan
        try:
            plan.save()
        except OSError:
            pass
        self._set_step(2)
        self._start("memo", GENERATE_STEPS, lambda ui, rep: generate(ui, self.settings, plan, rep))

    def _generate_shortcut(self, event=None):
        if self.gen_btn.instate(["!disabled"]):
            self._generate()
        return "break"


def _quiet(fn, *a):
    try:
        fn(*a)
    except Exception:
        pass


# --------------------------------------------------------------------------- screenshot (hidden CLI flag)
def _grab_window_png(win: tk.Misc, path: Path) -> bool:
    """Save `win` as PNG: PIL.ImageGrab when installed, else (Windows) GDI PrintWindow + a stdlib PNG writer."""
    win.update()
    try:
        from PIL import ImageGrab  # type: ignore[import-not-found]
        x, y = win.winfo_rootx(), win.winfo_rooty()
        ImageGrab.grab(bbox=(x, y, x + win.winfo_width(), y + win.winfo_height())).save(path)
        return True
    except ImportError:
        pass
    if sys.platform != "win32":
        return False
    import ctypes
    import struct
    import zlib
    from ctypes import wintypes

    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    for fn, res, args in ((user32.GetParent, wintypes.HWND, [wintypes.HWND]),
                          (user32.GetWindowDC, wintypes.HDC, [wintypes.HWND]),
                          (user32.ReleaseDC, ctypes.c_int, [wintypes.HWND, wintypes.HDC]),
                          (user32.PrintWindow, wintypes.BOOL, [wintypes.HWND, wintypes.HDC, wintypes.UINT]),
                          (user32.GetWindowRect, wintypes.BOOL, [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]),
                          (gdi32.CreateCompatibleDC, wintypes.HDC, [wintypes.HDC]),
                          (gdi32.CreateCompatibleBitmap, wintypes.HBITMAP, [wintypes.HDC, ctypes.c_int, ctypes.c_int]),
                          (gdi32.SelectObject, wintypes.HGDIOBJ, [wintypes.HDC, wintypes.HGDIOBJ]),
                          (gdi32.DeleteObject, wintypes.BOOL, [wintypes.HGDIOBJ]),
                          (gdi32.DeleteDC, wintypes.BOOL, [wintypes.HDC])):
        fn.restype, fn.argtypes = res, args

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                    ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                    ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]

    hwnd = user32.GetParent(win.winfo_id()) or win.winfo_id()
    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    w, h = rect.right - rect.left, rect.bottom - rect.top
    hdc_win = user32.GetWindowDC(hwnd)
    hdc = gdi32.CreateCompatibleDC(hdc_win)
    bmp = gdi32.CreateCompatibleBitmap(hdc_win, w, h)
    gdi32.SelectObject(hdc, bmp)
    user32.PrintWindow(hwnd, hdc, 2)                     # PW_RENDERFULLCONTENT
    bmi = BITMAPINFOHEADER(40, w, -h, 1, 32, 0, 0, 0, 0, 0, 0)
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(ctypes.c_void_p(hdc), ctypes.c_void_p(bmp), 0, h, buf, ctypes.byref(bmi), 0)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(hdc)
    user32.ReleaseDC(hwnd, hdc_win)
    data, rows = buf.raw, []
    for y in range(h):
        line = data[y * w * 4:(y + 1) * w * 4]
        rgb = bytearray(w * 3)
        rgb[0::3], rgb[1::3], rgb[2::3] = line[2::4], line[1::4], line[0::4]
        rows.append(b"\x00" + bytes(rgb))

    def chunk(kind: bytes, body: bytes) -> bytes:
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", zlib.crc32(kind + body) & 0xFFFFFFFF)

    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + \
        chunk(b"IDAT", zlib.compress(b"".join(rows), 6)) + chunk(b"IEND", b"")
    Path(path).write_bytes(png)
    return True


def _demo_plan(deal_path: str) -> SourcesPlan:
    """Sources table rows for the screenshot (no folder access, no network)."""
    from datetime import datetime, timezone

    from .settings import load_yaml_resource
    from .sources_plan import DEFAULT_NAMES, SourceRow
    from .storage import FileEntry

    rules = load_yaml_resource("config", "sources.yaml").get("sources", {})
    demo = {"term_sheet": ["Term Sheet/Term Sheet - 150 E Aurora St (signed).pdf"], "psa": ["PSA/Executed PSA.pdf"],
            "model": ["150 E Aurora St Model.xlsx"], "one_pager": ["Memo/150 E Aurora one pager.xlsx"],
            "resume": ["Resume/Shaya bio.pdf"], "pfs": ["PFS & SREO/Personal Financial Statement.pdf"],
            "background": ["Background Authorisation & Credit Report/Final Results SMargareten.pdf"],
            "appraisal": ["Appraisal/2025-3654 - 150 East Aurora Street_Final Appraisal_11.20.25.pdf"],
            "environmental": ["Environmental/Phase I ESA - 150 E Aurora.pdf"],
            "insurance": ["Insurance/Evidence of Property Insurance.pdf"],
            "emails": ["Borrower material X/Fwd 150 E Aurora St Bridge Loan Request.eml",
                       "Borrower material X/RE Fwd 150 E Aurora St.eml"], "photos": []}
    when = datetime(2025, 11, 20, 14, 3, tzinfo=timezone.utc)
    rows = {}
    for key, (name, used) in DEFAULT_NAMES.items():
        files = [FileEntry(rel, rel, rel.split("/")[-1], False, when) for rel in demo.get(key, [])]
        rule = rules.get(key, {})
        rows[key] = SourceRow(key, name, used, rule.get("pick", "best") != "best",
                              files[0].parent_rel if files else None, files,
                              memo_sections=list(rule.get("memo_sections", [])),
                              sheets=["150 E Aurora"] if key == "one_pager" else [])
    return SourcesPlan(deal_path, rows, [])


def ui_screenshot(out: str) -> int:
    """`python main.py --ui-screenshot out.png`: the window with sample values and a demo Sources table."""
    T.enable_dpi_awareness()
    app = App()
    deal = "/Bridge Lending/Current Portfolio/150 E Aurora St"
    app.path_var.set(deal)
    app.rate_var.set("4.31")
    app.index_var.set("1M Term SOFR")
    app.notes.focus_set()
    app._notes_focus(True)
    app.notes.insert("1.0", "Acquisition and renovation of a 89,855 SF industrial building; lease-up, then refinance.")
    app.tree.focus_set()
    app._show_plan(_demo_plan(deal))
    app._show_prior(deal, ["12 Main St (PFS & SREO/PFS - S. Margareten 2024.pdf)"])
    app.progress["value"] = 60
    app._set_status_text("Writing the memo sections…  (60%)")
    app.tree.selection_set("appraisal")
    app.update()
    ok = {"saved": False}

    def shoot():
        try:
            ok["saved"] = _grab_window_png(app, Path(out))
        finally:
            app.destroy()

    app.after(800, shoot)
    app.mainloop()
    print(f"Screenshot saved: {out}" if ok["saved"] else "Screenshot not saved (install Pillow, or run on Windows).")
    return 0 if ok["saved"] else 1


def main():
    T.enable_dpi_awareness()
    App().mainloop()
