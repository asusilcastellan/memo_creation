"""Castellan look for the Tkinter app: palette, fonts and ttk styles (pure tkinter/ttk, "clam" theme).

    enable_dpi_awareness()      call BEFORE creating Tk (crisp text on high-DPI Windows screens)
    apply_theme(root) -> Style  fonts + every ttk style used by gui.py
    color_title_bar(window)     Windows 11: NAVY caption with white text (silently skipped elsewhere)
    font(kind)                  "base" | "small" | "section" | "title" | "subtitle" | "fixed" | "bold"

Change colours and fonts here only; gui.py reads everything from this module.
"""
from __future__ import annotations

import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

# ---------------------------------------------------------------- palette
NAVY = "#002060"            # brand, matches the memo tables
BLUE = "#1F4E9A"
BLUE_HOVER = "#173D7A"
LIGHT_BLUE = "#E8EEF8"
ORANGE = "#E8711A"
ORANGE_HOVER = "#C85E0F"
LIGHT_ORANGE = "#FDEBDD"
BG = "#F4F6FA"
CARD = "#FFFFFF"
BORDER = "#D5DBE5"
TEXT = "#1A1F2B"
MUTED = "#5B6475"
RED = "#C00000"
GREEN = "#2E7D32"
DISABLED_BG = "#C9CED8"
DISABLED_FG = "#7A8190"
FIELD_DISABLED = "#EEF0F4"
ZEBRA = "#F7F9FC"
HEADER_SUB = "#C9D6EE"
GHOST_BORDER = "#B4BFD1"            # a little darker than BORDER so white buttons read as buttons

# ---------------------------------------------------------------- fonts
SIZES = {"base": 11, "small": 10, "section": 12, "title": 17, "subtitle": 10, "fixed": 10, "bold": 11}
FAMILY = "Segoe UI" if sys.platform == "win32" else ""     # "" = the platform default (set in apply_theme)
FIXED_FAMILY = "Consolas" if sys.platform == "win32" else "Courier"


def font(kind: str = "base") -> tuple:
    family = FAMILY or "TkDefaultFont"
    if kind == "fixed":
        return (FIXED_FAMILY, SIZES["fixed"])
    weight = "bold" if kind in ("section", "title", "bold") else "normal"
    return (family, SIZES.get(kind, SIZES["base"]), weight)


def enable_dpi_awareness() -> None:
    """Windows: render at the screen's real resolution instead of a blurry bitmap scale. Call before Tk()."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def color_title_bar(window: tk.Misc) -> None:
    """Windows 11 (build 22000+): NAVY title bar with white caption text. Older Windows: no effect."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id()) or window.winfo_id()

        def colorref(hex_color: str) -> ctypes.c_int:
            r, g, b = int(hex_color[1:3], 16), int(hex_color[3:5], 16), int(hex_color[5:7], 16)
            return ctypes.c_int(r | (g << 8) | (b << 16))

        for attr, colour in ((35, NAVY), (36, "#FFFFFF")):           # DWMWA_CAPTION_COLOR, DWMWA_TEXT_COLOR
            value = colorref(colour)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(value), ctypes.sizeof(value))
    except Exception:
        pass


def _named_fonts(root: tk.Misc) -> None:
    global FAMILY
    if not FAMILY:
        FAMILY = tkfont.nametofont("TkDefaultFont", root=root).actual("family")
    for name, kind in (("TkDefaultFont", "base"), ("TkTextFont", "base"), ("TkMenuFont", "base"),
                       ("TkHeadingFont", "bold"), ("TkTooltipFont", "small"), ("TkCaptionFont", "section")):
        try:
            f = tkfont.nametofont(name, root=root)
            fam, size, weight = font(kind)
            f.configure(family=fam, size=size, weight=weight)
        except tk.TclError:
            pass
    try:
        tkfont.nametofont("TkFixedFont", root=root).configure(family=FIXED_FAMILY, size=SIZES["fixed"])
    except tk.TclError:
        pass
    root.option_add("*TCombobox*Listbox.font", font("base"))
    root.option_add("*TCombobox*Listbox.selectBackground", LIGHT_ORANGE)
    root.option_add("*TCombobox*Listbox.selectForeground", TEXT)
    root.option_add("*Listbox.font", font("base"))
    root.option_add("*Menu.font", font("base"))


def _button(style: ttk.Style, name: str, bg: str, fg: str, hover: str, border: str, bold: bool = False,
            padding=(14, 7)) -> None:
    style.configure(name, background=bg, foreground=fg, bordercolor=border, lightcolor=bg, darkcolor=bg,
                    focuscolor=bg, focusthickness=0, relief="flat", padding=padding,
                    font=font("bold" if bold else "base"), anchor="center")
    style.map(name,
              background=[("disabled", DISABLED_BG), ("pressed", hover), ("active", hover)],
              lightcolor=[("disabled", DISABLED_BG), ("pressed", hover), ("active", hover)],
              darkcolor=[("disabled", DISABLED_BG), ("pressed", hover), ("active", hover)],
              foreground=[("disabled", DISABLED_FG)],
              bordercolor=[("disabled", DISABLED_BG), ("focus", hover)])


def _checkbox_image(root: tk.Misc, size: int, gap: int, fill: str, outline: str, tick: str | None) -> tk.PhotoImage:
    """A square checkbox drawn pixel by pixel (no image files): `fill` box, 1px `outline`, optional `tick` mark.
    The image is `gap` px wider than the box (transparent) so the label does not touch it."""
    img = tk.PhotoImage(master=root, width=size + gap, height=size)
    img.put(outline, to=(0, 0, size, size))
    img.put(fill, to=(1, 1, size - 1, size - 1))
    if tick:
        pts = [(0.22, 0.52), (0.42, 0.72), (0.78, 0.30)]          # short stroke down, long stroke up
        pts = [(x * size, y * size) for x, y in pts]
        half = max(1, round(size / 14))                            # stroke thickness ~ size / 7
        for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
            steps = int(max(abs(x1 - x0), abs(y1 - y0)) * 2) + 1
            for i in range(steps + 1):
                x = round(x0 + (x1 - x0) * i / steps)
                y = round(y0 + (y1 - y0) * i / steps)
                img.put(tick, to=(max(1, x - half), max(1, y - half), min(size - 1, x + half + 1), min(size - 1, y + half + 1)))
    return img


def _tick_indicator(root: tk.Misc, style: ttk.Style) -> None:
    """Checkboxes: a GREEN box with a white tick when ticked (clam's default draws a cross)."""
    line = tkfont.Font(root=root, font=font("base")).metrics("linespace")
    size, gap = max(14, round(line * 0.8)), max(8, round(line * 0.4))
    images = {
        "off": _checkbox_image(root, size, gap, "white", GHOST_BORDER, None),
        "on": _checkbox_image(root, size, gap, GREEN, GREEN, "white"),
        "off_disabled": _checkbox_image(root, size, gap, FIELD_DISABLED, BORDER, None),
        "on_disabled": _checkbox_image(root, size, gap, DISABLED_BG, DISABLED_BG, "white"),
    }
    root._castellan_check_images = images        # keep references, or Tk drops the images
    try:
        style.element_create("Tick.indicator", "image", images["off"],
                             ("disabled", "selected", images["on_disabled"]), ("disabled", images["off_disabled"]),
                             ("selected", images["on"]), sticky="")
    except tk.TclError:
        pass                                      # already created in this interpreter
    style.layout("TCheckbutton", [("Checkbutton.padding", {"sticky": "nswe", "children": [
        ("Tick.indicator", {"side": "left", "sticky": ""}),
        ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [("Checkbutton.label", {"sticky": "nswe"})]}),
    ]})])


def apply_theme(root: tk.Misc) -> ttk.Style:
    _named_fonts(root)
    style = ttk.Style(root)
    style.theme_use("clam")                      # the native "vista" theme ignores colours
    try:
        root.configure(background=BG)
    except tk.TclError:
        pass

    style.configure(".", background=CARD, foreground=TEXT, font=font("base"), bordercolor=BORDER,
                    troughcolor=BG, focuscolor=BLUE)
    # frames and cards
    style.configure("TFrame", background=CARD)
    style.configure("App.TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD, bordercolor=BORDER, relief="solid", borderwidth=0)
    style.configure("Card.TLabelframe", background=CARD, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                    relief="solid", borderwidth=1, padding=14)
    style.configure("Card.TLabelframe.Label", background=CARD, foreground=NAVY, font=font("section"))
    style.configure("TSeparator", background=BORDER)
    # labels
    style.configure("TLabel", background=CARD, foreground=TEXT, font=font("base"))
    style.configure("Section.TLabel", background=CARD, foreground=NAVY, font=font("section"))
    style.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=font("small"))
    style.configure("FieldHint.TLabel", background="white", foreground=MUTED, font=font("small"))
    style.configure("Empty.TLabel", background=CARD, foreground=MUTED, font=font("base"))
    style.configure("Status.TLabel", background=BG, foreground=TEXT, font=font("base"))
    style.configure("StatusOk.TLabel", background=BG, foreground=GREEN, font=font("bold"))
    style.configure("StatusError.TLabel", background=BG, foreground=RED, font=font("base"))
    style.configure("StatusMuted.TLabel", background=BG, foreground=MUTED, font=font("small"))
    style.configure("Dialog.TLabel", background=BG, foreground=TEXT, font=font("base"))
    style.configure("DialogTitle.TLabel", background=BG, foreground=NAVY, font=font("section"))
    style.configure("DialogMuted.TLabel", background=BG, foreground=MUTED, font=font("small"))
    # inputs: same padding everywhere so every field is the same height
    for name in ("TEntry", "TCombobox"):
        style.configure(name, fieldbackground="white", background="white", foreground=TEXT, bordercolor=BORDER,
                        lightcolor=BORDER, darkcolor=BORDER, padding=6, insertcolor=NAVY, arrowcolor=NAVY,
                        selectbackground=LIGHT_ORANGE, selectforeground=TEXT)
        style.map(name,
                  bordercolor=[("focus", BLUE)], lightcolor=[("focus", BLUE)], darkcolor=[("focus", BLUE)],
                  fieldbackground=[("disabled", FIELD_DISABLED), ("readonly", "white")],
                  foreground=[("disabled", DISABLED_FG)],
                  selectbackground=[("readonly", "white")], selectforeground=[("readonly", TEXT)],
                  arrowcolor=[("disabled", DISABLED_FG)], background=[("active", "white"), ("readonly", "white")])
    style.configure("TCheckbutton", background=CARD, foreground=TEXT, font=font("base"), padding=(0, 1),
                    indicatorbackground="white", indicatorforeground="white", indicatormargin=(0, 0, 8, 0),
                    upperbordercolor=BORDER, lowerbordercolor=BORDER, focuscolor=CARD)
    style.map("TCheckbutton",
              background=[("active", CARD)],
              indicatorbackground=[("disabled", FIELD_DISABLED), ("selected", ORANGE)],
              upperbordercolor=[("selected", ORANGE)], lowerbordercolor=[("selected", ORANGE)],
              foreground=[("disabled", DISABLED_FG)])
    _tick_indicator(root, style)
    # buttons
    _button(style, "Primary.TButton", ORANGE, "white", ORANGE_HOVER, ORANGE, bold=True, padding=(18, 8))
    _button(style, "Secondary.TButton", BLUE, "white", BLUE_HOVER, BLUE)
    _button(style, "Ghost.TButton", "white", BLUE, LIGHT_BLUE, GHOST_BORDER)
    style.configure("Ghost.TButton", lightcolor="white", darkcolor="white", bordercolor=GHOST_BORDER, borderwidth=1,
                    relief="solid")
    style.map("Ghost.TButton", foreground=[("disabled", DISABLED_FG)],
              lightcolor=[("active", LIGHT_BLUE)], darkcolor=[("active", LIGHT_BLUE)],
              bordercolor=[("active", BLUE), ("focus", BLUE)])
    style.configure("TButton", font=font("base"), padding=(12, 6))
    # table
    line = tkfont.Font(root=root, font=font("base")).metrics("linespace")
    style.configure("Treeview", background="white", fieldbackground="white", foreground=TEXT, font=font("base"),
                    rowheight=max(30, line + 12), bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
    style.map("Treeview", background=[("selected", LIGHT_ORANGE)], foreground=[("selected", TEXT)])
    style.configure("Treeview.Heading", background=NAVY, foreground="white", font=font("bold"), padding=6,
                    relief="flat", bordercolor=NAVY, lightcolor=NAVY, darkcolor=NAVY)
    style.map("Treeview.Heading", background=[("active", NAVY), ("pressed", NAVY)], foreground=[("active", "white")])
    style.configure("TScrollbar", background=DISABLED_BG, troughcolor=BG, bordercolor=BG, arrowcolor=NAVY,
                    lightcolor=DISABLED_BG, darkcolor=DISABLED_BG)
    style.map("TScrollbar", background=[("active", MUTED)])
    # progress
    style.configure("Orange.Horizontal.TProgressbar", background=ORANGE, troughcolor=LIGHT_BLUE, bordercolor=LIGHT_BLUE,
                    lightcolor=ORANGE, darkcolor=ORANGE, thickness=10)
    return style


def scale(root: tk.Misc) -> float:
    """Screen scaling factor (1.0 at 96 DPI, 1.5 at 150%). With DPI awareness on, Tk pixel sizes do not scale
    by themselves - multiply layout sizes (window, column minimums) by this."""
    try:
        return max(1.0, root.winfo_fpixels("1i") / 96.0)
    except tk.TclError:
        return 1.0


def style_listbox(lb: tk.Listbox) -> None:
    lb.configure(font=font("base"), background="white", foreground=TEXT, selectbackground=LIGHT_ORANGE,
                 selectforeground=TEXT, highlightthickness=1, highlightbackground=BORDER, highlightcolor=BLUE,
                 relief="flat", borderwidth=0, activestyle="none")


def card(parent: tk.Misc, title: str = "", padding: int = 14) -> tuple[tk.Frame, ttk.Frame]:
    """A white card with a 1px BORDER border. Returns (outer frame to grid, inner frame to fill).
    With a title, row 0 of the inner frame holds the section title (columnspan 3)."""
    outer = tk.Frame(parent, background=CARD, highlightthickness=1, highlightbackground=BORDER, highlightcolor=BORDER,
                     borderwidth=0)
    outer.columnconfigure(0, weight=1)
    outer.rowconfigure(0, weight=1)
    inner = ttk.Frame(outer, style="Card.TFrame", padding=padding)
    inner.grid(row=0, column=0, sticky="nsew")
    if title:
        ttk.Label(inner, text=title, style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))
    return outer, inner


def center_over(win: tk.Toplevel, parent: tk.Misc) -> None:
    win.update_idletasks()
    try:
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        w = max(win.winfo_width(), win.winfo_reqwidth())
        h = max(win.winfo_height(), win.winfo_reqheight())
        win.geometry(f"+{max(0, px + (pw - w) // 2)}+{max(0, py + (ph - h) // 2)}")
    except tk.TclError:
        pass
