"""Render boilerplate.yaml with the deal facts.

Every value printed with {{ }} is wrapped in invisible markers:
    ⟦value⟧          filled in by the tool       -> normal text in the memo
    ⟬text⟭           written by the AI model     -> normal text in the memo
    ⟪⁣uid⁣⟫ red remark linked to a facts.Flag -> "[REVIEW R3: title]" (red highlight, white bold)
The docx builder numbers the linked flags R1, R2, ... in document order; the review workbook lists the same
Flag objects, so the IDs always match. Free-text ⟪[text]⟫ remarks are still understood (the builder turns
them into flags).

Inside templates:
    {{ f.remark('background_adverse') }}   remark linked to the flag raised in facts.build_facts
    {{ 'REVIEW: some text'|flag }}         ad-hoc remark (creates a flag in the current section)
    a field that is None                   -> [MISSING Rn: <plain field name>] flag (facts.FIELD_LABELS)
"""
from __future__ import annotations

import contextvars
import re

from jinja2 import Environment, Undefined

VAR_O, VAR_C = "⟦", "⟧"
FLAG_O, FLAG_C = "⟪", "⟫"
AI_O, AI_C = "⟬", "⟭"
UID = "⁣"
MARKERS = (VAR_O, VAR_C, FLAG_O, FLAG_C, AI_O, AI_C)
UID_TOKEN = re.compile(UID + r"(\d+)" + UID)
REMARK_MAX_WORDS = 15

# (facts, memo section) for the template being rendered - lets |flag and [MISSING] create linked flags
_ctx: contextvars.ContextVar[tuple | None] = contextvars.ContextVar("render_ctx", default=None)


def short_remark(text: str, max_words: int = REMARK_MAX_WORDS) -> str:
    words = str(text).split()
    return str(text) if len(words) <= max_words else " ".join(words[:max_words]).rstrip(",;:") + " …"


_LEVEL = re.compile(r"^\s*(REVIEW|MISSING|CHECK)\b\s*:?\s*", re.I)


def flag(text: str) -> str:
    """Ad-hoc red remark. Inside render() it becomes a linked flag; outside it is a plain red remark."""
    ctx = _ctx.get()
    if ctx and ctx[0] is not None and hasattr(ctx[0], "new_remark"):
        f, section = ctx
        m = _LEVEL.match(text)
        level = m.group(1).lower() if m else "review"
        title = text[m.end():] if m else text
        title = re.sub(r"\s*-\s*see Review Notes\s*$", "", title, flags=re.I)
        return f.new_remark(level, short_remark(title), text, section)
    return f"{FLAG_O}[{short_remark(text)}]{FLAG_C}"


def ai(text: str) -> str:
    return f"{AI_O}{text}{AI_C}" if text else ""


def strip_markers(text: str, facts=None) -> str:
    """Plain text. Linked remarks are removed, or printed as '[REVIEW R3: title]' when `facts` is given."""
    t = text or ""
    if facts is not None:
        by_uid = {fl.uid: fl for fl in facts.flags}
        t = UID_TOKEN.sub(lambda m: by_uid[int(m.group(1))].remark_text() if int(m.group(1)) in by_uid else "", t)
    else:
        t = re.sub(FLAG_O + UID_TOKEN.pattern + FLAG_C, "", t)
    return re.sub("[" + "".join(MARKERS) + "]", "", t)


# ---------------------------------------------------------------- missing values
class Missing:
    """A None field of the facts, remembering its name so the [MISSING] remark can say what is missing."""
    __slots__ = ("name",)

    def __init__(self, name: str):
        self.name = name

    def __bool__(self):
        return False

    def __str__(self):
        return ""

    def __eq__(self, other):
        return other is None or isinstance(other, Missing)

    def __hash__(self):
        return hash(None)

    def __iter__(self):
        return iter(())

    def __len__(self):
        return 0


class FactsView:
    """Read-only view of DealFacts for templates: None attributes come back as Missing(name)."""

    def __init__(self, obj):
        object.__setattr__(self, "_obj", obj)

    def __getattr__(self, name):
        v = getattr(object.__getattribute__(self, "_obj"), name)
        if v is None:
            return Missing(name)
        if isinstance(v, list) and v and hasattr(v[0], "__dataclass_fields__"):
            return [FactsView(i) for i in v]
        return v


def _missing(v) -> bool:
    return v is None or isinstance(v, Missing) or v == ""


def _finalize(v):
    if isinstance(v, Undefined) or (isinstance(v, str) and v == ""):
        return ""
    if v is None or isinstance(v, Missing):
        ctx = _ctx.get()
        if ctx and ctx[0] is not None and hasattr(ctx[0], "missing_remark"):
            return ctx[0].missing_remark(v.name if isinstance(v, Missing) else "value", ctx[1])
        return f"{FLAG_O}[MISSING]{FLAG_C}"
    s = str(v)
    if any(m in s for m in MARKERS):
        return s
    return f"{VAR_O}{s}{VAR_C}"


# ---------------------------------------------------------------- number filters
def money(v):
    return v if _missing(v) else f"${float(v):,.0f}"


def mm(v):
    if _missing(v):
        return v
    v = float(v)
    return f"${v / 1e6:,.2f}MM" if abs(v) >= 1e6 else f"${v / 1e3:,.0f}K"


def pct(v):
    """Ratio -> whole percent (0.392 -> 39%)."""
    return v if _missing(v) else f"{float(v) * 100:.0f}%"


def pct2(v):
    """Percent number -> 2dp (9.64 -> 9.64%)."""
    return v if _missing(v) else f"{float(v):.2f}%"


def sf(v):
    return v if _missing(v) else f"{float(v):,.0f}"


def cap(v):
    return v[:1].upper() + v[1:] if isinstance(v, str) and v else v


def make_env() -> Environment:
    env = Environment(finalize=_finalize, autoescape=False, trim_blocks=True, lstrip_blocks=True)
    env.filters.update(money=money, mm=mm, pct=pct, pct2=pct2, sf=sf, cap=cap, flag=flag)
    return env


ENV = make_env()


def tidy(text: str) -> str:
    """Collapse the whitespace that folded YAML + Jinja tags leave behind."""
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\s+([.,;:])", r"\1", text)
    text = re.sub(r"\.\s*\.", ".", text)
    return text


def render(template: str, f, section: str = "") -> str:
    """Render one template. `section` is the memo section the text belongs to (for the flags it creates)."""
    token = _ctx.set((f, section))
    try:
        return tidy(ENV.from_string(template).render(f=FactsView(f) if f is not None else f))
    finally:
        _ctx.reset(token)


def when(expr: str, f) -> bool:
    return bool(ENV.compile_expression(str(expr), undefined_to_none=True)(f=FactsView(f) if f is not None else f))
