"""Code-only checks for the AI-written sections (no LLM): numbers, banned phrases, length, markdown, "we".

    allowed_numbers(*texts)            every number in the facts / sources (+ $MM / $K forms of money)
    verify_plan(plan, allowed)         -> (approved sentences, [(dropped sentence, unknown numbers)])
    check_text(text, playbook, allowed, opening) -> [problem, ...]   (empty list = passes)
"""
from __future__ import annotations

import re
from typing import Iterable, Optional

NUM = re.compile(r"(?<![\w.])\d[\d,]*(?:\.\d+)?")
MIN_CHECKED = 10            # single-digit integers ("two", "3 tenants", "Q1") are not checked


def norm_num(s: str) -> str:
    s = s.replace(",", "")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s or "0"


def numbers(text: str) -> list[str]:
    return [norm_num(m.rstrip(",.")) for m in NUM.findall(text or "")]


def allowed_numbers(*texts: str) -> set[str]:
    out: set[str] = set()
    for t in texts:
        for n in numbers(t):
            out.add(n)
            try:
                v = float(n)
            except ValueError:
                continue
            if v >= 100_000:                 # $2,240,000 may be written $2.24MM / $2.2M / $2MM
                for dp in (0, 1, 2):
                    out.add(norm_num(f"{v / 1e6:.{dp}f}"))
            if v >= 1_000:
                for dp in (0, 1):
                    out.add(norm_num(f"{v / 1e3:.{dp}f}"))
            if 0 < v < 1:                     # ratios 0.392 -> 39%
                out.add(norm_num(f"{v * 100:.0f}"))
    return out


def unknown_numbers(text: str, allowed: set[str]) -> list[str]:
    bad = []
    for n in numbers(text):
        try:
            if float(n) < MIN_CHECKED and "." not in n:
                continue
        except ValueError:
            continue
        if n not in allowed and n not in bad:
            bad.append(n)
    return bad


def verify_plan(plan, allowed: set[str]):
    approved, dropped = [], []
    for s in plan.sentences:
        bad = unknown_numbers(" ".join(s.facts_used), allowed)
        (dropped.append((s, bad)) if bad else approved.append(s))
    return approved, dropped


def _strip_allowed(text: str, attributions: Iterable[str]) -> str:
    out = text
    for a in attributions or []:
        out = re.sub(re.escape(a), " ", out, flags=re.I)
    return out


def check_text(text: str, playbook: dict, allowed: set[str], opening: Optional[str] = None) -> list[str]:
    problems: list[str] = []
    body = _strip_allowed(text, playbook.get("allowed_attributions", []))
    low = body.lower()
    for p in playbook.get("banned_phrases", []) or []:
        if p.lower() in low:
            problems.append(f'Remove the phrase "{p}" (rephrase without it).')
    tw = playbook.get("target_words") or {}
    words = len(text.split())
    if tw.get("min") and words < int(tw["min"] * 0.7):
        problems.append(f"Too short: {words} words; write about {tw['min']}-{tw.get('max', tw['min'])} words.")
    if tw.get("max") and words > int(tw["max"] * 1.3):
        problems.append(f"Too long: {words} words; cut to about {tw.get('min', tw['max'])}-{tw['max']} words.")
    bad = unknown_numbers(text, allowed)
    if bad:
        problems.append("These numbers are not in the facts - remove them or use the exact figure from the facts: "
                        + ", ".join(bad) + ".")
    if re.search(r"\*\*|__|`|^\s*#|^\s*[-*•]\s", text, re.M):
        problems.append("Remove markdown (bold markers, headings, bullet points).")
    we = len(re.findall(r"\bwe\b", text, re.I)) - len(re.findall(r"\bwe believe\b", text, re.I))
    if we > 0:
        problems.append('Do not write "we" (only the fixed phrase "We believe" is allowed); write "Castellan".')
    if opening and not text.startswith(opening):
        problems.append(f'Begin exactly with: "{opening}"')
    return problems
