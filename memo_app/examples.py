"""Few-shot examples: tagged past-memo sections, picked by similarity to the deal.

Files: resources/style/examples/<loan_type>/<section>/*.txt with a YAML front-matter header

    ---
    loan_type: acquisition
    property_type: industrial
    structure: senior            # senior | pari_passu | participation | mezz | note_on_note
    exit: refinance              # refinance | sale | sellout
    quality: gold                # gold (hand-picked) | silver (harvested, unreviewed)
    source_memo: Memo - 150 E Aurora St.docx
    ---
    <section text>

The older "# key: value" header lines are still understood. Missing loan_type = the folder name.
Score (same section only): +3 loan_type, +3 structure, +2 property_type, +1 exit, +2 gold -> top k.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import yaml

WEIGHTS = {"loan_type": 3, "structure": 3, "property_type": 2, "exit": 1}
GOLD_BONUS = 2


@dataclass
class Example:
    path: Path
    section: str
    text: str
    tags: dict = field(default_factory=dict)

    @property
    def label(self) -> str:
        t = self.tags
        bits = [t.get("loan_type"), t.get("structure"), t.get("property_type")]
        return " / ".join(str(b) for b in bits if b) + f" - {self.path.stem}"


def _val(key: str, v) -> str:
    """Tag values are matched case-insensitively; the source memo keeps its file name."""
    s = "" if v is None else str(v).strip()
    return s if key in ("source_memo", "source") else s.lower()


def parse_example(raw: str) -> tuple[dict, str]:
    """(tags, text) from a front-matter or legacy '# key: value' example file."""
    raw = raw.lstrip("﻿")
    m = re.match(r"^---\s*\n(.*?)\n---\s*\n?", raw, re.S)
    if m:
        try:
            tags = yaml.safe_load(m.group(1)) or {}
        except yaml.YAMLError:
            tags = {}
        return {str(k).lower(): _val(str(k).lower(), v) for k, v in tags.items()}, raw[m.end():].strip()
    tags, lines = {}, raw.splitlines()
    while lines and lines[0].startswith("#"):
        mm = re.match(r"#\s*(\w+)\s*:\s*(.+)", lines.pop(0))
        if mm:
            tags[mm.group(1).lower()] = _val(mm.group(1).lower(), mm.group(2))
    if "source" in tags and "source_memo" not in tags:
        tags["source_memo"] = tags.pop("source")
    return tags, "\n".join(lines).strip()


def front_matter(tags: dict) -> str:
    keys = ["loan_type", "property_type", "structure", "exit", "quality", "source_memo"]
    lines = [f"{k}: {tags[k]}" if tags.get(k) else f"{k}:" for k in keys]
    return "---\n" + "\n".join(lines) + "\n---\n"


def detect_structure(text: str) -> str:
    t = (text or "").lower()
    if "pari passu" in t or "pari-passu" in t:
        return "pari_passu"
    if "participation" in t:
        return "participation"
    if "note-on-note" in t or "note on note" in t:
        return "note_on_note"
    if "junior loan" in t or "mezz" in t or "subordinate loan" in t:
        return "mezz"
    return "senior"


def detect_exit(text: str) -> str:
    t = (text or "").lower()
    if re.search(r"sell-?out|sale of the (individual )?units|unit sales|condo(minium)? sales", t):
        return "sellout"
    if "refinanc" in t:
        return "refinance"
    if re.search(r"\bsale\b|\bsell\b|\bsold\b", t):
        return "sale"
    return ""


def load_examples(root: Path, section: str) -> list[Example]:
    out = []
    if not root.exists():
        return out
    for lt_dir in sorted(p for p in root.iterdir() if p.is_dir()):
        sec_dir = lt_dir / section
        if not sec_dir.is_dir():
            continue
        for fp in sorted(sec_dir.glob("*.txt")):
            tags, text = parse_example(fp.read_text(encoding="utf-8"))
            if not text:
                continue
            tags.setdefault("loan_type", lt_dir.name)
            if not tags.get("loan_type"):
                tags["loan_type"] = lt_dir.name
            out.append(Example(fp, section, text, tags))
    return out


def score(ex: Example, deal: dict) -> int:
    s = sum(w for k, w in WEIGHTS.items() if deal.get(k) and ex.tags.get(k) == str(deal[k]).lower())
    return s + (GOLD_BONUS if ex.tags.get("quality") == "gold" else 0)


def select(examples: list[Example], deal: dict, k: int) -> list[Example]:
    """Top-k by score; ties keep file order (stable)."""
    ranked = sorted(enumerate(examples), key=lambda ie: (-score(ie[1], deal), ie[0]))
    return [ex for _, ex in ranked[:k]]


def deal_tags(f) -> dict:
    return {"loan_type": f.loan_type, "structure": getattr(f, "structure", None),
            "property_type": f.property_type, "exit": detect_exit(f.exit_strategy or "") or None}


def example_exit(text: str, explicit: Optional[str] = None) -> str:
    return explicit or detect_exit(text)
