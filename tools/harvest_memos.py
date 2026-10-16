"""Build the few-shot style library from past memos.

Every deal folder in Current Portfolio has Memo/Memo - <address>.docx. This tool splits each
memo into sections and saves them as examples the writer model imitates:

    resources/style/examples/<loan_type>/<section>/<deal>.txt

Usage (from the project root, venv active):
    # from a local / synced folder that contains many deal folders
    python tools/harvest_memos.py --root "C:/Users/me/Castellan Dropbox/Sharepoint/Bridge Lending/Current Portfolio"
    # straight from Dropbox
    python tools/harvest_memos.py --root "/Bridge Lending/Current Portfolio" --dropbox
    # individual memo files with a known loan type
    python tools/harvest_memos.py --files memo1.docx memo2.docx --loan-type construction

Review the output folder afterwards: delete weak examples, keep the ones the team is proud of, and change
`quality: silver` to `quality: gold` for the best ones (gold examples are preferred by the writer).
Each file gets a YAML front-matter header: loan_type, property_type, structure (pari passu / participation /
junior loan -> mezz / senior), exit (refinance / sale / sellout), quality, source_memo.
Loan type is detected from the Transaction Overview; pass --loan-type to override.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from memo_app import paths  # noqa: E402
from memo_app.examples import detect_exit, detect_structure, front_matter  # noqa: E402

HEADINGS = {
    "transactionoverview": "transaction_overview",
    "sponsorshipoverview": "sponsorship",
    "submarketoverview": "submarket_overview",
    "marketoverview": "submarket_overview",
    "businessplan": "business_plan",
    "risksmitigants": "risks", "risksandmitigants": "risks",
    "loanfinancials": "loan_financials",
    "investmenthighlights": "highlights",
    "pictures": "end", "photo": "end", "photos": "end", "map": "end",
}
PROPERTY_WORDS = ["multifamily", "industrial", "retail", "office", "single-family", "condo", "land",
                  "hotel", "hospitality", "mixed-use", "self storage", "warehouse", "residential"]


def norm(s: str) -> str:
    return re.sub(r"[^a-z]", "", s.lower())


def body_paragraph_texts(path: Path) -> list[str]:
    """Paragraph texts in document order, including text inside table cells."""
    d = docx.Document(str(path))
    out = []
    for el in d.element.body.iterchildren():
        if el.tag == qn("w:p"):
            out.append("".join(t.text or "" for t in el.iter(qn("w:t"))))
        elif el.tag == qn("w:tbl"):
            # merged cells are stored once in the XML, so document order has no duplicates
            for p in el.iter(qn("w:p")):
                out.append("".join(t.text or "" for t in p.iter(qn("w:t"))))
    return [t.strip() for t in out]


def split_sections(texts: list[str]) -> dict[str, list[str]]:
    sections: dict[str, list[str]] = {}
    cur = None
    for t in texts:
        key = HEADINGS.get(norm(t))
        if key:
            cur = key
            sections.setdefault(cur, [])
            continue
        if cur and t and cur != "end":
            sections[cur].append(t)
    return sections


def sponsorship_parts(lines: list[str]) -> dict[str, str]:
    parts, cur = {}, None
    labels = {"businessbackground": "business_background", "personalfinancials": "personal_financials",
              "backgroundcreditcheckresults": "background"}
    for ln in lines:
        m = re.match(r"^\s*([A-Za-z &]+?)\s*:\s*(.*)$", ln)
        if m and norm(m.group(1)) in labels:
            cur = labels[norm(m.group(1))]
            parts[cur] = m.group(2).strip()
        elif cur and ln:
            parts[cur] = (parts[cur] + " " + ln).strip()
    return parts


def sponsor_experience(risk_lines: list[str]) -> str:
    text = " ".join(risk_lines)
    m = re.search(r"Sponsor Risk\s*:\s*(.*?)(?=\b[A-Z][A-Za-z /]+ Risk\s*:|$)", text)
    if not m:
        return ""
    sents = re.split(r"(?<=[.!?])\s+(?=[A-Z])", m.group(1))
    keep = [s for s in sents if not re.search(r"background|guarant|previously worked|first loan|bankruptc|carve", s, re.I)]
    return " ".join(keep).strip()


def detect_loan_type(to_text: str) -> str:
    t = to_text.lower()
    if "note-on-note" in t or "note on note" in t or "assignment of the underlying loan" in t:
        return "note_on_note"
    if "construction of" in t or "finance the construction" in t:
        return "construction"
    if "refinanc" in t.split(".")[0]:
        return "refinance"
    return "acquisition"


def detect_property_type(text: str) -> str:
    t = text.lower()
    for w in PROPERTY_WORDS:
        if w in t:
            return {"warehouse": "industrial", "hotel": "hospitality", "residential": "single-family",
                    "condo": "condominium", "self storage": "self-storage"}.get(w, w)
    return ""


def harvest(memo: Path, deal_name: str, out_root: Path, loan_type: str | None) -> list[Path]:
    secs = split_sections(body_paragraph_texts(memo))
    to = " ".join(secs.get("transaction_overview", []))
    if not to:
        print(f"  skip (no Transaction Overview found): {memo}")
        return []
    lt = loan_type or detect_loan_type(to)
    pt = detect_property_type(to)
    plan_text = " ".join(secs.get("business_plan", []))
    tags = {"loan_type": lt, "property_type": pt, "structure": detect_structure(to),
            "exit": detect_exit(to + " " + plan_text), "quality": "silver", "source_memo": memo.name}
    items = {
        "transaction_overview": "\n\n".join(secs.get("transaction_overview", [])),
        "submarket_overview": "\n\n".join(secs.get("submarket_overview", [])),
        "business_plan": "\n\n".join(secs.get("business_plan", [])),
        "sponsor_experience": sponsor_experience(secs.get("risks", [])),
    }
    sp = sponsorship_parts(secs.get("sponsorship", []))
    items["business_background"] = sp.get("business_background", "")
    written = []
    slug = re.sub(r"[^A-Za-z0-9]+", "_", deal_name).strip("_")[:60] or memo.stem
    for sec, text in items.items():
        text = text.strip()
        if len(text) < 40:
            continue
        d = out_root / lt / sec
        d.mkdir(parents=True, exist_ok=True)
        p = d / f"{slug}.txt"
        p.write_text(front_matter(tags) + text + "\n", encoding="utf-8")
        written.append(p)
    print(f"  {lt:13s} {pt:13s} {len(written)} sections  <- {memo}")
    return written


def find_memos_local(root: Path) -> list[tuple[str, Path]]:
    found = []
    for deal in sorted(p for p in root.iterdir() if p.is_dir()):
        memo_dir = next((d for d in deal.iterdir() if d.is_dir() and norm(d.name) in ("memo", "memos")), None)
        if not memo_dir:
            continue
        docs = [f for f in memo_dir.glob("*.docx") if not f.name.startswith("~$") and "memo" in f.name.lower()]
        if docs:
            found.append((deal.name, max(docs, key=lambda f: f.stat().st_mtime)))
    return found


def find_memos_dropbox(root: str) -> list[tuple[str, Path]]:
    from memo_app.settings import load_settings
    from memo_app.storage import DropboxStorage

    st = DropboxStorage(load_settings())
    entries = st.walk(root)
    by_deal: dict[str, list] = {}
    for e in entries:
        parts = e.rel.split("/")
        if e.is_dir or len(parts) != 3 or norm(parts[1]) not in ("memo", "memos"):
            continue
        if e.suffix == ".docx" and "memo" in e.name.lower() and not e.name.startswith("~$"):
            by_deal.setdefault(parts[0], []).append(e)
    return [(deal, st.fetch(max(es, key=lambda e: e.modified))) for deal, es in sorted(by_deal.items())]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", help="Folder containing many deal folders (e.g. Current Portfolio)")
    ap.add_argument("--dropbox", action="store_true", help="--root is a Dropbox API path")
    ap.add_argument("--files", nargs="*", help="Individual memo .docx files")
    ap.add_argument("--loan-type", choices=["acquisition", "construction", "refinance", "note_on_note"])
    ap.add_argument("--out", default=str(paths.resource("style", "examples")))
    a = ap.parse_args()

    out = Path(a.out)
    memos: list[tuple[str, Path]] = []
    if a.files:
        memos += [(Path(f).stem.replace("Memo - ", ""), Path(f)) for f in a.files]
    if a.root:
        memos += find_memos_dropbox(a.root) if a.dropbox else find_memos_local(Path(a.root))
    if not memos:
        ap.error("nothing to harvest - use --root or --files")
    total = 0
    for deal, path in memos:
        try:
            total += len(harvest(path, deal, out, a.loan_type))
        except Exception as exc:
            print(f"  ERROR {path}: {exc}")
    print(f"Wrote {total} example files under {out}")


if __name__ == "__main__":
    main()
