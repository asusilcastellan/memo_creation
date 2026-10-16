"""Turn source files into text (and page images when a PDF is scanned)."""
from __future__ import annotations

import base64
import re
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf  # PyMuPDF


@dataclass
class DocText:
    name: str
    pages: list[str] = field(default_factory=list)     # one string per page / sheet
    images_b64: list[str] = field(default_factory=list)  # PNG pages for scanned PDFs
    note: str = ""

    @property
    def text(self) -> str:
        return "\n\n".join(f"--- page {i + 1} ---\n{p}" for i, p in enumerate(self.pages) if p.strip())

    @property
    def char_count(self) -> int:
        return sum(len(p) for p in self.pages)


# --------------------------------------------------------------------------- PDF
def pdf_to_doc(path: Path, name: str, scanned_min_chars: int = 80, max_image_pages: int = 12,
               dpi: int = 110) -> DocText:
    doc = pymupdf.open(path)
    d = DocText(name)
    scanned_pages = []
    for i, page in enumerate(doc):
        t = page.get_text("text") or ""
        d.pages.append(t)
        if len(t.strip()) < scanned_min_chars:
            scanned_pages.append(i)
    if scanned_pages:
        if len(scanned_pages) == len(d.pages):
            d.note = "scanned PDF (no text layer) - sent as images"
        for i in scanned_pages[:max_image_pages]:
            d.images_b64.append(render_page_b64(doc, i, dpi))
    doc.close()
    return d


def render_page_b64(doc_or_path, index: int, dpi: int = 110) -> str:
    doc = pymupdf.open(doc_or_path) if not isinstance(doc_or_path, pymupdf.Document) else doc_or_path
    pix = doc[index].get_pixmap(dpi=dpi)
    return base64.b64encode(pix.tobytes("png")).decode()


def last_pages_b64(path: Path, n: int = 2, dpi: int = 110) -> list[str]:
    doc = pymupdf.open(path)
    idx = list(range(max(0, len(doc) - n), len(doc)))
    out = [render_page_b64(doc, i, dpi) for i in idx]
    doc.close()
    return out


def pdf_signature_hints(path: Path) -> list[str]:
    """Cheap, model-free signals that a PDF was signed (DocuSign, signature fields, /s/)."""
    hints = []
    try:
        doc = pymupdf.open(path)
        for page in doc:
            for w in page.widgets() or []:
                if w.field_type == pymupdf.PDF_WIDGET_TYPE_SIGNATURE:
                    hints.append("PDF signature field")
                    break
        tail = "".join(doc[i].get_text() for i in range(max(0, len(doc) - 3), len(doc)))
        head = doc[0].get_text() if len(doc) else ""
        if re.search(r"docusign envelope id|dotloop|adobe sign|hellosign|signnow", head + tail, re.I):
            hints.append("e-signature stamp")
        if re.search(r"/s/\s*\w", tail):
            hints.append("/s/ conformed signature")
        doc.close()
    except Exception:
        pass
    return hints


# --------------------------------------------------------------------------- DOCX
def docx_to_doc(path: Path, name: str) -> DocText:
    import docx

    d = docx.Document(str(path))
    lines = [p.text for p in d.paragraphs if p.text.strip()]
    for t in d.tables:
        for row in t.rows:
            cells = []
            for c in row.cells:
                txt = c.text.strip()
                if txt and (not cells or cells[-1] != txt):
                    cells.append(txt)
            if cells:
                lines.append(" | ".join(cells))
    return DocText(name, pages=["\n".join(lines)])


# --------------------------------------------------------------------------- XLSX
def _clean_format(section: str) -> str:
    """Drop the parts of an Excel number format that do not print: [colors]/[$-409], quotes, _x padding, *x fill."""
    s = re.sub(r"\[[^\]]*\]", "", section)
    s = re.sub(r'"([^"]*)"', r"\1", s)
    s = re.sub(r"_.", "", s)
    s = re.sub(r"\*.", "", s)
    return s.replace("\\", "").strip()


def format_cell(value, number_format: str | None):
    """The cell as Excel displays it: 869.065868 with '$#,##0' -> '$869'; 0.4895 with '0.00%' -> '48.95%'.
    Handles $ / accounting / #,##0 / 0.00 / percent / date formats; anything else -> the raw value."""
    import datetime as _dt

    if value is None or isinstance(value, (bool, str)):
        return value
    fmt = (number_format or "General").strip()
    if isinstance(value, (_dt.datetime, _dt.date)):
        low = fmt.lower()
        if "mmmm" in low:
            return f"{value:%B} {value.day}, {value.year}"
        if "mmm" in low:
            return f"{value:%b}-{value:%y}" if "yyyy" not in low else f"{value:%b} {value.day}, {value.year}"
        return f"{value.month}/{value.day}/{value.year}"
    if not isinstance(value, (int, float)):
        return value
    if fmt.lower() == "general" or fmt == "@":
        return int(value) if float(value).is_integer() else round(value, 6)
    sections = fmt.split(";")
    if value == 0 and len(sections) >= 3:
        zero = _clean_format(sections[2]).replace("?", "").replace("#", "").strip()
        if zero and not re.search(r"\d", zero):
            return zero
    use_neg = value < 0 and len(sections) >= 2
    sec = _clean_format(sections[1] if use_neg else sections[0])
    m = re.search(r"[0#?][0#?,]*(?:\.([0#?]+))?", sec)
    if not m:
        return int(value) if float(value).is_integer() else round(value, 6)
    decimals = len(m.group(1) or "")
    v = abs(value) if use_neg else value
    if "%" in sec:
        v = v * 100
    num = f"{v:,.{decimals}f}" if "," in m.group(0) else f"{v:.{decimals}f}"
    if num.startswith("-") and "$" in sec[:m.start()]:
        return "-" + sec[:m.start()] + num[1:] + sec[m.end():]
    return (sec[:m.start()] + num + sec[m.end():]).strip()


def xlsx_to_doc(path: Path, name: str, max_rows: int = 400, max_cols: int = 30,
                sheets: list[str] | None = None) -> DocText:
    """Dump every sheet (or only `sheets`) as 'R<row>: B=label | C=value' lines: cached values (not formulas),
    each rendered with its number format so the text matches what Excel shows ($869, 48.95%, Dec-25)."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    wb = load_workbook(path, data_only=True, read_only=True)
    d = DocText(name)
    for ws in wb.worksheets:
        if sheets and ws.title not in sheets:
            continue
        lines = [f"### Sheet: {ws.title}"]
        for r_idx, row in enumerate(ws.iter_rows(max_row=max_rows, max_col=max_cols), start=1):
            cells = []
            for c_idx, cell in enumerate(row, start=1):
                v = getattr(cell, "value", None)
                if v is None or (isinstance(v, str) and not v.strip()):
                    continue
                v = format_cell(v, getattr(cell, "number_format", "General"))
                cells.append(f"{get_column_letter(c_idx)}={v}")
            if cells:
                lines.append(f"R{r_idx}: " + " | ".join(cells))
        if len(lines) > 1:
            d.pages.append("\n".join(lines))
    wb.close()
    return d


def xlsx_first_image(path: Path) -> bytes | None:
    """The model's 'Model' tab usually carries a property photo; grab the first embedded image."""
    import zipfile

    try:
        with zipfile.ZipFile(path) as z:
            media = sorted(n for n in z.namelist() if n.startswith("xl/media/")
                           and n.lower().endswith((".png", ".jpg", ".jpeg")))
            if media:
                biggest = max(media, key=lambda n: z.getinfo(n).file_size)
                return z.read(biggest)
    except Exception:
        return None
    return None


# --------------------------------------------------------------------------- dispatcher
def load_doc(path: Path, name: str, cfg: dict, sheets: list[str] | None = None) -> DocText:
    suf = path.suffix.lower()
    if suf == ".pdf":
        return pdf_to_doc(path, name, cfg.get("scanned_page_min_chars", 80),
                          cfg.get("max_image_pages", 12), cfg.get("image_dpi", 110))
    if suf in (".docx",):
        return docx_to_doc(path, name)
    if suf in (".xlsx", ".xlsm"):
        return xlsx_to_doc(path, name, sheets=sheets)
    if suf in (".png", ".jpg", ".jpeg"):
        return DocText(name, images_b64=[base64.b64encode(path.read_bytes()).decode()])
    return DocText(name, note=f"unsupported file type {suf}")


def trim_to_relevant(doc: DocText, keywords: list[str], max_chars: int, always_first: int = 4) -> str:
    """Keep the first pages (cover / summary) plus the pages that score highest on keywords."""
    if doc.char_count <= max_chars:
        return doc.text
    kws = [k.lower() for k in keywords]
    scored = []
    for i, p in enumerate(doc.pages):
        low = p.lower()
        scored.append((sum(low.count(k) for k in kws), i))
    keep = set(range(min(always_first, len(doc.pages))))
    budget = max_chars - sum(len(doc.pages[i]) for i in keep)
    for score, i in sorted(scored, reverse=True):
        if score == 0 or i in keep:
            continue
        if len(doc.pages[i]) > budget:
            continue
        keep.add(i)
        budget -= len(doc.pages[i])
    return "\n\n".join(f"--- page {i + 1} ---\n{doc.pages[i]}" for i in sorted(keep))
