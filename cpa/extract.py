"""Step 1a: read text lines, thin lines, tables and pictures (with positions) from a PDF.

Run:
    python -m cpa.extract OLD NEW
Writes output/dump_<name>.txt (UTF-8) for each input.
"""
from __future__ import annotations

import argparse
import hashlib
import ctypes
import re

from dataclasses import dataclass, field
from pathlib import Path

import pdfplumber

from cpa.convert import get_password, to_pdf

BBox = tuple[float, float, float, float]  # x0, top, x1, bottom
MIN_IMAGE_PT = 6.0                         # smaller pictures (bullets, dots) are ignored
ITALIC_SKEW = 0.1
BOLD_FONT_RE = re.compile(r"(?i)bold|black|heavy|semibold|demibold|-W[6-9]$")
ITALIC_FONT_RE = re.compile(r"(?i)italic|oblique")

@dataclass
class Char:
    text: str
    x0: float
    top: float
    x1: float
    bottom: float
    font: str
    size: float
    color: tuple | None
    bold: bool = False
    italic: bool = False

@dataclass
class Line:
    page: int  # 1-based
    text: str
    x0: float
    top: float
    x1: float
    bottom: float
    chars: list[Char] = field(default_factory=list)

@dataclass
class PageInfo:
    number: int
    width: float
    height: float

@dataclass
class HSegment:
    """A thin horizontal shape: candidate strikethrough, underline or table border."""
    page: int
    x0: float
    x1: float
    y: float
    kind: str  # "rect" or "line"

@dataclass
class TableInfo:
    """A table grid found from its borders. rows[r][c] = cell bbox, or None (merged)."""
    page: int
    bbox: BBox
    rows: list[list[BBox | None]]

@dataclass
class ImageInfo:
    """A picture placed on the page (position in pt; digest = hash of the image data)."""
    page: int
    x0: float
    top: float
    x1: float
    bottom: float
    src_w: int
    src_h: int
    digest: str | None

@dataclass
class Document:
    path: Path
    pages: list[PageInfo]
    lines: list[Line]
    hsegments: list[HSegment]
    tables: list[TableInfo] = field(default_factory=list)
    images: list[ImageInfo] = field(default_factory=list)

def _color(c: dict) -> tuple | None:
    v = c.get("non_stroking_color")
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return (round(float(v), 3),)
    try:
        return tuple(round(float(x), 3) for x in v)
    except (TypeError, ValueError):
        return None

def _italic(c: dict) -> bool:
    """Italic from the font name or a slanted glyph (Word/LibreOffice fake italic)."""
    if ITALIC_FONT_RE.search(c.get("fontname", "")):
        return True
    m = c.get("matrix")
    if not m or len(m) < 4:
        return False
    a, b, cc, d = (float(x) for x in m[:4])
    return abs(b) < 1e-6 and abs(d) > 1e-6 and abs(cc / d) > ITALIC_SKEW

def _stroked_text(pdf_path, page_count: int) -> dict[int, list[BBox]]:
    """Areas of text drawn with fill + stroke = fake bold (Word/LibreOffice for Japanese fonts).

    pdfplumber cannot see the text render mode, so it is read with pdfium.
    Returns page -> [(x0, top, x1, bottom)] in pdfplumber coordinates.
    """
    import pypdfium2 as pdfium
    import pypdfium2.raw as pdfium_c

    stroked_modes = { pdfium_c.FPDF_TEXTRENDERMODE_FILL_STROKE,
                      pdfium_c.FPDF_TEXTRENDERMODE_FILL_STROKE_CLIP}
    out: dict[int, list[BBox]] = {}
    pdf = pdfium.PdfDocument(str(pdf_path), password=get_password(pdf_path))
    try:
        for pno in range(1, min(page_count, len(pdf))+1):
            page = pdf[pno -1]
            mx0, _, _, mtop = page.get_mediabox()
            for obj in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_TEXT], max_depth=15):
                if pdfium_c.FPDFTextObj_GetTextRenderMode(obj.raw) not in stroked_modes:
                    continue
                l, b, r, t = (ctypes.c_float() for _ in range(4))
                if pdfium_c.FPDFPageObj_GetBounds(obj.raw, l, b, r, t):
                    out.setdefault(pno, []).append(
                        (l.value - mx0, mtop -t.value, r.value - mx0, mtop - b.value))
            page.close()
    finally:
        pdf.close()
    return out

def _in_any(x: float, y: float, boxes: list[BBox]) -> bool:
    return any(b[0] - 0.5 <= x <= b[2] + 0.5 and b[1] - 0.5 <= y <= b[3] + 0.5 for b in boxes)

def _bbox(v) -> BBox:
    return (float(v[0]), float(v[1]), float(v[2]), float(v[3]))

def _digest(im: dict) -> str | None:
    s = im.get("stream")
    if s is None:
        return None
    for getter in ("get_rawdata", "get_data"):
        try:
            data = getattr(s, getter)()
        except Exception:
            continue
        if data:
            return hashlib.sha1(data).hexdigest()
    return None

def _logical_columns(rows: list[list[BBox | None]], min_w: float = 8.0,
                     snap: float = 1.0) -> list[list[BBox | None]]:
    """Re-index cells on the table's real columns (drops narrow shading/padding strips)."""
    bounds: list[list[float]] = []  # [x, number of rows that have this boundary]
    for r in rows:
        xs = sorted(x for c in r if c for x in (c[0], c[2]))
        seen: list[float] = []
        for x in xs:
            if not seen or x - seen[-1] > snap:
                seen.append(x)
        for x in seen:
            for b in bounds:
                if abs(b[0] - x) <= snap:
                    b[1] += 1
                    break
            else:
                bounds.append([x, 1])
    bounds.sort()

    while len(bounds) > 2:
        w, i = min((bounds[k + 1][0] - bounds[k][0], k) for k in range(len(bounds) - 1))
        if w >= min_w:
            break
        a, b = i, i + 1
        if a == 0:
            drop = b
        elif b == len(bounds) - 1:
            drop = a
        else:
            drop = a if bounds[a][1] < bounds[b][1] else b
        bounds.pop(drop)

    xs = [b[0] for b in bounds]
    if len(xs) < 2:
        return rows

    def nearest(x: float) -> int:
        return min(range(len(xs)), key=lambda k: abs(xs[k] - x))

    out: list[list[BBox | None]] = []
    for r in rows:
        nr: list[BBox | None] = [None] * (len(xs) - 1)
        for c in r:
            if not c:
                continue
            s, e = nearest(c[0]), nearest(c[2])
            if e <= s:
                continue
            box = (xs[s], c[1], xs[e], c[3])
            old = nr[s]
            nr[s] = box if old is None else (min(old[0], box[0]), min(old[1], box[1]),
                                             max(old[2], box[2]), max(old[3], box[3]))
        out.append(nr)
    return out

def extract(pdf_path: str | Path) -> Document:
    pages: list[PageInfo] = []
    lines: list[Line] = []
    segs: list[HSegment] = []
    tables: list[TableInfo] = []
    images: list[ImageInfo] = []

    with pdfplumber.open(pdf_path, password=get_password(pdf_path) or "") as pdf:
        stroked = _stroked_text(pdf_path, len(pdf.pages))
        for pno, page in enumerate(pdf.pages, start=1):
            pages.append(PageInfo(pno, float(page.width), float(page.height)))
            fake_bold = stroked.get(pno, [])

            for tl in page.extract_text_lines(strip=True, return_chars=True):
                chars = [
                    Char(
                        text=c["text"],
                        x0=float(c["x0"]), top=float(c["top"]),
                        x1=float(c["x1"]), bottom=float(c["bottom"]),
                        font=c.get("fontname", ""),
                        size=round(float(c.get("size", 0)), 2),
                        color=_color(c),
                        bold=bool(BOLD_FONT_RE.search(c.get("fontname", ""))) or (
                            bool(fake_bold) and _in_any((float(c["x0"]) + float(c["x1"])) /2, 
                                                        (float(c["top"]) + float(c["bottom"])) /2, 
                                                        fake_bold)),
                        italic=_italic(c),
                    )
                    for c in tl["chars"]
                ]
                lines.append(Line(pno, tl["text"], float(tl["x0"]), float(tl["top"]),
                                  float(tl["x1"]), float(tl["bottom"]), chars))

            for r in page.rects:
                h, w = r["bottom"] - r["top"], r["x1"] - r["x0"]
                if h <= 1.5 and w >= 2:
                    segs.append(HSegment(pno, float(r["x0"]), float(r["x1"]),
                                         float((r["top"] + r["bottom"]) / 2), "rect"))
            for ln in page.lines:
                if abs(ln["top"] - ln["bottom"]) <= 0.5 and (ln["x1"] - ln["x0"]) >= 2:
                    segs.append(HSegment(pno, float(ln["x0"]), float(ln["x1"]),
                                         float(ln["top"]), "line"))

            for t in page.find_tables():
                rows: list[list[BBox | None]] = []
                for r in t.rows:
                    cells = [_bbox(c) if c else None for c in r.cells]
                    real = [c for c in cells if c]
                    if not real or max(c[3] - c[1] for c in real) < 2:
                        continue
                    rows.append(cells)
                if rows:
                    tables.append(TableInfo(pno, _bbox(t.bbox), _logical_columns(rows)))

            for im in page.images:
                x0, x1 = sorted((float(im["x0"]), float(im["x1"])))
                top, bottom = sorted((float(im["top"]), float(im["bottom"])))
                x0, top = max(0.0, x0), max(0.0, top)
                x1, bottom = min(float(page.width), x1), min(float(page.height), bottom)
                if x1 - x0 < MIN_IMAGE_PT or bottom - top < MIN_IMAGE_PT:
                    continue
                src = im.get("srcsize") or (0, 0)
                images.append(ImageInfo(pno, x0, top, x1, bottom,
                                        int(src[0] or 0), int(src[1] or 0), _digest(im)))

    lines.sort(key=lambda l: (l.page, round(l.top, 1), l.x0))
    segs.sort(key=lambda s: (s.page, s.y, s.x0))
    tables.sort(key=lambda t: (t.page, t.bbox[1], t.bbox[0]))
    images.sort(key=lambda i: (i.page, round(i.top, 1), i.x0))
    return Document(Path(pdf_path), pages, lines, segs, tables, images)

def write_dump(doc: Document, path: Path) -> None:
    out: list[str] = [f"# {doc.path.name}"]
    for p in doc.pages:
        out.append(f"\n=== page {p.number}  ({p.width:.0f} x {p.height:.0f} pt)")
        for l in (l for l in doc.lines if l.page == p.number):
            sizes = sorted({c.size for c in l.chars})
            colors = sorted({str(c.color) for c in l.chars})
            style = "".join((" bold" if any(c.bold for c in l.chars) else "",
                             " italic" if any(c.italic for c in l.chars) else ""))
            out.append(f"[y={l.top:6.1f}-{l.bottom:6.1f} x={l.x0:6.1f}-{l.x1:6.1f} "
                       f"size={sizes} color={colors}{style}] {l.text}")
        out.append(f"--- thin horizontal shapes on page {p.number}")
        for s in (s for s in doc.hsegments if s.page == p.number):
            out.append(f"  y={s.y:6.1f} x={s.x0:6.1f}-{s.x1:6.1f} ({s.kind})")
        out.append(f"--- tables on page {p.number}")
        for t in (t for t in doc.tables if t.page == p.number):
            out.append(f"  table bbox=({t.bbox[0]:.1f}, {t.bbox[1]:.1f}, {t.bbox[2]:.1f}, "
                       f"{t.bbox[3]:.1f}) rows={len(t.rows)}")
            for ri, row in enumerate(t.rows):
                cells = ", ".join(f"x={c[0]:.1f}-{c[2]:.1f}" if c else "merged" for c in row)
                tops = [c[1] for c in row if c]
                bots = [c[3] for c in row if c]
                out.append(f"    row {ri}: y={min(tops):.1f}-{max(bots):.1f} | {cells}")
        out.append(f"--- images on page {p.number}")
        for im in (i for i in doc.images if i.page == p.number):
            out.append(f"  x={im.x0:6.1f}-{im.x1:6.1f} y={im.top:6.1f}-{im.bottom:6.1f} "
                       f"src={im.src_w}x{im.src_h} data={(im.digest or '-')[:10]}")
    path.write_text("\n".join(out), encoding="utf-8")

def main() -> None:
    ap = argparse.ArgumentParser(description="Dump PDF text, lines, tables and pictures.")
    ap.add_argument("files", nargs="+", help=".pdf or .docx files")
    ap.add_argument("--out", default="output", help="output folder (default: output)")
    args = ap.parse_args()

    out_dir = Path(args.out)
    for f in args.files:
        pdf = to_pdf(f, out_dir)
        doc = extract(pdf)
        dump_path = out_dir / f"dump_{Path(f).stem}.txt"
        write_dump(doc, dump_path)
        print(f"{Path(f).name}: {len(doc.pages)} page(s), {len(doc.lines)} lines, "
              f"{len(doc.hsegments)} thin shapes, {len(doc.tables)} table(s), "
              f"{len(doc.images)} image(s) -> {dump_path}")

if __name__ == "__main__":
    main()
