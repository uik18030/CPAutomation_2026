"""Step 1a: read text lines (with page, position and font info) from a PDF.

Run:
    python -m cpa.extract tests/data/File_Old_update.docx tests/data/File_New_update.docx
Writes output/dump_<name>.txt (UTF-8) for each input.
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from pathlib import Path

import pdfplumber

from cpa.convert import to_pdf

BBox = tuple[float, float, float, float]  # x0, top, x1, bottom

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
class Document:
    path: Path
    pages: list[PageInfo]
    lines: list[Line]
    hsegments: list[HSegment]
    tables: list[TableInfo] = field(default_factory=list)

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

def _bbox(v) -> BBox:
    return (float(v[0]), float(v[1]), float(v[2]), float(v[3]))

def _logical_columns(rows: list[list[BBox | None]], min_w: float = 8.0,
                     snap: float = 1.0) -> list[list[BBox | None]]:
    """Re-index cells on the table's real columns.

    Shaded cells / cell padding can make pdfplumber see extra narrow columns
    (e.g. 5 pt strips next to the borders). A column boundary that would create
    a column narrower than *min_w* is removed - the boundary seen in fewer rows
    goes first, the outer table edges are always kept.
    """
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
                continue  # narrow strip that collapsed onto one boundary
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

    with pdfplumber.open(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            pages.append(PageInfo(pno, float(page.width), float(page.height)))

            for tl in page.extract_text_lines(strip=True, return_chars=True):
                chars = [
                    Char(
                        text=c["text"],
                        x0=float(c["x0"]), top=float(c["top"]),
                        x1=float(c["x1"]), bottom=float(c["bottom"]),
                        font=c.get("fontname", ""),
                        size=round(float(c.get("size", 0)), 2),
                        color=_color(c),
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
                        continue  # skip empty rows and hairline rows (double borders)
                    rows.append(cells)
                if rows:
                    tables.append(TableInfo(pno, _bbox(t.bbox), _logical_columns(rows)))

    lines.sort(key=lambda l: (l.page, round(l.top, 1), l.x0))
    segs.sort(key=lambda s: (s.page, s.y, s.x0))
    tables.sort(key=lambda t: (t.page, t.bbox[1], t.bbox[0]))
    return Document(Path(pdf_path), pages, lines, segs, tables)

def write_dump(doc: Document, path: Path) -> None:
    out: list[str] = [f"# {doc.path.name}"]
    for p in doc.pages:
        out.append(f"\n=== page {p.number}  ({p.width:.0f} x {p.height:.0f} pt)")
        for l in (l for l in doc.lines if l.page == p.number):
            sizes = sorted({c.size for c in l.chars})
            colors = sorted({str(c.color) for c in l.chars})
            out.append(f"[y={l.top:6.1f}-{l.bottom:6.1f} x={l.x0:6.1f}-{l.x1:6.1f} "
                       f"size={sizes} color={colors}] {l.text}")
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
    path.write_text("\n".join(out), encoding="utf-8")

def main() -> None:
    ap = argparse.ArgumentParser(description="Dump PDF text lines with positions (Step 1a).")
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
              f"{len(doc.hsegments)} thin horizontal shapes, {len(doc.tables)} table(s) -> {dump_path}")

if __name__ == "__main__":
    main()
