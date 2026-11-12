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
class Document:
    path: Path
    pages: list[PageInfo]
    lines: list[Line]
    hsegments: list[HSegment]


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


def extract(pdf_path: str | Path) -> Document:
    pages: list[PageInfo] = []
    lines: list[Line] = []
    segs: list[HSegment] = []

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

    lines.sort(key=lambda l: (l.page, round(l.top, 1), l.x0))
    segs.sort(key=lambda s: (s.page, s.y, s.x0))
    return Document(Path(pdf_path), pages, lines, segs)


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
              f"{len(doc.hsegments)} thin horizontal shapes -> {dump_path}")


if __name__ == "__main__":
    main()
