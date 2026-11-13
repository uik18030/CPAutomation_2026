"""Step 1d: screenshots of each change with red boxes.

Old Spec : red box around removed content, otherwise plain
Latest   : red box around new/changed content, solid red marker where content was removed

Run:
    python -m cpa.shots tests/data/File_Old_update.docx tests/data/File_New_update.docx
Writes output/shots/NN_old.png, NN_new.png and output/changes.txt
"""
from __future__ import annotations

import argparse
import shutil
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw

from cpa.clean import Block, load
from cpa.compare import Box, Change, compare, write_changes

# ------------------------------------------------------------------ settings
DPI = 150
SCALE = DPI / 72          # PDF points -> pixels
PAD_Y = 30.0              # pt of context above and below the change
BOX_PAD = 2.0             # pt between text and red box
BOX_WIDTH = 3             # px line width of the red box
MARK_GROW = 1.0           # pt: markers are drawn slightly bigger so they are visible
PAGE_GAP = 8              # px grey gap between stacked page crops
RED = (255, 0, 0)
CACHE_PAGES = 4           # rendered pages kept in memory per document

class PageRenderer:
    """Render PDF pages to images, keeping only a few in memory."""

    def __init__(self, pdf_path: Path):
        self.pdf = pdfium.PdfDocument(str(pdf_path))
        self.cache: dict[int, Image.Image] = {}
        self.order: list[int] = []

    def get(self, pno: int) -> Image.Image:
        if pno in self.cache:
            return self.cache[pno]
        page = self.pdf[pno - 1]
        img = page.render(scale=SCALE).to_pil().convert("RGB")
        page.close()
        self.cache[pno] = img
        self.order.append(pno)
        if len(self.order) > CACHE_PAGES:
            del self.cache[self.order.pop(0)]
        return img

    def close(self) -> None:
        self.cache.clear()
        self.pdf.close()

def _regions(blocks: list[Block], *box_lists: list[Box]) -> dict[int, list[float]]:
    """Vertical range (top, bottom) to show on each page."""
    reg: dict[int, list[float]] = {}

    def add(p: int, top: float, bottom: float) -> None:
        r = reg.get(p)
        reg[p] = [top, bottom] if r is None else [min(r[0], top), max(r[1], bottom)]

    for b in blocks:
        for p, (_, top, _, bottom) in b.bboxes().items():
            add(p, top, bottom)
    for boxes in box_lists:
        for p, _, top, _, bottom in boxes:
            add(p, top, bottom)
    return reg

def _shot(rend: PageRenderer, regions: dict[int, list[float]],
          boxes: list[Box], markers: list[Box], out_path: Path) -> Path | None:
    if not regions:
        return None
    crops: list[Image.Image] = []
    for p in sorted(regions):
        img = rend.get(p).copy()
        d = ImageDraw.Draw(img)
        for bp, x0, top, x1, bottom in boxes:
            if bp == p:
                d.rectangle([(x0 - BOX_PAD) * SCALE, (top - BOX_PAD) * SCALE,
                             (x1 + BOX_PAD) * SCALE, (bottom + BOX_PAD) * SCALE],
                            outline=RED, width=BOX_WIDTH)
        for mp, x0, top, x1, bottom in markers:
            if mp == p:
                d.rectangle([(x0 - MARK_GROW) * SCALE, (top - MARK_GROW) * SCALE,
                             (x1 + MARK_GROW) * SCALE, (bottom + MARK_GROW) * SCALE], fill=RED)
        top, bottom = regions[p]
        y0 = max(0, int((top - PAD_Y) * SCALE))
        y1 = min(img.height, int((bottom + PAD_Y) * SCALE))
        crops.append(img.crop((0, y0, img.width, y1)))

    width = max(c.width for c in crops)
    height = sum(c.height for c in crops) + PAGE_GAP * (len(crops) - 1)
    out = Image.new("RGB", (width, height), (200, 200, 200))
    y = 0
    for c in crops:
        out.paste(c, (0, y))
        y += c.height + PAGE_GAP
    out.save(out_path)
    return out_path

def make_shots(changes: list[Change], old_pdf: Path, new_pdf: Path,
               out_dir: Path) -> dict[int, tuple[Path | None, Path | None]]:
    """Create NN_old.png / NN_new.png for every change. Returns {no: (old, new)}."""
    shots_dir = out_dir / "shots"
    if shots_dir.exists():
        shutil.rmtree(shots_dir)
    shots_dir.mkdir(parents=True)

    old_r, new_r = PageRenderer(old_pdf), PageRenderer(new_pdf)
    result: dict[int, tuple[Path | None, Path | None]] = {}
    try:
        for c in changes:
            old_reg = _regions(c.old_blocks, c.old_boxes, c.old_anchor)
            new_reg = _regions(c.new_blocks, c.new_boxes, c.new_markers)
            o = _shot(old_r, old_reg, c.old_boxes, [], shots_dir / f"{c.no:02d}_old.png")
            n = _shot(new_r, new_reg, c.new_boxes, c.new_markers, shots_dir / f"{c.no:02d}_new.png")
            result[c.no] = (o, n)
    finally:
        old_r.close()
        new_r.close()
    return result

def main() -> None:
    ap = argparse.ArgumentParser(description="Screenshots with red boxes (Step 1d).")
    ap.add_argument("old", help="old .pdf or .docx")
    ap.add_argument("new", help="new .pdf or .docx")
    ap.add_argument("--out", default="output", help="output folder (default: output)")
    args = ap.parse_args()

    out_dir = Path(args.out)
    old = load(args.old, out_dir)
    new = load(args.new, out_dir)
    changes = compare(old, new)
    write_changes(changes, old, new, out_dir / "changes.txt")
    shots = make_shots(changes, old.pdf, new.pdf, out_dir)
    made = sum((o is not None) + (n is not None) for o, n in shots.values())
    print(f"{len(changes)} changes, {made} screenshots -> {out_dir / 'shots'}")

if __name__ == "__main__":
    main()
