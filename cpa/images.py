"""Step 2a: image changes.

Pictures are matched across the whole document by how they look, so pictures pushed
to another page still match. Each pair is then compared:
    .docx input -> rotation/flip read from Word itself (exact)
    otherwise   -> compared as rendered (also rotated/flipped versions)
    rotated / flipped / resized / content changed -> Modified
    only in old -> Deleted, only in new -> New
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path

import imagehash
import pypdfium2 as pdfium
from PIL import Image

from cpa.convert import get_password

from cpa.clean import (COVER_HEADING, FOOTER_LABEL, HEADER_LABEL, HF_BOTTOM_RATIO,
                       HF_TOP_RATIO, CleanDoc)
from cpa.compare import Box, Change

# ------------------------------------------------------------------ settings
DPI = 150
SCALE = DPI / 72
SAME_DIST = 6            # hash distance (0-64) at or below this = looks the same
ROT_DIST = 16            # looser limit when the image data is identical (only orientation can differ)
PAIR_DIST = 20           # pictures at or below this (in any orientation) are the same picture
FLIP_MARGIN = 4          # a flip must match clearly better than a rotation to be reported as a flip
SIZE_TOL = 0.05          # size change above 5 % = resized
MAX_PAIR_CELLS = 250_000
HEADING_BONUS = 0.25


ORIENTATIONS = [
    ("rotated 180°", Image.Transpose.ROTATE_180),
    ("rotated 90° clockwise", Image.Transpose.ROTATE_270),
    ("rotated 90° counterclockwise", Image.Transpose.ROTATE_90),
    ("flipped horizontally", Image.Transpose.FLIP_LEFT_RIGHT),
    ("flipped vertically", Image.Transpose.FLIP_TOP_BOTTOM),
]

_NS = {
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}
_WP_TAGS = ("{%s}inline" % _NS["wp"], "{%s}anchor" % _NS["wp"])

@dataclass
class Pic:
    page: int
    box: tuple[float, float, float, float]       # x0, top, x1, bottom (pt)
    digest: str | None
    heading: str
    area: Box                                    # area to show in the screenshot
    hf: bool
    row: tuple | None                            # (table, row) if inside a table row
    h0: imagehash.ImageHash | None = None        # hash as rendered
    hx: dict = field(default_factory=dict)       # hashes of rotated/flipped versions
    xf: tuple | None = None                      # (rotation°, flipH, flipV) read from Word

    @property
    def w(self) -> float:
        return self.box[2] - self.box[0]

    @property
    def h(self) -> float:
        return self.box[3] - self.box[1]

    def as_box(self) -> Box:
        return (self.page, *self.box)

class _Pages:
    def __init__(self, pdf_path):
        self.pdf = pdfium.PdfDocument(str(pdf_path), password=get_password(pdf_path))
        self.cache: dict[int, Image.Image] = {}

    def get(self, pno: int) -> Image.Image:
        if pno not in self.cache:
            if len(self.cache) >= 3:
                self.cache.pop(next(iter(self.cache)))
            page = self.pdf[pno - 1]
            self.cache[pno] = page.render(scale=SCALE).to_pil().convert("RGB")
            page.close()
        return self.cache[pno]

    def close(self) -> None:
        self.cache.clear()
        self.pdf.close()

# ------------------------------------------------------------------ Word rotation
def _docx_xforms(path) -> list[tuple[int, bool, bool]] | None:
    """Rotation/flip of every picture in a .docx body, in document order (None if not .docx)."""
    p = Path(path)
    if p.suffix.lower() != ".docx" or not p.exists():
        return None
    try:
        with zipfile.ZipFile(p) as z:
            root = ET.fromstring(z.read("word/document.xml"))
    except Exception:
        return None
    out: list[tuple[int, bool, bool]] = []
    for el in root.iter():
        if el.tag not in _WP_TAGS or el.find(".//a:blip", _NS) is None:
            continue
        xf = el.find(".//pic:spPr/a:xfrm", _NS)
        rot = int(xf.get("rot", "0")) if xf is not None else 0
        fh = xf is not None and xf.get("flipH") in ("1", "true")
        fv = xf is not None and xf.get("flipV") in ("1", "true")
        out.append((round(rot / 60000) % 360, fh, fv))
    return out

def _xf_details(a: tuple, b: tuple) -> list[str]:
    ra, fha, fva = a
    rb, fhb, fvb = b
    fh, fv = fha != fhb, fva != fvb
    rot = (rb - ra) % 360
    if fh and fv:                               # flipped both ways = turned 180°
        rot, fh, fv = (rot + 180) % 360, False, False
    out: list[str] = []
    if rot == 90:
        out.append("image rotated 90° clockwise")
    elif rot == 270:
        out.append("image rotated 90° counterclockwise")
    elif rot:
        out.append(f"image rotated {rot}°")
    if fh:
        out.append("image flipped horizontally")
    if fv:
        out.append("image flipped vertically")
    return out

# ------------------------------------------------------------------ collecting
def _crop(img: Image.Image, box) -> Image.Image | None:
    x0, t = int(box[0] * SCALE) + 1, int(box[1] * SCALE) + 1
    x1, b = int(box[2] * SCALE) - 1, int(box[3] * SCALE) - 1
    x0, t, x1, b = max(0, x0), max(0, t), min(img.width, x1), min(img.height, b)
    if x1 - x0 < 4 or b - t < 4:
        return None
    return img.crop((x0, t, x1, b)).convert("L")

def _heading_at(doc: CleanDoc, page: int, top: float) -> str:
    current = COVER_HEADING
    for b in doc.blocks:
        if b.kind != "heading":
            continue
        if (b.page, b.top) <= (page, top + 0.5):
            current = b.title
        else:
            break
    return current

def _row_at(doc: CleanDoc, page: int, box) -> tuple[Box, tuple | None]:
    cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
    for key, (p, x0, t, x1, b) in (getattr(doc, "table_rows", None) or {}).items():
        if p == page and x0 <= cx <= x1 and t <= cy <= b:
            return (p, x0, t, x1, b), key
    return (page, *box), None

def _pics(doc: CleanDoc, pages: _Pages) -> list[Pic]:
    heights = {p.number: p.height for p in doc.pages}
    out: list[Pic] = []
    for im in getattr(doc, "images", None) or []:
        h = heights[im.page]
        box = (im.x0, im.top, im.x1, im.bottom)
        if im.top < h * HF_TOP_RATIO or im.bottom > h * HF_BOTTOM_RATIO:
            heading = HEADER_LABEL if im.top < h * HF_TOP_RATIO else FOOTER_LABEL
            pic = Pic(im.page, box, im.digest, heading, (im.page, *box), True, None)
        else:
            area, row = _row_at(doc, im.page, box)
            pic = Pic(im.page, box, im.digest, _heading_at(doc, im.page, im.top), area, False, row)
        crop = _crop(pages.get(im.page), box)
        if crop is not None:
            pic.h0 = imagehash.phash(crop)
            pic.hx = {name: imagehash.phash(crop.transpose(t)) for name, t in ORIENTATIONS}
        out.append(pic)

    # .docx: attach Word's own rotation/flip (only if the picture counts agree)
    xfs = _docx_xforms(doc.source)
    body = [p for p in out if not p.hf]
    if xfs is not None and len(xfs) == len(body):
        for p, xf in zip(body, xfs):
            p.xf = xf
    return out

def _dedupe(pics: list[Pic]) -> list[Pic]:
    """Header/footer pictures repeat on every page - keep the first of each."""
    seen: set[str] = set()
    out: list[Pic] = []
    for p in pics:
        k = f"{p.heading}|{p.digest or p.h0}"
        if k not in seen:
            seen.add(k)
            out.append(p)
    return out

# ------------------------------------------------------------------ comparing
def _dist_same(o: Pic, n: Pic) -> int:
    if o.h0 is None or n.h0 is None:
        return 0 if (o.digest and o.digest == n.digest) else 64
    return o.h0 - n.h0

def _best(o: Pic, n: Pic) -> tuple[str | None, int]:
    """Closest orientation. Rotations are preferred: a flip only wins if it is clearly better."""
    d0 = _dist_same(o, n)
    if o.h0 is None or n.h0 is None:
        return (None, d0)
    rot = min(((name, hx - n.h0) for name, hx in o.hx.items() if name.startswith("rotated")),
              key=lambda x: x[1])
    flip = min(((name, hx - n.h0) for name, hx in o.hx.items() if name.startswith("flipped")),
               key=lambda x: x[1])
    best: tuple[str | None, int] = (None, d0)
    if rot[1] < best[1]:
        best = rot
    if flip[1] + FLIP_MARGIN < best[1]:
        best = flip
    return best

def _score(o: Pic, n: Pic) -> float:
    if o.digest and o.digest == n.digest:
        s = 1.0
    else:
        _, d = _best(o, n)
        s = 1.0 - d / 64 if d <= PAIR_DIST else 0.0
    if s >0 and o.heading == n.heading:
        s += HEADING_BONUS
    return s

def _align(O: list[Pic], N: list[Pic]) -> list[tuple[int, int]]:
    """Pairs (old index, new index), keeping document order."""
    n, m = len(O), len(N)
    if n == 0 or m == 0:
        return []
    if n * m > MAX_PAIR_CELLS:
        keys_o = [p.digest or str(p.h0) for p in O]
        keys_n = [p.digest or str(p.h0) for p in N]
        sm = SequenceMatcher(None, keys_o, keys_n, autojunk=False)
        return [(i1 + k, j1 + k) for tag, i1, i2, j1, j2 in sm.get_opcodes()
                if tag == "equal" for k in range(i2 - i1)]

    sim = [[_score(O[a], N[b]) for b in range(m)] for a in range(n)]
    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    for a in range(n - 1, -1, -1):
        for b in range(m - 1, -1, -1):
            best = max(dp[a + 1][b], dp[a][b + 1])
            if sim[a][b] > 0:
                best = max(best, sim[a][b] + dp[a + 1][b + 1])
            dp[a][b] = best
    pairs: list[tuple[int, int]] = []
    a = b = 0
    while a < n and b < m:
        if sim[a][b] > 0 and abs(dp[a][b] - (sim[a][b] + dp[a + 1][b + 1])) < 1e-9:
            pairs.append((a, b))
            a += 1
            b += 1
        elif abs(dp[a][b] - dp[a + 1][b]) < 1e-9:
            a += 1
        else:
            b += 1

    # leftovers between two pairs, under the same heading: picture was replaced
    bounds = [(-1, -1)] + pairs + [(n, m)]
    extra: list[tuple[int, int]] = []
    for (pi, pj), (qi, qj) in zip(bounds, bounds[1:]):
        for i, j in zip(range(pi + 1, qi), range(pj + 1, qj)):
            if O[i].heading == N[j].heading:
                extra.append((i, j))
    return sorted(pairs + extra)

def _size_details(o: Pic, n: Pic, turned_90: bool) -> list[str]:
    ow, oh = (o.h, o.w) if turned_90 else (o.w, o.h)
    if ow > 0 and oh > 0 and (abs(n.w / ow - 1) > SIZE_TOL or abs(n.h / oh - 1) > SIZE_TOL):
        return [f"image resized ({o.w:.0f}x{o.h:.0f} → {n.w:.0f}x{n.h:.0f} pt)"]
    return []

def _classify(o: Pic, n: Pic) -> list[str]:
    same_data = bool(o.digest) and o.digest == n.digest

    # 1) Word documents: exact rotation/flip from the file
    if o.xf is not None and n.xf is not None and (same_data or not (o.digest and n.digest)):
        details = _xf_details(o.xf, n.xf)
        turned = any("90°" in d for d in details)
        return details + _size_details(o, n, turned)

    # 2) otherwise compare as rendered
    d_same = _dist_same(o, n)
    name, d = _best(o, n)
    limit = ROT_DIST if same_data else SAME_DIST
    details: list[str] = []
    if d_same <= SAME_DIST:
        name = None                                  # looks the same
    elif name is not None and d <= limit:
        details.append(f"image {name}")
    else:
        name = None
        details.append(f"image content changed (difference {d_same}/64)")
    return details + _size_details(o, n, bool(name and "90°" in name))

def _place(pic: Pic, idx: int, mapping: dict[int, int], dst: list[Pic], dst_doc: CleanDoc):
    """Where an unpaired picture would be in the other document: (crop area, marker)."""
    if pic.row is not None:
        r = (getattr(dst_doc, "table_rows", None) or {}).get(pic.row)
        if r:
            p, _, t, _, _ = r
            return r, (p, pic.box[0], t + 1.0, pic.box[2], t + 3.0)
    later = [k for k in mapping if k > idx]
    if later:
        q = dst[mapping[min(later)]]
        y = q.box[1] - 3.0
        return q.area, (q.page, q.box[0], y - 1.0, q.box[2], y + 1.0)
    earlier = [k for k in mapping if k < idx]
    if earlier:
        q = dst[mapping[max(earlier)]]
        y = q.box[3] + 3.0
        return q.area, (q.page, q.box[0], y - 1.0, q.box[2], y + 1.0)
    return None, None

def _changes(O: list[Pic], N: list[Pic], old: CleanDoc, new: CleanDoc) -> list[Change]:
    pairs = _align(O, N)
    o2n = dict(pairs)
    n2o = {j: i for i, j in pairs}
    out: list[Change] = []
    for i, j in pairs:
        o, n = O[i], N[j]
        details = _classify(o, n)
        if details:
            out.append(Change("Modified", n.heading, [], [], [], [n.as_box()], [], details,
                              old_anchor=[o.area], new_anchor=[n.area]))
    for i, o in enumerate(O):
        if i in o2n:
            continue
        ch = Change("Deleted", o.heading, [], [], [o.as_box()], [], [], ["image deleted"],
                    old_anchor=[o.area])
        area, marker = _place(o, i, o2n, N, new)
        if area:
            ch.new_anchor = [area]
            ch.new_markers = [marker]
        out.append(ch)
    for j, n in enumerate(N):
        if j in n2o:
            continue
        ch = Change("New", n.heading, [], [], [], [n.as_box()], [], ["image added"],
                    new_anchor=[n.area])
        area, _ = _place(n, j, n2o, O, old)
        if area:
            ch.old_anchor = [area]
        out.append(ch)
    return out

def image_changes(old: CleanDoc, new: CleanDoc) -> tuple[list[Change], list[Change]]:
    """Return (body image changes, header/footer image changes)."""
    po, pn = _Pages(old.pdf), _Pages(new.pdf)
    try:
        O = _pics(old, po)
        N = _pics(new, pn)
    finally:
        po.close()
        pn.close()
    body = _changes([p for p in O if not p.hf], [p for p in N if not p.hf], old, new)
    hf = _changes(_dedupe([p for p in O if p.hf]), _dedupe([p for p in N if p.hf]), old, new)
    return body, hf
