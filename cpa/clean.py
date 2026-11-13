"""Step 1b: turn extracted lines into clean blocks for comparison.

- strikethrough: thin line through the middle of characters -> char.struck
- table text sorted into cells (table / row / column)
- body lines joined into paragraphs
- headers/footers kept separately, page numbers removed
- numbered headings found; every block gets the nearest heading above it
- characters normalised for comparison (full/half width, punctuation)

Run:
    python -m cpa.clean tests/data/File_Old_update.docx tests/data/File_New_update.docx
Writes output/clean_<name>.txt (UTF-8).
"""
from __future__ import annotations

import argparse
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from cpa.convert import to_pdf
from cpa.extract import Char, Document, PageInfo, TableInfo, extract

# ------------------------------------------------------------------ settings
HF_TOP_RATIO = 0.075        # text above 7.5 % of the page height = header
HF_BOTTOM_RATIO = 0.925     # text below 92.5 % of the page height = footer
STRIKE_BAND = (0.35, 0.65)  # strike line must cross this part of the character height
STRIKE_Y_TOL = 1.5          # pt: pieces at about the same height are joined
STRIKE_X_GAP = 1.0          # pt: max gap between pieces that are joined
STRIKE_MIN_COVER = 0.5      # line must cover at least this part of a character's width
HEADING_MAX_LEN = 40
COVER_HEADING = "表紙"
HEADER_LABEL = "ヘッダー"
FOOTER_LABEL = "フッター"

# ------------------------------------------------------------------ data
@dataclass
class TChar:
    t: str                   # normalised character used for comparison (always 1 char)
    raw: str                 # character as in the PDF ("" for inserted spaces/breaks)
    page: int
    x0: float = 0.0
    top: float = 0.0
    x1: float = 0.0
    bottom: float = 0.0
    size: float = 0.0
    font: str = ""
    color: tuple | None = None
    struck: bool = False
    virtual: bool = False    # inserted space or paragraph break, no box on the page

@dataclass
class Block:
    kind: str                # "heading" | "para" | "cell" | "hf"
    page: int
    top: float
    heading: str = ""        # value for the Sheet/Header column
    title: str = ""          # heading blocks: text without its number
    table: int | None = None
    row: int | None = None
    col: int | None = None
    chars: list[TChar] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "".join(c.t for c in self.chars)

    def bboxes(self) -> dict[int, tuple[float, float, float, float]]:
        """Bounding box of the block on each page it appears on."""
        out: dict[int, tuple[float, float, float, float]] = {}
        for c in self.chars:
            if c.virtual:
                continue
            b = out.get(c.page)
            out[c.page] = (c.x0, c.top, c.x1, c.bottom) if b is None else (
                min(b[0], c.x0), min(b[1], c.top), max(b[2], c.x1), max(b[3], c.bottom))
        return out

@dataclass
class CleanDoc:
    source: Path
    pdf: Path
    pages: list[PageInfo]
    blocks: list[Block]

@dataclass
class _Sub:
    """One text line (or the part of a line inside one table cell)."""
    page: int
    top: float
    bottom: float
    x0: float
    x1: float
    tchars: list[TChar]
    raw: str
    size: float

# ------------------------------------------------------------------ normalising
_PUNCT = {
    "、": ",", "。": ".", "，": ",", "．": ".", "：": ":", "；": ";",
    "“": '"', "”": '"', "„": '"', "〝": '"', "〟": '"', "＂": '"',
    "‘": "'", "’": "'", "＇": "'", "‐": "-", "‑": "-", "−": "-",
}
_BULLETS = set("•・●○■□◆◇▪▫➢►▶▷※")

def norm_char(ch: str) -> str:
    if ch.isspace():
        return " "
    if unicodedata.category(ch) == "Co":       # private-use glyph (Symbol/Wingdings bullet)
        return "•"
    o = ord(ch)
    if 0x2460 <= o <= 0x24FF or 0x3251 <= o <= 0x32BF:
        n = ch                                  # keep ① ㉑ ... as they are
    else:
        n = unicodedata.normalize("NFKC", ch)
    return "".join(_PUNCT.get(x, x) for x in n)

def _starts_with_bullet(raw: str) -> bool:
    s = raw.lstrip()
    return bool(s) and (s[0] in _BULLETS or unicodedata.category(s[0]) == "Co")

# ------------------------------------------------------------------ headings
_D = r"[0-9０-９]"
_HEADING_RE = re.compile(
    rf"^(?P<num>"
    rf"第(?:{_D}+|[一二三四五六七八九十百]+)[章節項条部編]"
    rf"|[\(（]{_D}+[\)）]"
    rf"|[\u2460-\u2473]"
    rf"|{_D}+(?:[\.．]{_D}+)*[\.．\)）]?"
    rf")(?P<sep>[\s\u3000]*)(?P<title>.+)$"
)
_DIGITS = "0123456789０１２３４５６７８９"

def heading_title(raw: str) -> str | None:
    """Return the heading text without its number, or None if *raw* is not a heading."""
    s = raw.strip()
    if not s or len(s) > HEADING_MAX_LEN or s.endswith(("。", "．")):
        return None
    m = _HEADING_RE.match(s)
    if not m:
        return None
    num, sep, title = m["num"], m["sep"], m["title"].strip()
    if not title or title[0] in _DIGITS + "/／:：-－,，.．%％":
        return None
    if num[-1] in _DIGITS and not sep:
        # "3つの条件", "10.5mm" are not headings; "1.1エンディング" is
        if not ("." in num or "．" in num) or title[0].isascii():
            return None
    return title

# ------------------------------------------------------------------ page numbers
_PAGE_NO_RE = re.compile(
    r"(?i)(?:page\s*\d+(?:\s*(?:of|/)\s*\d+)?"
    r"|\bp\.\s*\d+"
    r"|-\s*\d+\s*-"
    r"|(?<![\d/])\d{1,4}\s*/\s*\d{1,4}(?![\d/])"
    r"|\d+\s*of\s*\d+"
    r"|\d+\s*(?:ページ|頁))"
)

def _strip_page_numbers(tcs: list[TChar]) -> list[TChar]:
    s = "".join(c.t for c in tcs)
    drop: set[int] = set()
    for m in _PAGE_NO_RE.finditer(s):
        drop.update(range(m.start(), m.end()))
    kept = [c for i, c in enumerate(tcs) if i not in drop]
    if re.fullmatch(r"[\s\d]*", "".join(c.t for c in kept)):
        return []                               # only a bare page number was left
    return kept

# ------------------------------------------------------------------ strikethrough
def _merge_hsegments(doc: Document) -> dict[int, list[list[float]]]:
    """Join touching thin pieces. Returns page -> [x0, x1, ylo, yhi]."""
    out: dict[int, list[list[float]]] = {}
    for s in sorted(doc.hsegments, key=lambda s: (s.page, s.y, s.x0)):
        merged = out.setdefault(s.page, [])
        for m in merged:
            if (abs((m[2] + m[3]) / 2 - s.y) <= STRIKE_Y_TOL
                    and s.x0 <= m[1] + STRIKE_X_GAP and s.x1 >= m[0] - STRIKE_X_GAP):
                m[0], m[1] = min(m[0], s.x0), max(m[1], s.x1)
                m[2], m[3] = min(m[2], s.y), max(m[3], s.y)
                break
        else:
            merged.append([s.x0, s.x1, s.y, s.y])
    return out

def _struck_ids(doc: Document) -> set[int]:
    segs_by_page = _merge_hsegments(doc)
    lo, hi = STRIKE_BAND
    struck: set[int] = set()
    for line in doc.lines:
        segs = segs_by_page.get(line.page)
        if not segs:
            continue
        for c in line.chars:
            h, w = c.bottom - c.top, c.x1 - c.x0
            if h <= 0 or w <= 0 or c.text.isspace():
                continue
            band_lo, band_hi = c.top + h * lo, c.top + h * hi
            for x0, x1, ylo, yhi in segs:
                if band_lo <= (ylo + yhi) / 2 <= band_hi and min(c.x1, x1) - max(c.x0, x0) >= w * STRIKE_MIN_COVER:
                    struck.add(id(c))
                    break
    return struck

    
def _is_cjk(ch: str) -> bool:
    """Japanese/Chinese characters and full-width punctuation (no spaces next to these)."""
    return unicodedata.east_asian_width(ch) in ("W", "F")


# ------------------------------------------------------------------ lines -> subs
def _make_sub(page: int, chars: list[Char], struck: set[int]) -> _Sub | None:
    chars = sorted(chars, key=lambda c: c.x0)
    real = [c for c in chars if not c.text.isspace()]
    if not real:
        return None
    tcs: list[TChar] = []
    raw_parts: list[str] = []
    prev: Char | None = None
    for c in chars:
        if prev is not None and not c.text.isspace() and not prev.text.isspace():
            if (c.x0 - prev.x1 > max(1.0, 0.15 * max(c.size, 1.0))
                and not _is_cjk(prev.text[-1]) and not _is_cjk(c.text[0])):
                    tcs.append(TChar(" ", "", page, virtual=True))
                    raw_parts.append(" ")
        is_struck = id(c) in struck
        for ch in c.text:
            for k in norm_char(ch):
                tcs.append(TChar(k, ch, page, c.x0, c.top, c.x1, c.bottom,
                                 c.size, c.font, c.color, is_struck, False))
        raw_parts.append(c.text)
        prev = c
    return _Sub(page, min(c.top for c in real), max(c.bottom for c in real),
                min(c.x0 for c in real), max(c.x1 for c in real),
                tcs, "".join(raw_parts), max(c.size for c in real))


def _cell_of(c: Char, tables: list[tuple[int, TableInfo]]) -> tuple[int, int, int] | None:
    cx, cy = (c.x0 + c.x1) / 2, (c.top + c.bottom) / 2
    for ti, t in tables:
        x0, t0, x1, b0 = t.bbox
        if not (x0 - 0.5 <= cx <= x1 + 0.5 and t0 - 0.5 <= cy <= b0 + 0.5):
            continue
        for ri, row in enumerate(t.rows):
            for ci, cell in enumerate(row):
                if cell and cell[0] - 0.5 <= cx <= cell[2] + 0.5 and cell[1] - 0.5 <= cy <= cell[3] + 0.5:
                    return (ti, ri, ci)
    return None

# ------------------------------------------------------------------ joining
_END = ("。", "．", "!", "！", "?", "？", ":", "：")

def _continues(prev: _Sub, nxt: _Sub) -> bool:
    """Does *nxt* continue the paragraph that *prev* is in?"""
    if prev.raw.rstrip().endswith(_END) or _starts_with_bullet(nxt.raw):
        return False
    if nxt.page == prev.page:
        lh = max(prev.bottom - prev.top, 1.0)
        if nxt.top < prev.top or nxt.top - prev.bottom > lh:
            return False
    return True

def _collapse(tcs: list[TChar]) -> list[TChar]:
    out: list[TChar] = []
    for c in tcs:
        if c.t == " " and (not out or out[-1].t in (" ", "\n")):
            continue
        if c.t == "\n" and out and out[-1].t == " ":
            out.pop()
        out.append(c)
    while out and out[-1].t in (" ", "\n"):
        out.pop()
    return out

def _join_subs(subs: list[_Sub]) -> list[TChar]:
    out: list[TChar] = []
    for s in subs:
        tcs = s.tchars
        last = next((c for c in reversed(out) if not c.virtual), None)
        if last is not None and tcs:
            first = tcs[0]
            if last.t == "-" and first.t.isascii() and first.t.islower():
                out.remove(last)                      # "trans-" + "mission"
            elif last.t.isascii() and last.t.isalnum() and first.t.isascii() and first.t.isalnum():
                out.append(TChar(" ", "", s.page, virtual=True))
        out.extend(tcs)
    return out

def _paragraphs(subs: list[_Sub]) -> list[list[_Sub]]:
    paras: list[list[_Sub]] = []
    for s in subs:
        if paras and _continues(paras[-1][-1], s):
            paras[-1].append(s)
        else:
            paras.append([s])
    return paras

# ------------------------------------------------------------------ tables across pages
def _row_text(cell_subs: dict, ti: int, ri: int) -> str:
    keys = sorted(k for k in cell_subs if k[0] == ti and k[1] == ri)
    return "|".join("".join(c.t for s in cell_subs[k] for c in s.tchars).strip() for k in keys)

def _same_columns(a: TableInfo, b: TableInfo) -> bool:
    xa = sorted(c[0] for c in a.rows[0] if c)
    xb = sorted(c[0] for c in b.rows[0] if c)
    return len(xa) == len(xb) and all(abs(p - q) <= 2.0 for p, q in zip(xa, xb))

def _table_mapping(doc: Document, cell_subs: dict) -> dict[int, tuple[int, int, bool]]:
    """table index -> (logical table id, row offset, drop first row)."""
    page_h = {p.number: p.height for p in doc.pages}
    mapping: dict[int, tuple[int, int, bool]] = {}
    base: list[int] = []
    total_rows: list[int] = []
    prev: tuple[int, TableInfo] | None = None
    for ti, t in enumerate(doc.tables):
        cont = False
        if prev is not None:
            pti, pt = prev
            cont = (t.page == pt.page + 1
                    and t.bbox[1] < page_h[t.page] * 0.25
                    and pt.bbox[3] > page_h[pt.page] * 0.6
                    and _same_columns(pt, t))
        if cont:
            lid = mapping[prev[0]][0]
            head = _row_text(cell_subs, ti, 0)
            drop = bool(head) and head == _row_text(cell_subs, base[lid], 0)
            mapping[ti] = (lid, total_rows[lid] - (1 if drop else 0), drop)
            total_rows[lid] += len(t.rows) - (1 if drop else 0)
        else:
            mapping[ti] = (len(base), 0, False)
            base.append(ti)
            total_rows.append(len(t.rows))
        prev = (ti, t)
    return mapping

# ------------------------------------------------------------------ main step
def clean(doc: Document, source: Path | None = None) -> CleanDoc:
    struck = _struck_ids(doc)
    page_h = {p.number: p.height for p in doc.pages}
    tables_by_page: dict[int, list[tuple[int, TableInfo]]] = {}
    for ti, t in enumerate(doc.tables):
        tables_by_page.setdefault(t.page, []).append((ti, t))

    body_subs: list[_Sub] = []
    hf_subs: list[tuple[str, _Sub]] = []
    cell_subs: dict[tuple[int, int, int], list[_Sub]] = {}

    for line in doc.lines:
        h = page_h[line.page]
        if line.top < h * HF_TOP_RATIO or line.bottom > h * HF_BOTTOM_RATIO:
            sub = _make_sub(line.page, line.chars, struck)
            if sub:
                hf_subs.append((HEADER_LABEL if line.top < h * HF_TOP_RATIO else FOOTER_LABEL, sub))
            continue
        groups: dict[tuple[int, int, int] | None, list[Char]] = {}
        for c in line.chars:
            groups.setdefault(_cell_of(c, tables_by_page.get(line.page, [])), []).append(c)
        for key, chars in groups.items():
            sub = _make_sub(line.page, chars, struck)
            if sub is None:
                continue
            if key is None:
                body_subs.append(sub)
            else:
                cell_subs.setdefault(key, []).append(sub)

    content: list[tuple[tuple, Block]] = []

    # body: paragraphs and headings
    page_right: dict[int, float] = {}
    for s in body_subs:
        page_right[s.page] = max(page_right.get(s.page, 0.0), s.x1)

    def heading_of(s: _Sub) -> str | None:
        t = heading_title(s.raw)
        if t is None or s.x1 >= page_right[s.page] - 2 * s.size:   # full-width line = wrapped text
            return None
        return t

    paras: list[list[_Sub]] = []
    for s in body_subs:
        if heading_of(s) is not None:
            paras.append([s])
            paras.append([])                       # nothing joins onto a heading
        elif paras and paras[-1] and heading_of(paras[-1][-1]) is None and _continues(paras[-1][-1], s):
            paras[-1].append(s)
        else:
            paras.append([s])
    for p in (p for p in paras if p):
        first = p[0]
        title = heading_of(first) if len(p) == 1 else None
        b = Block("heading" if title else "para", first.page, first.top,
                  title=title or "", chars=_collapse(_join_subs(p)))
        if b.chars:
            content.append(((first.page, first.top, 0, 0), b))

    # table cells
    mapping = _table_mapping(doc, cell_subs)
    merged_cells: dict[tuple[int, int, int], list[_Sub]] = {}
    row_tops: dict[tuple[int, int], tuple[int, float]] = {}
    for (ti, ri, ci), subs in cell_subs.items():
        lid, off, drop = mapping[ti]
        if drop and ri == 0:
            continue
        merged_cells.setdefault((lid, off + ri, ci), []).extend(subs)
        t = doc.tables[ti]
        row_tops.setdefault((lid, off + ri), (t.page, min(c[1] for c in t.rows[ri] if c)))
    for (lid, row, col), subs in merged_cells.items():
        subs.sort(key=lambda s: (s.page, s.top, s.x0))
        tcs: list[TChar] = []
        for p in _paragraphs(subs):
            if tcs:
                tcs.append(TChar("\n", "", p[0].page, virtual=True))
            tcs.extend(_join_subs(p))
        b = Block("cell", subs[0].page, subs[0].top, table=lid, row=row, col=col, chars=_collapse(tcs))
        if b.chars:
            page, rtop = row_tops[(lid, row)]
            content.append(((page, rtop, 1, row, col), b))


    # nearest heading above each block
    content.sort(key=lambda kv: kv[0])
    current = COVER_HEADING
    blocks: list[Block] = []
    for _, b in content:
        if b.kind == "heading":
            current = b.title
        b.heading = current
        blocks.append(b)

    # headers / footers (page numbers removed)
    for label, s in hf_subs:
        tcs = _collapse(_strip_page_numbers(s.tchars))
        if tcs:
            blocks.append(Block("hf", s.page, s.top, heading=label, chars=tcs))

    return CleanDoc(source or doc.path, doc.path, doc.pages, blocks)

def load(src: str | Path, out_dir: str | Path = "output") -> CleanDoc:
    """Convert (if needed), extract and clean one document."""
    pdf = to_pdf(src, out_dir)
    return clean(extract(pdf), source=Path(src))

# ------------------------------------------------------------------ dump
def _marked(b: Block) -> str:
    out: list[str] = []
    on = False
    for c in b.chars:
        if not c.virtual and c.struck != on:
            out.append("~~")
            on = c.struck
        out.append(" ⏎ " if c.t == "\n" else (c.t if c.t == "•" else (c.raw or c.t)))
    if on:
        out.append("~~")
    return "".join(out)

def write_clean_dump(cd: CleanDoc, path: Path) -> None:
    out = [f"# {cd.source.name}  ({len(cd.blocks)} blocks)   ~~text~~ = struck through"]
    for i, b in enumerate(cd.blocks, start=1):
        loc = f" T{b.table}R{b.row}C{b.col}" if b.table is not None else ""
        title = f" title={b.title}" if b.kind == "heading" else ""
        out.append(f"\n[{i:03d}] {b.kind:<7} p{b.page} y={b.top:6.1f}{loc} heading={b.heading}{title}")
        out.append(f"      {_marked(b)}")
    path.write_text("\n".join(out), encoding="utf-8")

def main() -> None:
    ap = argparse.ArgumentParser(description="Clean extracted text into blocks (Step 1b).")
    ap.add_argument("files", nargs="+", help=".pdf or .docx files")
    ap.add_argument("--out", default="output", help="output folder (default: output)")
    args = ap.parse_args()

    out_dir = Path(args.out)
    for f in args.files:
        cd = load(f, out_dir)
        dump_path = out_dir / f"clean_{Path(f).stem}.txt"
        write_clean_dump(cd, dump_path)
        kinds = {k: sum(1 for b in cd.blocks if b.kind == k) for k in ("heading", "para", "cell", "hf")}
        struck_n = sum(1 for b in cd.blocks for c in b.chars if c.struck)
        print(f"{Path(f).name}: {kinds['heading']} headings, {kinds['para']} paragraphs, "
              f"{kinds['cell']} cells, {kinds['hf']} header/footer, {struck_n} struck chars -> {dump_path}")

if __name__ == "__main__":
    main()
