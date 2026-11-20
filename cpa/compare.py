"""Step 1c: compare old and new documents block by block, then character by character.

Run:
    python -m cpa.compare OLD NEW
Writes output/changes.txt (UTF-8).
"""
from __future__ import annotations

import argparse
import re
from collections import Counter
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path

from cpa.clean import Block, CleanDoc, TChar, _marked, load

# ------------------------------------------------------------------ settings
MATCH_MIN = 0.5          # blocks at least this similar are paired as "Modified"
COL_BONUS = 0.1          # table cells in the same column are preferred
MAX_PAIR_CELLS = 250_000 # safety limit for the pairing table (n_old x n_new)
SIZE_TOL = 0.3
COLOR_TOL = 0.15
SYSTEMATIC_RATIO = 0.3
SYSTEMATIC_MIN = 200
MOVED_MIN_CHARS = 4
MOVED_SIM = 0.85


Box = tuple[int, float, float, float, float]  # page, x0, top, x1, bottom

@dataclass
class Change:
    kind: str                                   # "New" | "Deleted" | "Modified"
    heading: str                                # Sheet/Header column
    old_blocks: list[Block] = field(default_factory=list)
    new_blocks: list[Block] = field(default_factory=list)
    old_boxes: list[Box] = field(default_factory=list)    # red boxes on Old Spec
    new_boxes: list[Box] = field(default_factory=list)    # red boxes on Latest Spec
    new_markers: list[Box] = field(default_factory=list)  # red markers on Latest Spec
    details: list[str] = field(default_factory=list)
    no: int = 0
    first_seq: int = 0
    last_seq: int = 0
    old_anchor: list[Box] = field(default_factory=list)   # area to show in old (crop only)
    new_anchor: list[Box] = field(default_factory=list)   # area to show in new (crop only)

# ------------------------------------------------------------------ block helpers
_COMPACT: dict[int, list[tuple[int, TChar]]] = {}

def _compact(b: Block) -> list[tuple[int, TChar]]:
    """Characters used for comparison: no spaces, no inserted (virtual) characters."""
    k = id(b)
    if k not in _COMPACT:
        _COMPACT[k] = [(i, c) for i, c in enumerate(b.chars) if not c.virtual and not c.t.isspace()]
    return _COMPACT[k]

def _ctext(b: Block) -> str:
    return "".join(c.t for _, c in _compact(b))

def _group(b: Block) -> str:
    if b.kind == "cell":
        return "cell"
    if b.kind == "hf":
        return "hf:" + b.heading
    return "text"

_HEAD_NO_RE = re.compile(r"^(?:第\d+[章節項条部編]|\(\d+\)|[\u2460-\u2473]|\d+(?:\.\d+)*[.)]?)")
_CAPTION_NO_RE = re.compile(r"^((?:表|図|Table|Figure|Fig\.?))\d+(?:[.\-]\d+)*")

def _unnumbered(b: Block) -> str | None:
    """Heading / caption text without its number (renumbering alone is not a change)."""
    t = _ctext(b)
    if b.kind == "heading":
        return "#" + _HEAD_NO_RE.sub("", t)
    if b.kind == "para" and _CAPTION_NO_RE.match(t):
        return _CAPTION_NO_RE.sub(r"\1#",t)
    return None

def _key(b:Block) -> tuple[str, str]:
    return (_group(b), _unnumbered(b) or _ctext(b))

def _plain(b: Block) -> str:
    return "".join(" ⏎ " if c.t == "\n" else (c.raw or c.t) for c in b.chars).strip()

def _short(b: Block, n: int = 60) -> str:
    t = _plain(b)
    return t if len(t) <= n else t[:n] + "…"

def _boxes(chars: list[TChar]) -> list[Box]:
    """One box per text line covered by *chars*."""
    out: list[list] = []
    for c in chars:
        if c.virtual:
            continue
        h = max(c.bottom - c.top, 1.0)
        if out and out[-1][0] == c.page and abs(out[-1][2] - c.top) < h / 2:
            b = out[-1]
            b[1], b[2], b[3], b[4] = min(b[1], c.x0), min(b[2], c.top), max(b[3], c.x1), max(b[4], c.bottom)
        else:
            out.append([c.page, c.x0, c.top, c.x1, c.bottom])
    return [tuple(b) for b in out]

def _is_word(ch: str) -> bool:
    return ch.isascii() and (ch.isalnum() or ch == "_")

def _expand_words(seq: list[tuple[int, TChar]], idx: set[int]) -> set[int]:
    """Grow boxed English letters/digits to the whole word (3D5, T_dis_RSP)."""
    out = set(idx)
    for i in idx:
        if not _is_word(seq[i][1].t):
            continue
        k = i
        while k > 0 and _is_word(seq[k - 1][1].t) and seq[k - 1][0] == seq[k][0] - 1:
            k -= 1
            out.add(k)
        k = i
        while k < len(seq) - 1 and _is_word(seq[k + 1][1].t) and seq[k + 1][0] == seq[k][0] + 1:
            k += 1
            out.add(k)
    return out

def _word_span(seq: list[tuple[int, TChar]], a: int, b: int) -> tuple[int, int]:
    """Grow the changed range [a, b) to whole English words / numbers (3235 -> 3300)."""
    def tok(k: int) -> bool:         # letter/digit, or . - / : between two of them (2023.10.16, PM-2)
        if not 0 <= k < len(seq):
            return False
        t = seq[k][1].t
        return _is_word(t) or (t in ".-/:" and 0 < k < len(seq) - 1
                               and _is_word(seq[k - 1][1].t) and _is_word(seq[k + 1][1].t))

    def joined(k: int) -> bool:      # seq[k-1] and seq[k] belong to one word / number / code
        return 0 < k < len(seq) and tok(k - 1) and tok(k) and seq[k][0] == seq[k - 1][0] + 1
    while joined(a):
        a -= 1
    while joined(b):
        b += 1
    return a, b

def _text_details(oc, nc, edits: list[tuple[int, int, int, int]]):
    """One detail per changed word or text run: 「old」 → 「new」, added 「…」, removed 「…」.
    Also returns the characters to box: removed text in old, new/changed text in new."""
    spans: list[list[int]] = []
    for i1, i2, j1, j2 in edits:
        a, b = _word_span(oc, i1, i2)
        c, d = _word_span(nc, j1, j2)
        if spans and a <= spans[-1][1] and c <= spans[-1][3]:
            spans[-1] = [min(spans[-1][0], a), max(spans[-1][1], b), min(spans[-1][2], c), max(spans[-1][3], d)]
        else:
            spans.append([a, b, c, d])
    out: list[str] = []
    old_box: set[int] = set()
    new_box: set[int] = set()
    removed_at: list[int] = []                   # where text was only removed (marker in new)
    for a, b, c, d in spans:
        if a == b:
            out.append(f"added 「{_txt_spaced(nc, c, d)}」")
        elif c == d:
            out.append(f"removed 「{_txt_spaced(oc, a, b)}」")
            old_box.update(range(a, b))
            removed_at.append(c)
        else:
            out.append(f"「{_txt_spaced(oc, a, b)}」 → 「{_txt_spaced(nc, c, d)}」")
        new_box.update(range(c, d))
    return out, old_box, new_box, removed_at
        

def _boxes_for(seq: list[tuple[int, TChar]], idx: set[int]) -> list[Box]:
    out: list[Box] = []
    run: list[TChar] = []
    last = None
    for i in sorted(idx):
        if last is not None and i != last + 1:
            out += _boxes(run)
            run = []
        run.append(seq[i][1])
        last = i
    return out + _boxes(run)

def _txt(seq: list[tuple[int, TChar]], a: int, b: int) -> str:
    return "".join(c.raw or c.t for _, c in seq[a:b])

def _txt_spaced(seq: list[tuple[int, TChar]], a:int, b:int) -> str:
    """Like _txt, with a space where the document had a space or line break (4.1 5.0)."""
    out: list[str] =[]
    for k in range(a,b):
        if k > a and seq[k][0] != seq[k-1][0] + 1:
            out.append(" ")
        out.append(seq[k][1].raw or seq[k][1].t)
    return "".join(out)

# ------------------------------------------------------------------ markers
def _point_marker(seq: list[tuple[int, TChar]], j: int) -> Box | None:
    """Thin vertical marker in the new text where characters were removed."""
    if not seq:
        return None
    c = seq[j - 1][1] if j > 0 else seq[0][1]
    x = c.x1 if j > 0 else c.x0
    return (c.page, x - 1.0, c.top, x + 1.0, c.bottom)

def _gap_marker(N: list[Block], j: int) -> Box | None:
    """Thin horizontal marker where a block used to be (before N[j])."""
    if j < len(N):
        bb = N[j].bboxes()
        if not bb:
            return None
        p = min(bb)
        x0, top, x1, _ = bb[p]
        y = top - 2.0
    elif N:
        bb = N[-1].bboxes()
        if not bb:
            return None
        p = max(bb)
        x0, _, x1, bottom = bb[p]
        y = bottom + 2.0
    else:
        return None
    return (p, x0, y - 1.0, x1, y + 1.0)

# ------------------------------------------------------------------ matching
def _sim(o: Block, n: Block) -> float:
    if _group(o) != _group(n):
        return 0.0
    a, b = _ctext(o), _ctext(n)
    if not a and not b:
        return 1.0
    sm = SequenceMatcher(None, a, b, autojunk=False)
    if sm.real_quick_ratio() < MATCH_MIN or sm.quick_ratio() < MATCH_MIN:
        return 0.0
    r = sm.ratio()
    if o.kind == "cell" and o.col == n.col:
        r = min(1.0, r + COL_BONUS)
    return r

def _pair_range(O: list[Block], N: list[Block], i1: int, i2: int, j1: int, j2: int, sim_fn=_sim) -> list[tuple]:
    n, m = i2 - i1, j2 - j1
    if n == 0:
        return [("ins", j1 + k, i1) for k in range(m)]
    if m == 0:
        return [("del", i1 + k, j1) for k in range(n)]
    if n * m > MAX_PAIR_CELLS:
        return [("del", i1 + k, j1) for k in range(n)] + [("ins", j1 + k, i2) for k in range(m)]

    sim = [[sim_fn(O[i1 + a], N[j1 + b]) for b in range(m)] for a in range(n)]
    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    for a in range(n - 1, -1, -1):
        for b in range(m - 1, -1, -1):
            best = max(dp[a + 1][b], dp[a][b + 1])
            if sim[a][b] >= MATCH_MIN:
                best = max(best, sim[a][b] + dp[a + 1][b + 1])
            dp[a][b] = best

    ops: list[tuple] = []
    a = b = 0
    while a < n and b < m:
        s = sim[a][b]
        if s >= MATCH_MIN and abs(dp[a][b] - (s + dp[a + 1][b + 1])) < 1e-9:
            ops.append(("pair", i1 + a, j1 + b))
            a += 1
            b += 1
        elif abs(dp[a][b] - dp[a + 1][b]) < 1e-9:
            ops.append(("del", i1 + a, j1 + b))
            a += 1
        else:
            ops.append(("ins", j1 + b, i1 + a))
            b += 1
    while a < n:
        ops.append(("del", i1 + a, j2))
        a += 1
    while b < m:
        ops.append(("ins", j1 + b, i2))
        b += 1
    return ops

# ------------------------------------------------------------------ units: paragraph or whole table row
@dataclass
class _Unit:
    """What is lined up between old and new: one paragraph/heading, or one whole table row."""
    idx: list[int]                     # block indices (a row: its cells, left to right)
    cells: dict[int, str] | None       # table row: column -> text; None for a paragraph
    key: tuple[str, str] = ("", "")    # (group, text) - what the unit says
    where: str = ""                    # section heading it is in (keeps repeated rows apart)

def _units(B: list[Block]) -> list[_Unit]:
    out: list[_Unit] = []
    for i, b in enumerate(B):
        last = B[out[-1].idx[0]] if out else None
        if (b.kind == "cell" and last is not None and last.kind == "cell"
                and (last.table, last.row) == (b.table, b.row)):
            out[-1].idx.append(i)
        else:
            out.append(_Unit([i], {} if b.kind == "cell" else None))
    for u in out:
        u.where = B[u.idx[0]].heading
        if u.cells is None:
            u.key = _key(B[u.idx[0]])
        else:
            u.idx.sort(key=lambda i: B[i].col)
            u.cells = {B[i].col: _ctext(B[i]) for i in u.idx}
            u.key = ("row", "|".join(u.cells.values()))
    return out

def _make_unit_sim(O: list[Block], N: list[Block]):
    def sim(uo: _Unit, un: _Unit) -> float:
        if (uo.cells is None) != (un.cells is None):
            return 0.0
        if uo.cells is None:
            return _sim(O[uo.idx[0]], N[un.idx[0]])
        # table rows: whole-row text, or share of cells that are the same in the same column
        same = sum(1 for c, t in uo.cells.items() if t and t == un.cells.get(c))
        by_col = same / max(len(uo.cells), len(un.cells), 1)
        a, b = uo.key[1], un.key[1]
        sm = SequenceMatcher(None, a, b, autojunk=False)
        if sm.real_quick_ratio() < MATCH_MIN and by_col < MATCH_MIN:
            return 0.0
        return max(sm.ratio(), by_col)
    return sim

def _pair_moved(UO: list[_Unit], UN: list[_Unit], uops: list[tuple], sim_fn=None) -> list[tuple]:
    """A deleted unit whose exact text was added somewhere else only moved: pair them.

    Only text that occurs once in each document counts (a repeated table header row
    or caption is not "moved", it belongs to its own table).
    """
    once_old = Counter(u.key for u in UO)
    once_new = Counter(u.key for u in UN)
    added: dict[tuple, list[int]] = {}
    for k, op in enumerate(uops):
        if op[0] == "ins":
            added.setdefault(UN[op[1]].key, []).append(k)
    out = list(uops)
    drop: set[int] = set()
    for k, op in enumerate(uops):
        if op[0] != "del":
            continue
        key = UO[op[1]].key
        if (len(key[1].replace("|", "")) < MOVED_MIN_CHARS or not added.get(key)
                or once_old[key] != 1 or once_new[key] != 1):
            continue
        kk = added[key].pop(0)
        out[k] = ("pair", op[1], uops[kk][1])
        drop.add(kk)

    # moved and edited: leftover deleted/added units that are each other's best close match
    if sim_fn is not None:
        dels = [k for k, op in enumerate(out) if op[0] == "del"]
        adds = [k for k, op in enumerate(out) if op[0] == "ins" and k not in drop]
        if dels and adds and len(dels) * len(adds) <= MAX_PAIR_CELLS:
            score = {(a, b): sim_fn(UO[out[a][1]], UN[out[b][1]]) for a in dels for b in adds}
            for a in dels:
                b = max(adds, key=lambda x: score[(a, x)], default=None)
                if (b is None or b in drop or score[(a, b)] < MOVED_SIM
                        or max(dels, key=lambda x: score[(x, b)]) != a):
                    continue
                out[a] = ("pair", out[a][1], out[b][1])
                drop.add(b)
    return [op for k, op in enumerate(out) if k not in drop]

def _expand_pair(O: list[Block], N: list[Block], uo: _Unit, un: _Unit) -> list[tuple]:
    """Paired units -> paired blocks. Rows: cells lined up by text, then left to right."""
    if uo.cells is None:
        return [("pair", uo.idx[0], un.idx[0])]
    oi, ni = uo.idx, un.idx
    sm = SequenceMatcher(None, [_ctext(O[i]) for i in oi], [_ctext(N[j]) for j in ni], autojunk=False)
    ops: list[tuple] = []

    def nxt_new(j: int) -> int:
        return ni[j] if j < len(ni) else ni[-1] + 1

    def nxt_old(i: int) -> int:
        return oi[i] if i < len(oi) else oi[-1] + 1

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        k = 0
        if tag in ("equal", "replace"):
            for k in range(min(i2 - i1, j2 - j1)):       # same place in the row = Modified
                ops.append(("pair", oi[i1 + k], ni[j1 + k]))
            k = min(i2 - i1, j2 - j1)
        ops += [("del", oi[i], nxt_new(j1 + k)) for i in range(i1 + k, i2)]
        ops += [("ins", ni[j], nxt_old(i1 + k)) for j in range(j1 + k, j2)]
    return ops

def _align(O: list[Block], N: list[Block]) -> list[tuple]:
    """("pair", i, j) | ("del", i, next_new_index) | ("ins", j, next_old_index)."""
    UO, UN = _units(O), _units(N)
    sm = SequenceMatcher(None, [(u.where, u.key) for u in UO], [(u.where, u.key) for u in UN],
                         autojunk=False)
    uops: list[tuple] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            uops += [("pair", i1 + k, j1 + k) for k in range(i2 - i1)]
        else:
            uops += _pair_range(UO, UN, i1, i2, j1, j2, _make_unit_sim(O, N))
    uops = _pair_moved(UO, UN, uops, _make_unit_sim(O, N))

    ops: list[tuple] = []
    for op in uops:
        if op[0] == "pair":
            ops += _expand_pair(O, N, UO[op[1]], UN[op[2]])
        elif op[0] == "del":
            nxt = UN[op[2]].idx[0] if op[2] < len(UN) else len(N)
            ops += [("del", i, nxt) for i in UO[op[1]].idx]
        else:
            nxt = UO[op[2]].idx[0] if op[2] < len(UO) else len(O)
            ops += [("ins", j, nxt) for j in UN[op[1]].idx]
    return ops


# ------------------------------------------------------------------ formatting
_COLOR_NAMES = {
    "black":(0, 0, 0), "white":(1, 1, 1), "gray":(0.5, 0.5, 0.5), "red":(1, 0, 0),
    "dark red":(0.75, 0, 0), "orange":(1, 0.65, 0), "yellow":(1, 1, 0), "green":(0, 0.69, 0.31),
    "dark green": (0, 0.5, 0), "cyan": (0, 1, 1), "blue": (0, 0, 1), "dark blue": (0, 0, 0.5),
    "light blue": (0, 0.69, 0.94), "purple": (0.44, 0.19, 0.63), "magenta": (1, 0, 1),   
}

_STYLE_STATS: Counter = Counter()
_IGNORED:set[str] = set()

def _rgb(c: tuple | None) -> tuple[float, float, float] | None:
    """Gray / RGB / CMYK colour -> RGB (0-1). None if unknown."""
    if not c:
        return None
    try:
        v = [min(1.0, max(0.0, float(x))) for x in c]
    except (TypeError, ValueError):
        return None
    if len(v) == 1:
        return (v[0], v[0], v[0])
    if len(v) == 3:
        return (v[0], v[1], v[2])
    if len(v) == 4:
        k = v[3]
        return tuple((1 - x) * (1-k) for x in v[:3])
    return None

def _color_name(rgb: tuple[float, float, float]) -> str:
    name = min(_COLOR_NAMES, key=lambda k: sum((a - b) ** 2 for a, b in zip(_COLOR_NAMES[k], rgb)))
    hexv = "#" + "".join(f"{round(x * 255):02X}" for x in rgb)
    return f"{name} {hexv}"

def _size(v:float) -> str:
    return f"{v:.1f}".rstrip("0").rstrip(".")

def _format_diff(a: TChar, b: TChar) -> tuple[str, ...]:
    """What looks different between two equal characters (empty = nothing)."""
    out: list[str] = []
    if a.struck != b.struck:
        out.append(f"strikethrough {'removed' if a.struck else 'added'}")
    for attr in ("bold", "italic", "underline"):
        if attr not in _IGNORED and getattr(a, attr) != getattr(b, attr):
            out.append(f"{attr} {'removed' if getattr(a, attr) else 'added'}")
    if "colour" not in _IGNORED:
        ca, cb = _rgb(a.color), _rgb(b.color)
        if ca and cb and max(abs(x - y) for x, y in zip(ca, cb)) > COLOR_TOL:
            out.append(f"font colour {_color_name(ca)} → {_color_name(cb)}")
    if "size" not in _IGNORED and a.size and b.size and abs(a.size - b.size) > SIZE_TOL:
        out.append(f"font size {_size(a.size)} → {_size(b.size)}")
    return tuple(out)

def _count_styles(diff: tuple[str, ...]) -> None:
    _STYLE_STATS["chars"] += 1
    for d in diff:
        _STYLE_STATS[_attr_of(d)] += 1

def _attr_of(detail: str) -> str:
    for attr in ("strikethrough", "bold", "italic", "underline", "colour", "size"):
        if attr in detail:
            return attr
    return detail

# ------------------------------------------------------------------ character diff
def _diff_pair(o:Block, n:Block):
    un = _unnumbered(o)
    if un is not None and un == _unnumbered(n):
        return None
    oc, nc = _compact(o), _compact(n)
    sm = SequenceMatcher(None, [c.t for _, c in oc], [c.t for _, c in nc], autojunk=False)
    if oc and nc and sm.ratio() < MATCH_MIN:
        # completely different text in the same place ("Released" -> "DRAFT"): one replacement
        return ([], _boxes_for(nc, set(range(len(nc)))), [],
                [f"「{_txt(oc, 0, len(oc))}」 → 「{_txt(nc, 0, len(nc))}」"])     
    old_idx: set[int] = set()
    new_idx: set[int] = set()
    markers: list[Box] = []
    details: list[str] = []
    edits:list[tuple[int, int, int, int]] = [] # text changes, described after the loop
    fmt_runs:list[tuple[int, int, tuple]] = [] # formatting changes (new start, end, what)

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            # formatting: one detail line per run of characters with the same difference
            run: list[int] = []
            run_diff: tuple[str, ...] = ()
            for k in range (i2 - i1):
                diff = _format_diff(oc[i1 + k][1], nc[j1 + k][1])
                _count_styles(diff)
                if run and (not diff or diff != run_diff or j1 + k != run[-1] + 1):
                    fmt_runs.append((run[0], run[-1] + 1, run_diff))
                    run = []
                if diff:
                    run.append(j1 + k)
                    run_diff = diff
                    new_idx.add(j1 + k)
            if run:
                fmt_runs.append((run[0], run[-1] + 1, run_diff))
        elif tag in ("replace", "insert"):
            edits.append((i1, i2, j1, j2))
        else:  # delete
            edits.append((i1, i2, j1, j2))

    text, o_box, n_box, removed_at = _text_details(oc, nc, edits)
    for j in removed_at:
        m = _point_marker(nc, j)
        if m:
            markers.append(m)
    for a, b, what in fmt_runs:
        if all(k in n_box for k in range(a, b)):
            new_idx.difference_update(range(a, b))    # inside a changed word: already reported
            continue
        details.append(f"{', '.join(what)} 「{_txt(nc, a, b)}」")
    details += text
    old_idx |= o_box
    new_idx |= n_box

    if not old_idx and not new_idx:
        return None
    return (_boxes_for(oc, _expand_words(oc, old_idx)),
            _boxes_for(nc, _expand_words(nc, new_idx)),
            markers, details)

# ------------------------------------------------------------------ table rows
def _row_boxes(blocks: list[Block], key: tuple) -> list[Box]:
    """Area of one table row from its text (all its cells), one box per page."""
    out: dict[int, tuple[float, float, float, float]] = {}
    for b in blocks:
        if b.kind == "cell" and (b.table, b.row) == key:
            for p, (x0, t, x1, bt) in b.bboxes().items():
                r = out.get(p)
                out[p] = (x0, t, x1, bt) if r is None else (
                    min(r[0], x0), min(r[1], t), max(r[2], x1), max(r[3], bt))
    return [(p, *v) for p, v in sorted(out.items())]

def _partner_row(src: list[Block], idx_map: dict[int, int], dst: list[Block],
                 b: Block) -> tuple | None:
    """Same table row in the other document.

    1) a cell of the same row exists in both documents -> use its row
    2) otherwise any matched cell of the same table -> same row number,
       shifted by that cell's row difference (handles added/removed rows above)
    """
    fallback = None
    for si, di in idx_map.items():
        s, d = src[si], dst[di]
        if s.kind != "cell" or d.kind != "cell" or s.table != b.table:
            continue
        if s.row == b.row:
            return (d.table, d.row)
        if fallback is None or abs(s.row - b.row) < abs(fallback[0] - b.row):
            fallback = (s.row, d.table, d.row - s.row)
    if fallback:
        return (fallback[1], b.row + fallback[2])
    return None

def _row_area(blocks: list[Block], table_rows: dict | None, key: tuple) -> list[Box]:
    """Area of a table row: from the table grid (works for empty rows), else from its text."""
    g = table_rows.get(key) if table_rows else None
    if g:
        return [g]
    return _row_boxes(blocks, key)

# ------------------------------------------------------------------ changes
def _row_key(c: Change):
    b = (c.new_blocks or c.old_blocks)[0]
    return (b.table, b.row) if b.kind == "cell" else None

def _first_kind(c: Change) -> str:
    return (c.old_blocks or c.new_blocks)[0].kind

def _merge(changes: list[Change]) -> list[Change]:
    """One row per table row; a deleted/new section (heading + content) = one row."""
    out: list[Change] = []
    for c in changes:
        p = out[-1] if out else None
        if p is not None:
            rk = _row_key(c)
            same_row = rk is not None and rk == _row_key(p)
            section = (c.kind == p.kind and c.kind in ("Deleted", "New")
                       and _first_kind(p) == "heading" and _first_kind(c) != "heading"
                       and c.first_seq == p.last_seq + 1)
            if same_row or section:
                if p.kind != c.kind:
                    p.kind = "Modified"
                p.old_blocks += c.old_blocks
                p.new_blocks += c.new_blocks
                p.old_boxes += c.old_boxes
                p.new_boxes += c.new_boxes
                p.new_markers += c.new_markers
                p.old_anchor += c.old_anchor
                p.new_anchor += c.new_anchor
                p.details += c.details
                p.last_seq = c.last_seq
                continue
        out.append(c)
    return out

def _changes_for(O: list[Block], N: list[Block],
                 old_rows: dict | None = None, new_rows: dict | None = None) -> list[Change]:
    ops = _align(O, N)
    n2o = {op[2]: op[1] for op in ops if op[0] == "pair"}
    o2n = {op[1]: op[2] for op in ops if op[0] == "pair"}
    out: list[Change] = []
    for seq, op in enumerate(ops):
        if op[0] == "pair":
            o, n = O[op[1]], N[op[2]]
            d = _diff_pair(o, n)
            if d:
                ob, nb, mk, det = d
                out.append(Change("Modified", n.heading, [o], [n], ob, nb, mk, det,
                                  first_seq=seq, last_seq=seq))
        elif op[0] == "del":
            o = O[op[1]]
            ch = Change("Deleted", o.heading, [o], [], _boxes(o.chars), [], [],
                        [f"deleted 「{_short(o)}」"], first_seq=seq, last_seq=seq)
            key = _partner_row(O, o2n, N, o) if o.kind == "cell" else None
            row = _row_area(N, new_rows, key) if key else []
            if row:
                ch.new_anchor = row
                p, x0, top, x1, _ = row[0]
                ch.new_markers = [(p, x0, top - 3.0, x1, top - 1.0)]
            else:
                mk = _gap_marker(N, op[2])
                if mk:
                    ch.new_markers = [mk]
            out.append(ch)
        else:
            n = N[op[1]]
            ch = Change("New", n.heading, [], [n], [], _boxes(n.chars), [],
                        [f"added 「{_short(n)}」"], first_seq=seq, last_seq=seq)
            key = _partner_row(N, n2o, O, n) if n.kind == "cell" else None
            row = _row_area(O, old_rows, key) if key else []
            if row:
                ch.old_anchor = row
            else:
                anchor = _gap_marker(O, op[2])
                if anchor:
                    ch.old_anchor = [anchor]
            out.append(ch)
    return _merge(out)

def _dedupe(blocks: list[Block]) -> list[Block]:
    """Headers/footers repeat on every page - keep the first occurrence of each text."""
    seen: set[tuple[str, str]] = set()
    out: list[Block] = []
    for b in blocks:
        k = (b.heading, _ctext(b))
        if k not in seen:
            seen.add(k)
            out.append(b)
    return out

def _first_box(*box_lists: list[Box]):
    for boxes in box_lists:
        if boxes:
            p, _, t, _, _ = boxes[0]
            return (p, t)
    return None

def _pos(c: Change) -> tuple:
    """Position used to number changes in document order."""
    if c.new_blocks:
        return min((b.page, b.top) for b in c.new_blocks)
    p = _first_box(c.new_boxes, c.new_markers, c.new_anchor)
    if p:
        return p
    if c.old_blocks:
        return min((b.page, b.top) for b in c.old_blocks)
    return _first_box(c.old_boxes, c.old_anchor) or (10 ** 9, 0.0)

TOC_LABEL = "目次"
_TOC_RE = re.compile(r"(?:\.{4,}|…{2,}|・{4,}|‥{2,})\s*\d+$")

def _is_toc(c: Change) -> bool:
    blocks = c.new_blocks or c.old_blocks
    return bool(blocks) and all(b.kind == "para" and _TOC_RE.search(_ctext(b)) for b in blocks)

def _group_toc(changes: list[Change]) -> list[Change]:
    """Table of contents: all its changes on one page = one row (page numbers shift a lot)."""
    out: list[Change] = []
    for c in changes:
        if not _is_toc(c):
            out.append(c)
            continue
        page = (c.new_blocks or c.old_blocks)[0].page
        p = out[-1] if out else None
        if p is not None and p.heading == TOC_LABEL and (p.new_blocks or p.old_blocks)[0].page == page:
            if p.kind != c.kind:
                p.kind = "Modified"
            p.old_blocks += c.old_blocks
            p.new_blocks += c.new_blocks
            p.old_boxes += c.old_boxes
            p.new_boxes += c.new_boxes
            p.new_markers += c.new_markers
            p.old_anchor += c.old_anchor
            p.new_anchor += c.new_anchor
            p.details += c.details
            continue
        c.heading = TOC_LABEL
        out.append(c)
    return out

def _fold_images(body:list[Change], images: list[Change]) -> list[Change]:
    """A picture deleted/added together with its whole section joins that section's row."""
    sections = {(c.kind, c.heading):c for c in body 
            if c.kind in ("Deleted", "New") and _first_kind(c) == "heading"}
    rest: list[Change] =[]
    for im in images:
        sec = sections.get((im.kind, im.heading))
        if sec is None:
            rest.append(im)
            continue
        sec.old_boxes += im.old_boxes
        sec.new_boxes += im.new_boxes
        sec.details += im.details
    return rest

def _systematic() -> set[str]:
    """Format attributes that differ almost everywhere: the two PDFs were made differently
    (e.g . Word vs supplier tool), so these differnces are not real changes."""
    total = _STYLE_STATS["chars"]
    if total < SYSTEMATIC_MIN:
        return set()
    return {a for a in ("bold", "italic", "underline", "colour", "size")
            if _STYLE_STATS[a] > total * SYSTEMATIC_RATIO}

def compare(old: CleanDoc, new: CleanDoc) -> list[Change]:
    from cpa.images import image_changes  # imported here to avoid a circular import

    old_rows = getattr(old, "table_rows", None) or {}
    new_rows = getattr(new, "table_rows", None) or {}
    ob = [b for b in old.blocks if b.kind != "hf"]
    nb = [b for b in new.blocks if b.kind != "hf"]
    _STYLE_STATS.clear()
    _IGNORED.clear()
    body = _changes_for(ob, nb, old_rows, new_rows)
    skip = _systematic()
    if skip:
        print(f"WARNING: {', '.join(sorted(skip))} differ(s) on most of the text "
              f"(PDFs made by different tools?) - these format changes are ignored; "
              f"check by hand if the whole document's format really changed")
        _IGNORED.update(skip)
        body = _changes_for(ob, nb, old_rows, new_rows)
    hf = _changes_for(_dedupe([b for b in old.blocks if b.kind == "hf"]),
                      _dedupe([b for b in new.blocks if b.kind == "hf"]))
    img_body, img_hf = image_changes(old, new)
    img_body = _fold_images(body, img_body)
    changes = _group_toc(sorted(body + img_body, key = _pos)) + sorted(hf + img_hf, key=_pos)
    for i, c in enumerate(changes, start=1):
        c.no = i
    return changes

# ------------------------------------------------------------------ report
def _pages(blocks: list[Block], *box_lists: list[Box]) -> str:
    ps = {p for b in blocks for p in b.bboxes()}
    ps |= {bx[0] for boxes in box_lists for bx in boxes}
    return "p" + ",".join(map(str, sorted(ps))) if ps else "-"

def _areas(boxes: list[Box]) -> str:
    return ", ".join(f"p{p} y={t:.0f}-{b:.0f}" for p, _, t, _, b in boxes) or "-"

def write_changes(changes: list[Change], old: CleanDoc, new: CleanDoc, path: Path) -> None:
    out = [f"# {old.source.name} → {new.source.name}   ({len(changes)} changes)   "
           f"~~text~~ = struck through"]
    for c in changes:
        loc = (f"old {_pages(c.old_blocks, c.old_boxes, c.old_anchor)} → "
               f"new {_pages(c.new_blocks, c.new_boxes, c.new_markers, c.new_anchor)}")
        blocks = c.new_blocks or c.old_blocks
        if blocks and blocks[0].kind == "cell":
            loc += f" | table T{blocks[0].table} row {blocks[0].row}"
        elif not blocks:
            loc += " | image"
        out.append(f"\n[{c.no:02d}] {c.kind} | {c.heading} | {loc}")
        if c.old_blocks:
            out.append("     old: " + " | ".join(_marked(x) for x in c.old_blocks))
        if c.new_blocks:
            out.append("     new: " + " | ".join(_marked(x) for x in c.new_blocks))
        for d in c.details:
            out.append(f"     - {d}")
        out.append(f"     boxes: old {len(c.old_boxes)}, new {len(c.new_boxes)}, "
                   f"markers {len(c.new_markers)}")
        if c.old_anchor or c.new_anchor:
            out.append(f"     crop area: old {_areas(c.old_anchor)} | new {_areas(c.new_anchor)}")
    path.write_text("\n".join(out), encoding="utf-8")    


def main() -> None:
    ap = argparse.ArgumentParser(description="Compare old and new documents (Step 1c).")
    ap.add_argument("old", help="old .pdf or .docx")
    ap.add_argument("new", help="new .pdf or .docx")
    ap.add_argument("--out", default="output", help="output folder (default: output)")
    args = ap.parse_args()

    out_dir = Path(args.out)
    old = load(args.old, out_dir)
    new = load(args.new, out_dir)
    changes = compare(old, new)
    path = out_dir / "changes.txt"
    write_changes(changes, old, new, path)
    kinds = {k: sum(1 for c in changes if c.kind == k) for k in ("New", "Deleted", "Modified")}
    print(f"{len(changes)} changes (New {kinds['New']}, Deleted {kinds['Deleted']}, "
          f"Modified {kinds['Modified']}) -> {path}")

if __name__ == "__main__":
    main()
