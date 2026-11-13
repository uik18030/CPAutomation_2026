"""Step 1c: compare old and new documents block by block, then character by character.

Run:
    python -m cpa.compare tests/data/File_Old_update.docx tests/data/File_New_update.docx
Writes output/changes.txt (UTF-8).
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path

from cpa.clean import Block, CleanDoc, TChar, _marked, load

# ------------------------------------------------------------------ settings
MATCH_MIN = 0.5          # blocks at least this similar are paired as "Modified"
COL_BONUS = 0.1          # table cells in the same column are preferred
MAX_PAIR_CELLS = 250_000 # safety limit for the pairing table (n_old x n_new)

Box = tuple[int, float, float, float, float]  # page, x0, top, x1, bottom

@dataclass
class Change:
    kind: str                                   # "New" | "Deleted" | "Modified"
    heading: str                                # Sheet/Header column
    old_blocks: list[Block] = field(default_factory=list)
    new_blocks: list[Block] = field(default_factory=list)
    old_boxes: list[Box] = field(default_factory=list)    # red boxes on Old Spec
    new_boxes: list[Box] = field(default_factory=list)    # red boxes on Latest Spec
    new_markers: list[Box] = field(default_factory=list)  # red markers on Latest Spec (deletions)
    details: list[str] = field(default_factory=list)
    no: int = 0
    first_seq: int = 0
    last_seq: int = 0
    old_anchor: list[Box] = field(default_factory=list)  # red boxes on Old Spec for anchors (deleted blocks)

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
    return "text"                               # headings and paragraphs

def _key(b: Block) -> tuple[str, str]:
    return (_group(b), _ctext(b))

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

# ------------------------------------------------------------------ markers
def _point_marker(seq: list[tuple[int, TChar]], j: int) -> Box | None:
    """Thin vertical marker in the new text where characters were removed."""
    if not seq:
        return None
    c = seq[j - 1][1] if j > 0 else seq[0][1]
    x = c.x1 if j > 0 else c.x0
    return (c.page, x - 1.0, c.top, x + 1.0, c.bottom)

def _gap_marker(N: list[Block], j: int) -> Box | None:
    """Thin horizontal marker where a deleted block used to be (before N[j])."""
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

def _pair_range(O: list[Block], N: list[Block], i1: int, i2: int, j1: int, j2: int) -> list[tuple]:
    """Pair similar blocks inside an unmatched stretch (keeps document order)."""
    n, m = i2 - i1, j2 - j1
    if n == 0:
        return [("ins", j1 + k, i1) for k in range(m)]
    if m == 0:
        return [("del", i1 + k, j1) for k in range(n)]
    if n * m > MAX_PAIR_CELLS:
        return [("del", i1 + k, j1) for k in range(n)] + [("ins", j1 + k, i2) for k in range(m)]

    sim = [[_sim(O[i1 + a], N[j1 + b]) for b in range(m)] for a in range(n)]
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

def _align(O: list[Block], N: list[Block]) -> list[tuple]:
    """("pair", i, j) | ("del", i, next_new_index) | ("ins", j, next_old_index)."""
    sm = SequenceMatcher(None, [_key(b) for b in O], [_key(b) for b in N], autojunk=False)
    ops: list[tuple] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            ops += [("pair", i1 + k, j1 + k) for k in range(i2 - i1)]
        else:
            ops += _pair_range(O, N, i1, i2, j1, j2)
    return ops

# ------------------------------------------------------------------ character diff
def _diff_pair(o: Block, n: Block):
    oc, nc = _compact(o), _compact(n)
    sm = SequenceMatcher(None, [c.t for _, c in oc], [c.t for _, c in nc], autojunk=False)
    old_idx: set[int] = set()
    new_idx: set[int] = set()
    markers: list[Box] = []
    details: list[str] = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            run: list[int] = []
            run_removed = None
            for k in range(i2 - i1):
                a, b = oc[i1 + k][1], nc[j1 + k][1]
                if a.struck != b.struck:
                    if run and (j1 + k != run[-1] + 1 or a.struck != run_removed):
                        details.append(f"strikethrough {'removed' if run_removed else 'added'} "
                                       f"「{_txt(nc, run[0], run[-1] + 1)}」")
                        run = []
                    run.append(j1 + k)
                    run_removed = a.struck
                    new_idx.add(j1 + k)
            if run:
                details.append(f"strikethrough {'removed' if run_removed else 'added'} "
                               f"「{_txt(nc, run[0], run[-1] + 1)}」")
        elif tag == "replace":
            new_idx.update(range(j1, j2))
            details.append(f"「{_txt(oc, i1, i2)}」 → 「{_txt(nc, j1, j2)}」")
        elif tag == "insert":
            new_idx.update(range(j1, j2))
            details.append(f"added 「{_txt(nc, j1, j2)}」")
        else:  # delete
            old_idx.update(range(i1, i2))
            details.append(f"removed 「{_txt(oc, i1, i2)}」")
            m = _point_marker(nc, j1)
            if m:
                markers.append(m)

    if not old_idx and not new_idx:
        return None
    return (_boxes_for(oc, _expand_words(oc, old_idx)),
            _boxes_for(nc, _expand_words(nc, new_idx)),
            markers, details)

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
                p.details += c.details
                p.last_seq = c.last_seq
                continue
        out.append(c)
    return out

def _changes_for(O: list[Block], N: list[Block]) -> list[Change]:
    out: list[Change] = []
    for seq, op in enumerate(_align(O, N)):
        if op[0] == "pair":
            o, n = O[op[1]], N[op[2]]
            d = _diff_pair(o, n)
            if d:
                ob, nb, mk, det = d
                out.append(Change("Modified", n.heading, [o], [n], ob, nb, mk, det,
                                  first_seq=seq, last_seq=seq))
        elif op[0] == "del":
            o = O[op[1]]
            mk = _gap_marker(N, op[2])
            out.append(Change("Deleted", o.heading, [o], [], _boxes(o.chars), [],
                              [mk] if mk else [], [f"deleted 「{_short(o)}」"],
                              first_seq=seq, last_seq=seq))
        else:
            n = N[op[1]]
            anchor = _gap_marker(O, op[2])
            ch = Change("New", n.heading, [], [n], [], _boxes(n.chars), [],
                            [f"added 「{_short(n)}」"], first_seq=seq, last_seq=seq)
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

def compare(old: CleanDoc, new: CleanDoc) -> list[Change]:
    ob = [b for b in old.blocks if b.kind != "hf"]
    nb = [b for b in new.blocks if b.kind != "hf"]
    changes = _changes_for(ob, nb)
    changes += _changes_for(_dedupe([b for b in old.blocks if b.kind == "hf"]),
                            _dedupe([b for b in new.blocks if b.kind == "hf"]))
    for i, c in enumerate(changes, start=1):
        c.no = i
    return changes

# ------------------------------------------------------------------ report
def _pages(blocks: list[Block]) -> str:
    ps = sorted({p for b in blocks for p in b.bboxes()})
    return "p" + ",".join(map(str, ps)) if ps else "-"

def write_changes(changes: list[Change], old: CleanDoc, new: CleanDoc, path: Path) -> None:
    out = [f"# {old.source.name} → {new.source.name}   ({len(changes)} changes)   "
           f"~~text~~ = struck through"]
    for c in changes:
        loc = f"old {_pages(c.old_blocks)} → new {_pages(c.new_blocks)}"
        b = (c.new_blocks or c.old_blocks)[0]
        if b.kind == "cell":
            loc += f" | table T{b.table} row {b.row}"
        out.append(f"\n[{c.no:02d}] {c.kind} | {c.heading} | {loc}")
        if c.old_blocks:
            out.append("     old: " + " | ".join(_marked(x) for x in c.old_blocks))
        if c.new_blocks:
            out.append("     new: " + " | ".join(_marked(x) for x in c.new_blocks))
        for d in c.details:
            out.append(f"     - {d}")
        out.append(f"     boxes: old {len(c.old_boxes)}, new {len(c.new_boxes)}, "
                   f"markers {len(c.new_markers)}")
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
