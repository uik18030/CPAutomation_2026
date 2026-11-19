"""Step 1f: check the tool's result against the answer key.

Run (from the repo folder, with .venv active):
    python tests/check_answer_key.py
Options:
    --old / --new   documents to compare (default: the test pair in tests/data)
    --key           answer key JSON (default: tests/expected/answer_key.json)
    --out           output folder (default: output/check)
Writes <out>/answer_key_report.txt (UTF-8). Exit code 0 = no failures.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from cpa.clean import load, norm_char            # noqa: E402
from cpa.compare import Change, _ctext, compare  # noqa: E402

PENDING_WORDS = ()       # not detected before Step 2

def norm(s: str) -> str:
    """Same normalising as the tool (full/half width, punctuation), spaces removed."""
    return "".join(k for ch in str(s) for k in norm_char(ch) if not k.isspace())

def texts(c: Change) -> tuple[str, str, str]:
    old = "".join(_ctext(b) for b in c.old_blocks)
    new = "".join(_ctext(b) for b in c.new_blocks)
    return old, new, " ".join(c.details)

def is_pending(exp: dict) -> bool:
    if exp.get("old") or exp.get("new") or exp.get("text"):
        return False
    ch = str(exp.get("change", "")).lower()
    return any(w in ch for w in PENDING_WORDS)

def matches(exp: dict, c: Change) -> bool:
    if exp.get("header") and exp["header"] != c.heading:
        return False
    if exp.get("type") and exp["type"] != c.kind:
        return False
    old_t, new_t, det = texts(c)
    o, n = norm(exp.get("old", "")), norm(exp.get("new", ""))
    if o and o not in old_t:
        return False
    if n and n not in new_t:
        return False
    if n and not o and n in old_t:        # "added" text must not already be in old
        return False
    ch = exp.get("change")
    t = norm(exp.get("test", ""))
    if t:
        #formatting change: one detail line must name both the change and the text
        if not any(str(ch or "").lower() in d.lower() and t in norm(d) for d in c.details):
            return False
        elif ch and not (o or n) and str(ch).lower() not in det.lower():
            return False
    return True

def describe(exp: dict) -> str:
    parts = [f"#{exp.get('no', '?')}", str(exp.get("header", "")), str(exp.get("type", ""))]
    if exp.get("old") or exp.get("new"):
        parts.append(f"「{exp.get('old', '')}」→「{exp.get('new', '')}」")
    if exp.get("text"):
        parts.append(f"「{exp['text']}」")
    if exp.get("change"):
        parts.append(f"({exp['change']})")
    return " ".join(parts)

def main() -> int:
    ap = argparse.ArgumentParser(description="Check results against the answer key (Step 1f).")
    ap.add_argument("--old", default=str(ROOT / "tests/data/File_Old_update.docx"))
    ap.add_argument("--new", default=str(ROOT / "tests/data/File_New_update.docx"))
    ap.add_argument("--key", default=str(ROOT / "tests/expected/answer_key.json"))
    ap.add_argument("--out", default=str(ROOT / "output/check"))
    args = ap.parse_args()

    expected = json.loads(Path(args.key).read_text(encoding="utf-8-sig"))
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    old = load(args.old, out_dir)
    new = load(args.new, out_dir)
    changes = compare(old, new)

    used: set[int] = set()
    report: list[str] = [f"# {Path(args.old).name} → {Path(args.new).name}",
                         f"# answer key: {Path(args.key).name}", ""]
    terminal: list[str] = []
    passed = pending = failed = 0

    for exp in expected:
        no = exp.get("no", "?")
        hit = next((c for c in changes if c.no not in used and matches(exp, c)), None)
        if hit:
            used.add(hit.no)
            passed += 1
            report.append(f"PASS     {describe(exp)}  -> change [{hit.no:02d}]")
            terminal.append(f"PASS     #{no} -> change [{hit.no:02d}]")
        elif is_pending(exp):
            pending += 1
            report.append(f"PENDING  {describe(exp)}  (image changes come in Step 2)")
            terminal.append(f"PENDING  #{no} (Step 2)")
        else:
            failed += 1
            report.append(f"FAIL     {describe(exp)}  -> not found")
            terminal.append(f"FAIL     #{no} -> not found")

    extra = [c for c in changes if c.no not in used]
    report.append("")
    if extra:
        failed += len(extra)
        report.append(f"UNEXPECTED changes (not in the answer key): {len(extra)}")
        for c in extra:
            report.append(f"  [{c.no:02d}] {c.kind} | {c.heading} | " + " / ".join(c.details))
            terminal.append(f"UNEXPECTED change [{c.no:02d}]")
    else:
        report.append("Unexpected changes: none")

    summary = (f"Result: {passed} passed, {pending} pending, {failed} failed  "
               f"({len(changes)} changes found)")
    report += ["", summary]
    path = out_dir / "answer_key_report.txt"
    path.write_text("\n".join(report), encoding="utf-8")

    print("\n".join(terminal))
    print(f"{summary} -> {path}")
    return 0 if failed == 0 else 1

if __name__ == "__main__":
    sys.exit(main())
