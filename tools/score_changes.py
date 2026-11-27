import json, re, sys
from pathlib import Path
exp = json.load(open(sys.argv[1], encoding="utf-8"))
txt = Path(sys.argv[2]).read_text(encoding="utf-8")
entries = re.split(r"\n(?=\[\d+\] )", txt)[1:]
ch = []
for e in entries:
    m = re.match(r"\[(\d+)\] (\w+) \| ([^|]+?) \|", e)
    ch.append((int(m[1]), m[2], m[3].strip(), e))
used = set(); found = []; missed = []
for kind, tok in sorted(exp, key=lambda x: -len(x[1])):
    hit = next((c for c in ch if c[0] not in used and c[1] == kind and tok in c[3]), None)
    if hit: used.add(hit[0]); found.append((kind, tok, hit[0]))
    else: missed.append((kind, tok))
extra = [c for c in ch if c[0] not in used]
print(f"expected {len(exp)} | found {len(found)} | missed {len(missed)} | extra (false) {len(extra)} | total rows {len(ch)}")
for k, t in missed: print("  MISSED", k, t)
for c in extra: print("  EXTRA ", f"[{c[0]:02d}]", c[1], c[2], "|", " / ".join(l.strip()[2:] for l in c[3].splitlines() if l.strip().startswith("- "))[:150])
