"""Draw a picture of the CPA_JPN sheet (A, B, C, F, G) from changes.txt + shots, like Excel shows it."""
import re, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

out = Path(sys.argv[1]); rows_wanted = sys.argv[2] if len(sys.argv) > 2 else ""
dest = Path(sys.argv[3]) if len(sys.argv) > 3 else out / "sheet_view.png"
F = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
font, bold = ImageFont.truetype(F, 18), ImageFont.truetype(F, 19)
rows = [(int(m[1]), m[2], m[3]) for m in re.finditer(r"^\[(\d+)\] (\w+) \| ([^|]+?) \|", (out / "changes.txt").read_text(encoding="utf-8"), re.M)]
if rows_wanted:
    keep = set()
    for part in rows_wanted.split(","):
        a, _, b = part.partition("-"); keep.update(range(int(a), int(b or a) + 1))
    rows = [r for r in rows if r[0] in keep]
W = {"A": 70, "B": 230, "C": 130, "F": 620, "G": 620}
X = {}; x = 0
for k in W: X[k] = x; x += W[k]
HDR = 46
heights = []
for no, kind, head in rows:
    h = 40
    for side in ("old", "new"):
        p = out / "shots" / f"{no:02d}_{side}.png"
        if p.exists():
            with Image.open(p) as im:
                h = max(h, int(im.height * (W["F"] - 10) / im.width) + 10)
    heights.append(min(h, 545))
img = Image.new("RGB", (x + 1, HDR + sum(heights) + 1), "white"); d = ImageDraw.Draw(img)
d.rectangle([0, 0, x, HDR], fill=(252, 228, 214))
for k, t in zip(W, ["No.", "Sheet/Header", "Description", "Old Spec", "Latest Spec"]):
    d.text((X[k] + 8, 13), t, font=bold, fill="black"); d.line([X[k], 0, X[k], img.height], fill=(120, 120, 120))
d.line([x, 0, x, img.height], fill=(120, 120, 120))
y = HDR
for (no, kind, head), h in zip(rows, heights):
    d.line([0, y, x, y], fill=(120, 120, 120))
    d.text((X["A"] + 8, y + 8), str(no), font=font, fill="black")
    d.text((X["B"] + 8, y + 8), head.strip(), font=font, fill="black")
    d.text((X["C"] + 8, y + 8), kind, font=font, fill="black")
    for side, col in (("old", "F"), ("new", "G")):
        p = out / "shots" / f"{no:02d}_{side}.png"
        if p.exists():
            with Image.open(p) as im:
                s = (W[col] - 10) / im.width
                pic = im.convert("RGB").resize((W[col] - 10, max(1, int(im.height * s))))
                pic = pic.crop((0, 0, pic.width, min(pic.height, h - 10)))
                img.paste(pic, (X[col] + 5, y + 5))
        elif side == "old" and kind == "New":
            d.text((X[col] + 8, y + 8), "N/A", font=font, fill="black")
    y += h
d.line([0, y, x, y], fill=(120, 120, 120))
img.save(dest)
print(dest, img.size)
