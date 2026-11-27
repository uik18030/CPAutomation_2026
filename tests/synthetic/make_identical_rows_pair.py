"""Table with several identical rows: one deleted (a), one changed (b), one added (c)."""
import os
from pathlib import Path

# output folder: <repo>/output/synthetic/<script name>/   (git-ignored)
OUT = Path(__file__).resolve().parents[2] / "output" / "synthetic" / Path(__file__).stem
OUT.mkdir(parents=True, exist_ok=True)
os.chdir(OUT)
FONT = os.environ.get("CPA_TEST_FONT", "MS Gothic" if os.name == "nt" else "IPAGothic")
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

SAME = ("-", "Reserved", "$0")
def build(path, rows):
    d = Document(); st = d.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    d.add_paragraph("1. 信号一覧"); d.add_paragraph("表1 信号一覧")
    t = d.add_table(rows=len(rows) + 1, cols=3); t.style = "Table Grid"
    for c, h in enumerate(("No.", "信号", "値")): t.cell(0, c).text = h
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row): t.cell(r, c).text = v
    d.save(path)

base = [("S1", "Hood", "$1"), SAME, SAME, SAME, ("S2", "Trunk", "$1"), SAME, SAME, ("S3", "IG", "$2")]
build("dup_old.docx", base)
build("dup_a.docx", base[:2] + base[3:])                                         # one identical row deleted
b = list(base); b[6] = ("-", "Reserved", "$5"); build("dup_b.docx", b)             # one identical row changed
build("dup_c.docx", base[:6] + [SAME] + base[6:])                                 # one identical row added
print("expected: dup_old vs dup_a / dup_b / dup_c -> 1 change each (Deleted / Modified / New), heading 信号一覧")
