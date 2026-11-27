"""Cover table where cells have struck-through + highlighted values (Version / Status / Issue Date)."""
import os
from pathlib import Path

# output folder: <repo>/output/synthetic/<script name>/   (git-ignored)
OUT = Path(__file__).resolve().parents[2] / "output" / "synthetic" / Path(__file__).stem
OUT.mkdir(parents=True, exist_ok=True)
os.chdir(OUT)
FONT = os.environ.get("CPA_TEST_FONT", "MS Gothic" if os.name == "nt" else "IPAGothic")
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.shared import Pt

def build(path, new):
    d = Document(); st = d.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    d.add_paragraph("IPC POWER METER Requirements Specification")
    t = d.add_table(rows=4, cols=2); t.style = "Table Grid"
    def cell(r, label, lines):
        t.cell(r, 0).text = label
        c = t.cell(r, 1); c.paragraphs[0].text = ""
        for k, (txt, strike, hl) in enumerate(lines):
            p = c.paragraphs[0] if k == 0 else c.add_paragraph()
            run = p.add_run(txt); run.font.strike = strike
            if hl: run.font.highlight_color = WD_COLOR_INDEX.TURQUOISE
    if new:
        cell(0, "Version", [("5.1", True, True), ("6.0", False, True)])
        cell(1, "Status", [("DRAFT", False, False)])
        cell(2, "Spec No.", [("PM-1", False, False)])
        cell(3, "Issue Date", [("2026.8.18", False, False)])
    else:
        cell(0, "Version", [("4.1 ", True, True)]); t.cell(0, 1).paragraphs[0].add_run("5.0").font.highlight_color = WD_COLOR_INDEX.TURQUOISE
        cell(1, "Status", [("Released", False, False)])
        cell(2, "Spec No.", [("PM-1", False, False)])
        cell(3, "Issue Date", [("2021.7.28", True, True), ("2023.10.2", False, True)])
    d.add_paragraph("1. 概要"); d.add_paragraph("本書は仕様を示す。")
    d.save(path)

build("cov_old.docx", False); build("cov_new.docx", True)
print("expected: 3 Modified rows (Version, Status, Issue Date), no Deleted rows")
