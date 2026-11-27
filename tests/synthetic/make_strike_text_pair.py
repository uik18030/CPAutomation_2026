"""Strikethrough removed + a text change + an underline present in both files (must not be reported)."""
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

for name, new in (("reg_old", False), ("reg_new", True)):
    d = Document(); st = d.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    d.add_paragraph("1. エンディングモード制御")
    p = d.add_paragraph(); r = p.add_run("リヤシートリマインダ表示は以下条件に従い制御すること。"); r.font.strike = not new
    p = d.add_paragraph(); p.add_run("信号T_dis_RSR" if not new else "信号T_dis_RSP"); p.add_run("を参照。")
    p = d.add_paragraph(); r = p.add_run("下線付きの共通テキスト"); r.underline = True
    d.save(name + ".docx")
print("expected: 2 Modified (strikethrough removed, T_dis_RSR -> T_dis_RSP)")
