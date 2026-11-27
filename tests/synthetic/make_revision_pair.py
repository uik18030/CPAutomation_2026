import os
from pathlib import Path

# output folder: <repo>/output/synthetic/<script name>/   (git-ignored)
OUT = Path(__file__).resolve().parents[2] / "output" / "synthetic" / Path(__file__).stem
OUT.mkdir(parents=True, exist_ok=True)
os.chdir(OUT)
FONT = os.environ.get("CPA_TEST_FONT", "MS Gothic" if os.name == "nt" else "IPAGothic")
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_BREAK

HIST_OLD = [
    ("1.0", "2014.11.11", "-", "Ver1.0 Issued"),
    ("1.1", "2015.11.11", "-", "Ver1.1 Issued"),
    ("1.2", "2016.6.10", "-", "Ver1.2 Issued"),
    ("2.0", "2020.01.30", "表紙", "仕様書名称変更"),
    ("", "", "改訂履歴", "表紙の次ページに移動"),
    ("", "", "4.1", "誤記修正、記載文章の見直し"),
    ("4.1", "2021.7.28", "4.4.1", "誤記修正"),
    ("", "", "-", "部署名変更"),
    ("", "", "1.2", "略語削除"),
    ("5.0", "2023.10.16", "1.3.2", "セグメントタイプ用の定数を追加"),
    ("", "", "3.2", "機能概要図のECU名称を変更"),
    ("", "", "4", "Type追加に伴い、制御仕様の説明を修正"),
    ("", "", "4.3", "新規Type追加"),
]
HIST_NEW_EXTRA = [   # inserted after the "5.0" group
    ("6.0", "2026.8.18", "4.4、4.4.1、4.5", "英語表記を日本語表記へ変更"),
    ("", "", "4.5", "Type分けに関する説明を追加"),
    ("", "", "4.6", "新規Typeの指示条件を追加"),
]

from docx.enum.table import WD_ALIGN_VERTICAL
def table(d, rows, header, merge=False):
    t = d.add_table(rows=1 + len(rows), cols=len(header)); t.style = "Table Grid"
    for c, h in enumerate(header):
        t.cell(0, c).text = h
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row):
            if v or not merge or c > 1:
                t.cell(r, c).text = v
    if merge:   # Ver./Date cells span their group, text centred vertically
        starts = [r for r, row in enumerate(rows, start=1) if row[0]] + [len(rows) + 1]
        for a, b in zip(starts, starts[1:]):
            for c in (0, 1):
                if b - 1 > a:
                    m = t.cell(a, c).merge(t.cell(b - 1, c))
                    m.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for p in m.paragraphs[1:]:
                        p._element.getparent().remove(p._element)

def build(path, new):
    d = Document(); st = d.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    d.add_paragraph("IPC POWER METER Requirements Specification")
    table(d, [("Version", "6.0" if new else "5.0"), ("Status", "DRAFT" if new else "Released"),
              ("Spec No.", "PM-1"), ("Issue Date", "2026.8.18" if new else "2023.10.16")], ("項目", "内容"))
    hist = list(HIST_OLD)
    if new:
        hist.insert(9, ("", "", "1.2.1", "用語定義を追加"))          # middle of the 4.1 group
        hist = [r for r in hist if r[3] != "誤記修正、記載文章の見直し"]  # row deleted
        hist = [(a, b, "3.3" if c == "3.2" else c, d) for a, b, c, d in hist]  # short cell changed
        hist += HIST_NEW_EXTRA
    if not new:
        d.add_paragraph("改訂履歴")
        table(d, hist, ("Ver.", "Date", "Section", "Description"), merge=True)
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.add_paragraph("1. 序章")
    d.add_paragraph("本書はIPCのPOWER METERの制御に関する仕様書である。")
    d.add_paragraph("本書内で以下の仕様書を指示する場合は、Spec No.で示す。")
    if new:
        d.add_paragraph("新規Typeの表示条件は4.6章に従うこと。")
        d.add_paragraph("改訂履歴")
        table(d, hist, ("Ver.", "Date", "Section", "Description"), merge=True)
    d.add_paragraph("2. 条件信号")
    table(d, [("Hoodが開いている", "[3D5] Hood Switch Status == $1"),
              ("Trunkが開いている", "[3D5] Trunk Switch Status == $1")], ("説明", "条件"))
    d.save(path)

build("rev_old.docx", False); build("rev_new.docx", True)
