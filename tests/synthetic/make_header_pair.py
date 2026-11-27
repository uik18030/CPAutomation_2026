"""Spec with a tall running page header (title + Page x of y) and long tables whose header row repeats."""
import os
from pathlib import Path

# output folder: <repo>/output/synthetic/<script name>/   (git-ignored)
OUT = Path(__file__).resolve().parents[2] / "output" / "synthetic" / Path(__file__).stem
OUT.mkdir(parents=True, exist_ok=True)
os.chdir(OUT)
FONT = os.environ.get("CPA_TEST_FONT", "MS Gothic" if os.name == "nt" else "IPAGothic")
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Mm

def field(run, instr):
    for kind, text in (("begin", None), (None, instr), ("separate", None), (None, "1"), ("end", None)):
        if kind:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), kind); run._r.append(e)
        elif text == instr:
            e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = f" {instr} "; run._r.append(e)
        else:
            e = OxmlElement("w:t"); e.text = text; run._r.append(e)

def no_split(t):
    for row in t.rows:
        trPr = row._tr.get_or_add_trPr(); e = OxmlElement("w:cantSplit"); e.set(qn("w:val"), "true"); trPr.append(e)

def repeat_header(row):
    trPr = row._tr.get_or_add_trPr(); h = OxmlElement("w:tblHeader"); h.set(qn("w:val"), "true"); trPr.append(h)

def build(path, ver, new):
    d = Document(); st = d.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    sec = d.sections[0]; sec.header_distance = Mm(14); sec.top_margin = Mm(42)
    h = sec.header
    h.paragraphs[0].text = "ACME"
    p = h.add_paragraph("CONFIDENTIAL"); p.alignment = 2
    p = h.add_paragraph(f"IPC POWER METER Requirements Specification_Ver{ver}    Page "); r = p.add_run(); field(r, "PAGE")
    p.add_run(" of "); r = p.add_run(); field(r, "NUMPAGES")
    d.add_paragraph("改訂履歴")
    rows = [(f"{v}.{k}", f"20{10+v}.{k+1}.1{k}", "-" if k == 0 else f"{v}.{k}", f"Ver{v}.{k} 変更内容{v}{k}") for v in range(1, 8) for k in range(4)]
    if new:
        rows.append(("6.0", "2026.8.18", "4.6", "新規Typeの指示条件を追加"))
    t = d.add_table(rows=len(rows) + 1, cols=4); t.style = "Table Grid"
    for c, x in enumerate(("Ver.", "Date", "Section", "Description")): t.cell(0, c).text = x
    repeat_header(t.rows[0])
    for r, row in enumerate(rows, 1):
        for c, x in enumerate(row): t.cell(r, c).text = x
    d.add_paragraph("1. ラベル定義")
    if new:
        d.add_paragraph("ラベルの一覧を以下に示す。各ラベルの閾値は表2に従うこと。")
    t = d.add_table(rows=61, cols=6); t.style = "Table Grid"
    for c, x in enumerate(("LabelName", "Unit", "Resolution", "Min", "Max", "Description")): t.cell(0, c).text = x
    repeat_header(t.rows[0])
    for r in range(1, 61):
        for c, x in enumerate((f"Pw_Se-{r}d", "-", "1", "-860", "3235" if not (new and r == 50) else "3300", "セグメント表示のヒス閾値")):
            t.cell(r, c).text = x
    no_split(t)
    d.add_paragraph("2. 表示制御")
    d.add_paragraph("表示はIG-ON中に更新すること。")
    d.save(path)
build("hdr_old.docx", "5.0", False); build("hdr_new.docx", "6.0", True)
