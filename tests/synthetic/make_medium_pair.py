"""Medium-size spec-like test pair with known changes (see EXPECTED at the bottom)."""
import os
from pathlib import Path

# output folder: <repo>/output/synthetic/<script name>/   (git-ignored)
OUT = Path(__file__).resolve().parents[2] / "output" / "synthetic" / Path(__file__).stem
OUT.mkdir(parents=True, exist_ok=True)
os.chdir(OUT)
FONT = os.environ.get("CPA_TEST_FONT", "MS Gothic" if os.name == "nt" else "IPAGothic")
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Mm
from PIL import Image, ImageDraw

# --- a simple block diagram picture (rotated 180° in the new file)
img = Image.new("RGB", (480, 220), "white"); d = ImageDraw.Draw(img)
d.rectangle([20, 60, 160, 160], outline="black", width=3); d.text((45, 100), "Meter ECU", fill="black")
d.rectangle([320, 60, 460, 160], outline="black", width=3); d.text((350, 100), "Gateway", fill="black")
d.line([160, 110, 320, 110], fill="black", width=3); d.polygon([(320, 110), (305, 100), (305, 120)], fill="black")
d.text((20, 20), "CAN-FD", fill="black")
img.save("diagram.png"); img.rotate(180).save("diagram_rot.png")

REV = [("1.0", "2014.11.11", "-", "Ver1.0 Issued"), ("1.1", "2015.11.11", "-", "Ver1.1 Issued"),
       ("1.2", "2016.6.10", "-", "Ver1.2 Issued"), ("2.0", "2020.01.30", "表紙", "仕様書名称変更"),
       ("", "", "改訂履歴", "表紙の次ページに移動"), ("", "", "適用", "削除"),
       ("4.1", "2021.7.28", "4.4.1", "誤記修正"), ("", "", "-", "部署名変更"), ("", "", "1.2", "略語削除"),
       ("5.0", "2023.10.16", "1.3.2", "セグメントタイプ用の定数を追加"), ("", "", "3.2", "機能概要図のECU名称を変更"),
       ("", "", "4", "Type追加に伴い、制御仕様の説明を修正"), ("", "", "4.3", "新規Type追加")]
REV_NEW = [("6.0", "2026.8.18", "4.4、4.4.1、4.5", "英語表記を日本語表記へ変更"),
           ("", "", "4.5", "Type分けに関する説明を追加"), ("", "", "4.6", "新規Typeの指示条件を追加")]
SIGNALS = [("Hoodが開いている", "[3D5] Hood Switch Status", "$1"),
           ("Trunkが開いている", "[3D5] Trunk Switch Status", "$1"),
           ("運転席ドアが開いている", "[3D6] Door Status D", "$1"),
           ("助手席ドアが開いている", "[3D6] Door Status P", "$1"),
           ("IG-ON", "[1A0] IG Status", "$2"),
           ("車速あり", "[1B4] Vehicle Speed", "≧5km/h")]
BULLETS = ["IG-OFF判定成立でエンディングモードへ遷移すること。",
           "エンディングモードの表示完了前であっても、オープニングモードの開始要求があった場合は、オープニングモードの開始を優先すること。",
           "メーター照明はイルミ信号に従い調光すること。",
           "燃料残量警告はFuel Level信号が10%以下で点灯すること。",
           "表示言語の切替はセットアップ画面から行うこと。",
           "シフト位置表示はShift Position信号に従い表示すること。",
           "外気温表示は-40℃から60℃の範囲とすること。",
           "時計表示は12時間表示と24時間表示を切り替え可能とすること。"]

def setup(doc):
    st = doc.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

def table(doc, header, rows, merge=False, widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(header)); t.style = "Table Grid"
    for c, h in enumerate(header):
        t.cell(0, c).text = h
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row):
            if v or not merge or c > 1:
                t.cell(r, c).text = v
    if merge:
        starts = [r for r, row in enumerate(rows, start=1) if row[0]] + [len(rows) + 1]
        for a, b in zip(starts, starts[1:]):
            for c in (0, 1):
                if b - 1 > a:
                    m = t.cell(a, c).merge(t.cell(b - 1, c))
                    m.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for p in m.paragraphs[1:]:
                        p._element.getparent().remove(p._element)
    return t

def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

def build(path, new):
    doc = Document(); setup(doc)
    doc.add_paragraph("IPC POWER METER Requirements Specification")
    table(doc, ("項目", "内容"), [("Version", "6.0" if new else "5.0"), ("Status", "DRAFT" if new else "Released"),
                                 ("Spec No.", "PM-1"), ("Issue Date", "2026.8.18" if new else "2023.10.16")])
    rev = REV + (REV_NEW if new else [])
    if not new:                                   # old: revision table on the cover
        doc.add_paragraph("改訂履歴")
        table(doc, ("Ver.", "Date", "Section", "Description"), rev, merge=True)
    page_break(doc)
    if new:                                       # new: revision table moved to page 2
        doc.add_paragraph("改訂履歴")
        table(doc, ("Ver.", "Date", "Section", "Description"), rev, merge=True)
        page_break(doc)

    doc.add_paragraph("1. 序章")
    doc.add_paragraph("本書はIPCのPOWER METERの" + ("制御および表示に関する" if new else "制御に関する")
                      + "仕様書である。本書の補足として表1記載の仕様書を参照のこと。")
    doc.add_paragraph("本書内で以下の仕様書を指示する場合は、Spec No.で示す。")
    doc.add_paragraph("表1 参照仕様書一覧")
    refs = [("PM-2", "CAN通信仕様書"), ("PM-5", "警告表示仕様書"), ("PM-7", "照明制御仕様書")]
    if new:
        refs = [("PM-3", "CAN通信仕様書"), ("PM-5", "警告表示仕様書"), ("PM-6", "診断仕様書"), ("PM-7", "照明制御仕様書")]
    table(doc, ("Spec No.", "名称"), refs)

    doc.add_paragraph("2. システム構成")
    doc.add_paragraph("システム構成を図1に示す。")
    doc.add_picture("diagram_rot.png" if new else "diagram.png", width=Mm(100))
    doc.add_paragraph("図1 システム構成図")

    doc.add_paragraph("3. 条件信号")
    doc.add_paragraph("各表示の条件信号を表2に示す。")
    sig = list(SIGNALS)
    if not new:
        sig[1] = ("Trunkが開いている", "[3DS] Trunk Switch Status", "$1")   # [3DS] -> [3D5]
    else:
        sig = [s for s in sig if not s[0].startswith("助手席")]              # row deleted
    table(doc, ("説明", "条件", "値"), sig)

    doc.add_paragraph("4. 表示制御")
    doc.add_paragraph("4.1 エンディングモード制御")
    for i, text in enumerate(BULLETS):
        if new and i == 4:
            continue                                                     # bullet deleted
        p = doc.add_paragraph("• ")
        if new and i == 2:
            r = p.add_run("メーター照明"); r.bold = True; p.add_run(text[len("メーター照明"):])
        elif new and i == 3:
            a, b = text.split("10%", 1); p.add_run(a); r = p.add_run("10%"); r.font.color.rgb = RGBColor(255, 0, 0); p.add_run(b)
        else:
            p.add_run(text)
        if new and i == 6:
            doc.add_paragraph("• 平均燃費表示はリセット操作で初期化すること。")  # bullet added
    doc.add_paragraph("4.2 警告灯制御")
    if not new:
        doc.add_paragraph("警告灯の点灯条件は別紙の警告表示仕様書に従うこと。")   # moved below in new
    doc.add_paragraph("警告灯はIG-ON後3秒間のバルブチェックを行うこと。")
    doc.add_paragraph("5. 警告音")
    p = doc.add_paragraph()
    r = p.add_run("警告音は条件成立後0.5秒以内に吹鳴すること。")
    if new:
        r.font.strike = True                                             # strikethrough added
    doc.add_paragraph("警告音の音量は3段階で切り替え可能とすること。")
    if new:
        doc.add_paragraph("警告灯の点灯条件は別紙の警告表示仕様書に従うこと。")   # moved here, same text
        doc.add_paragraph("6. 診断")
        doc.add_paragraph("診断要求を受信した場合は、診断仕様書に従い応答すること。")
        doc.add_paragraph("診断セッション中はメーター表示を継続すること。")
    doc.save(path)

build("med_old.docx", False); build("med_new.docx", True)

EXPECTED = """1 Version 5.0->6.0 | 2 Status Released->DRAFT | 3 Issue Date | 4-6 three new revision rows
7 序章 text edit | 8 ref table PM-2->PM-3 | 9 ref table new row PM-6 | 10 picture rotated 180
11 [3DS]->[3D5] | 12 signal row deleted | 13 bold メーター照明 | 14 red 10% | 15 bullet deleted
16 bullet added | 17 strikethrough added | 18 new section 6.診断 (one row)
NOT reported: revision table moving to page 2, paragraph moved from 4.2 to 5"""
open("expected.txt", "w").write(EXPECTED)
