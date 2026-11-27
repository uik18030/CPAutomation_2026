"""Realistic dummy spec pair (old Ver5.0 / new Ver6.0) with a known list of changes."""
import os
from pathlib import Path

# output folder: <repo>/output/synthetic/<script name>/   (git-ignored)
OUT = Path(__file__).resolve().parents[2] / "output" / "synthetic" / Path(__file__).stem
OUT.mkdir(parents=True, exist_ok=True)
os.chdir(OUT)
FONT = os.environ.get("CPA_TEST_FONT", "MS Gothic" if os.name == "nt" else "IPAGothic")
import json
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_BREAK, WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Mm, RGBColor
from PIL import Image, ImageDraw

CYAN = WD_COLOR_INDEX.TURQUOISE

def field(run, instr):
    for kind in ("begin", "instr", "separate", "text", "end"):
        if kind in ("begin", "separate", "end"):
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), kind)
        elif kind == "instr":
            e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = f" {instr} "
        else:
            e = OxmlElement("w:t"); e.text = "1"
        run._r.append(e)

def row_flag(row, tag):
    trPr = row._tr.get_or_add_trPr(); e = OxmlElement(f"w:{tag}"); e.set(qn("w:val"), "true"); trPr.append(e)

def runs(p, parts):
    """parts: list of (text, dict(strike, hl, bold, color))"""
    for text, f in parts:
        r = p.add_run(text)
        r.font.strike = f.get("strike", False)
        if f.get("hl"): r.font.highlight_color = CYAN
        if f.get("bold"): r.bold = True
        if f.get("color"): r.font.color.rgb = RGBColor(*f["color"])
        if f.get("size"): r.font.size = Pt(f["size"])
    return p

def table(doc, header, rows, merge_cols=(), repeat=True, widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(header)); t.style = "Table Grid"
    for c, h in enumerate(header):
        t.cell(0, c).text = h
    if repeat: row_flag(t.rows[0], "tblHeader")
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row):
            cell = t.cell(r, c)
            if isinstance(v, list):
                cell.paragraphs[0].text = ""
                for k, parts in enumerate(v):
                    p = cell.paragraphs[0] if k == 0 else cell.add_paragraph()
                    runs(p, parts)
            elif v or c not in merge_cols:
                cell.text = v
    if merge_cols:
        starts = [r for r, row in enumerate(rows, 1) if row[merge_cols[0]]] + [len(rows) + 1]
        for a, b in zip(starts, starts[1:]):
            for c in merge_cols:
                if b - 1 > a:
                    m = t.cell(a, c).merge(t.cell(b - 1, c)); m.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for p in m.paragraphs[1:]: p._element.getparent().remove(p._element)
    return t

def pic(name, rotate=False, label="ECU"):
    im = Image.new("RGB", (560, 230), "white"); d = ImageDraw.Draw(im)
    d.rectangle([20, 60, 180, 170], outline="black", width=3); d.text((50, 108), label, fill="black")
    d.rectangle([380, 30, 540, 200], outline="black", width=3); d.text((420, 108), "Meter", fill="black")
    d.line([180, 115, 380, 115], fill="black", width=3); d.polygon([(380, 115), (364, 105), (364, 125)], fill="black")
    d.text((20, 15), f"{name}", fill="black"); d.ellipse([240, 150, 300, 210], outline="black", width=3)
    if rotate: im = im.rotate(180)
    path = f"{name}{'_r' if rotate else ''}.png"; im.save(path); return path

def build(path, new):
    V = "6.0" if new else "5.0"
    doc = Document(); st = doc.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    sec = doc.sections[0]; sec.header_distance = Mm(10); sec.top_margin = Mm(40); sec.bottom_margin = Mm(22)
    # ---- running page header: logo | title_Ver | CONFIDENTIAL, then Page x of y
    h = sec.header
    runs(h.paragraphs[0], [("ACME", {"bold": True, "size": 14}), ("                                                    CONFIDENTIAL", {"bold": True})])
    p = h.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    runs(p, [("IPC POWER METER Requirements Specification_Ver", {}), (V, {"hl": True}), ("    Page ", {})])
    field(p.add_run(), "PAGE"); p.add_run(" of "); field(p.add_run(), "NUMPAGES")
    f = sec.footer.paragraphs[0]; f.alignment = WD_ALIGN_PARAGRAPH.CENTER; f.text = "Internal"

    # ---- cover: approval table, big title, info box, company/department/e-mail
    t = table(doc, ("設計 Design", "確認 Check", "確認 Check", "承認 Approval"),
              [("Tanaka", "Suzuki", "Sato", "Ito"),
               ("2026.8.10" if new else "2023.10.10", "2026.8.12" if new else "2023.10.12",
                "2026.8.12" if new else "2023.10.12", "2026.8.15" if new else "2023.10.14")], repeat=False)
    for _ in range(3): doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    runs(p, [("IPC POWER METER", {"bold": True, "size": 24})])
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    runs(p, [("Requirements Specification", {"bold": True, "size": 20})])
    for _ in range(3): doc.add_paragraph()
    if new:
        ver = [[("5.0", {"strike": True, "hl": True})], [("6.0", {"hl": True})]]
        date = [[("2023.10.16", {"strike": True, "hl": True})], [("2026.8.18", {"hl": True})]]
    else:
        ver = [[("4.1", {"strike": True, "hl": True}), (" ", {}), ("5.0", {"hl": True})]]
        date = [[("2021.7.28", {"strike": True, "hl": True})], [("2023.10.16", {"hl": True})]]
    table(doc, ("項目", "内容"), [("Version", ver), ("Status", "DRAFT" if new else "Released"),
                                 ("Spec No.", "PM-1"), ("Issue Date", date)], repeat=False)
    doc.add_paragraph()
    for line in ("ACME MOTOR CORPORATION",
                 "Cockpit Development Div. Meter and HUD Development Dept." if new else "Cockpit Development Div. Meter Development Dept.",
                 "e-mail: meter-spec@example.com"):
        p = doc.add_paragraph(line); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---- 改訂履歴 (merged Ver/Date, crosses a page, repeated header)
    doc.add_paragraph("改訂履歴")
    rev = [("1.0", "2014.11.11", "-", "Ver1.0 Issued"), ("1.1", "2015.11.11", "-", "Ver1.1 Issued"),
           ("1.2", "2016.6.10", "-", "Ver1.2 Issued")]
    for v, d, items in (("2.0", "2020.01.30", ["表紙:仕様書名称変更", "改訂履歴:表紙の次ページに移動", "適用:削除", "序章,1章,2章:新規追加(ロジック変更無し)"]),
                        ("3.0", "2020.9.1", ["4.1:誤記修正、記載文章の見直し", "4.2:Type-2追加(ロジック変更有り)", "5:警告音の条件を追加"]),
                        ("4.0", "2021.3.3", ["2.2:信号名を変更", "3.1:表示条件を修正", "3.4:部署名変更"]),
                        ("4.1", "2021.7.28", ["4.4.1:誤記修正", "-:部署名変更", "1.2:略語削除"]),
                        ("5.0", "2023.10.16", ["1.3.2:セグメントタイプ(制御仕様 Type-3)用の定数を追加", "3.2:機能概要図のECU名称を変更",
                                               "4:Type追加に伴い、制御仕様の説明を修正", "4.3:新規Type追加"])):
        for k, it in enumerate(items):
            sec_, desc = it.split(":", 1)
            rev.append((v if k == 0 else "", d if k == 0 else "", sec_, desc))
    if new:
        for k, (sec_, desc) in enumerate((("4.4、4.4.1、4.5", "英語表記を日本語表記へ変更"), ("2.3", "診断機能を追加"),
                                          ("4.6", "新規Typeの指示条件を追加"))):
            rev.append(("6.0" if k == 0 else "", "2026.8.18" if k == 0 else "",
                        [[(sec_, {"hl": True})]], [[(desc, {"hl": True})]]))
    table(doc, ("Ver.", "Date", "Section", "Description"), rev, merge_cols=(0, 1))
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---- 目次
    doc.add_paragraph("目次")
    toc = [("1. 序章", 4), ("1.1 適用範囲", 4), ("1.2 用語", 4), ("2. 機能概要", 5), ("2.1 システム構成", 5), ("2.2 条件信号", 5)]
    toc += [("2.3 診断機能", 6)] if new else []
    toc += [("3. 表示制御", 7 if new else 6), ("3.1 エンディングモード制御", 7 if new else 6), ("3.2 ラベル定義", 8 if new else 7), ("4. 警告音", 10 if new else 9)]
    for title, pg in toc:
        doc.add_paragraph(f"{title} ........................................ {pg}")
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---- 1. 序章
    doc.add_paragraph("1. 序章")
    doc.add_paragraph("1.1 適用範囲")
    doc.add_paragraph("本書はIPCのPOWER METERの" + ("制御および表示に関する" if new else "制御に関する") +
                      "仕様書である。本書の補足として表1記載の仕様書を参照のこと。")
    doc.add_paragraph("本書内で以下の仕様書を指示する場合は、\"Spec No.\"で示す。")
    doc.add_paragraph("表1 参照仕様書一覧")
    refs = [("PM-2", "CAN通信仕様書", "Ver3.0"), ("PM-5", "警告表示仕様書", "Ver2.1"), ("PM-7", "照明制御仕様書", "Ver1.4")]
    if new:
        refs = [("PM-2", "CAN通信仕様書", "Ver3.1"), ("PM-5", "警告表示仕様書", "Ver2.1"), ("PM-6", "診断仕様書", "Ver1.0"), ("PM-7", "照明制御仕様書", "Ver1.4")]
    table(doc, ("Spec No.", "名称", "Version"), refs)
    doc.add_paragraph("1.2 用語")
    for term in ("IG: イグニッション", "IPC: Instrument Panel Cluster", "ECU: Electronic Control Unit"):
        p = doc.add_paragraph("• " + term); p.paragraph_format.left_indent = Mm(5)

    # ---- 2. 機能概要
    doc.add_paragraph("2. 機能概要")
    doc.add_paragraph("2.1 システム構成")
    doc.add_paragraph("システム構成を図1に示す。")
    doc.add_picture(pic("fig1"), width=Mm(110)); doc.add_paragraph("図1 システム構成図")
    doc.add_paragraph("2.2 条件信号")
    doc.add_paragraph("各表示の条件信号を表2に示す。")
    sig = [("Hoodが開いている", [[("以下条件がすべて成立した場合", {})], [("• [3D5] Hood Switch Status == $1", {})], [("• [3DS]通信正常状態確定" if not new else "• [3D5]通信正常状態確定", {})]], "$1"),
           ("Trunk/Rear Doorが開いている" + ("" if new else "*1"), "[3D5] Trunk Switch Status == $1", "$1"),
           ("運転席ドアが開いている", "[3D6] Door Status D == $1", "$1"),
           ("助手席ドアが開いている", "[3D6] Door Status P == $1", "$1"),
           ("IG-ON", "[1A0] IG Status == $2", "$2")]
    if new:
        sig = [s for s in sig if not s[0].startswith("助手席")]
        sig.insert(3, ("後席ドアが開いている", "[3D7] Door Status R == $1", "$1"))
    table(doc, ("説明", "条件", "値"), sig)
    if new:
        doc.add_paragraph("2.3 診断機能")
        doc.add_paragraph("診断要求を受信した場合は、診断仕様書(PM-6)に従い応答すること。")
        table(doc, ("サービス", "内容"), [("0x22", "データ読出し"), ("0x2E", "データ書込み")])
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---- 3. 表示制御 (indentation levels)
    doc.add_paragraph("3. 表示制御")
    doc.add_paragraph("3.1 エンディングモード制御")
    doc.add_paragraph("以下条件に従い、エンディングモード制御をすること。")
    items = [(0, "• IG-OFF判定成立でエンディングモードへ遷移すること。"),
             (0, "• エンディングモードの表示完了前であっても、" + ("新たに" if new else "") + "オープニングモードの開始要求があった場合は、オープニングモードの開始を優先すること。"),
             (1, "➢ 開始要求はIG-ON判定成立時とする。"),
             (1, "➢ 表示中の警告は中断しないこと。"),
             (0, "• リヤシートリマインダ表示は以下条件に従い制御すること。"),
             (1, "① IG-ON中に[3D0] Multi-Display Information for Rear Seat Reminderの受信を確定した場合、リヤシートリマインダをT_dis_RSR間表示すること。"),
             (1, "② 受信確定の詳細なフローは、\"14.4.1リヤシートリマインダ\"を参照のこと。"),
             (0, "• メーター照明はイルミ信号に従い調光すること。")]
    if new:
        items[5] = (1, items[5][1].replace("T_dis_RSR", "T_dis_RSP"))
        items.insert(4, (1, "➢ 表示切替時はフェード処理を行うこと。"))
    for lvl, text in items:
        p = doc.add_paragraph(); p.paragraph_format.left_indent = Mm(5 + 8 * lvl)
        if new and text.startswith("• メーター照明"):
            runs(p, [("• ", {}), ("メーター照明", {"bold": True}), (text[len("• メーター照明"):], {})])
        elif not new and text.startswith("• IG-OFF"):
            runs(p, [(text, {"strike": True})])            # old: struck -> new: strike removed
        else:
            p.add_run(text)
    doc.add_paragraph("3.2 ラベル定義")
    doc.add_paragraph("各ラベルの定義を表3に示す。")
    labels = []
    for k in range(1, 31):
        desc = "セグメント表示のヒス閾値。表示値がこの範囲を超えた場合は、前回値を保持し、範囲内に戻った時点で表示を更新すること。" if k % 4 == 0 else "セグメント表示のヒス閾値"
        labels.append((f"Pw_Se-{k}d", "-", "1", "-860", "3300" if (new and k == 27) else "3235", desc))
    table(doc, ("LabelName", "Unit", "Resolution", "Min", "Max", "Description"), labels)
    doc.add_paragraph("3.3 外部表示")
    doc.add_paragraph("外部表示の構成を図2に示す。")
    doc.add_picture(pic("fig2", rotate=new, label="HUD"), width=Mm(100)); doc.add_paragraph("図2 外部表示構成図")

    # ---- 4. 警告音
    doc.add_paragraph("4. 警告音")
    p = doc.add_paragraph()
    runs(p, [("警告音は条件成立後0.5秒以内に吹鳴すること。", {"strike": new, "hl": new})])
    if not new:
        doc.add_paragraph("警告音の音量はセットアップ画面で切り替えること。")
    doc.add_paragraph("警告音の優先順位は表4に従うこと。")
    table(doc, ("優先度", "警告音", "周期"), [("1", "シートベルト", "1.0s" if not new else "0.8s"), ("2", "ドア開", "1.5s"), ("3", "ライト消し忘れ", "2.0s")])
    doc.save(path)

build("spec_old.docx", False); build("spec_new.docx", True)
EXPECTED = [
 ("Modified", "ヘッダー", "Ver5.0 -> Ver6.0 (page header, once)"),
 ("Modified", "表紙", "approval dates row"),
 ("Modified", "表紙", "Version 4.1 5.0 -> 5.0 6.0"),
 ("Modified", "表紙", "Status Released -> DRAFT"),
 ("Modified", "表紙", "Issue Date 2021.7.28 2023.10.16 -> 2023.10.16 2026.8.18"),
 ("Modified", "表紙", "Department name"),
 ("New", "表紙", "rev row 6.0 英語表記"), ("New", "表紙", "rev row 2.3 診断機能"), ("New", "表紙", "rev row 4.6"),
 ("?", "目次", "TOC: new 2.3 line + page numbers 6->7 etc (one row per TOC page per requirement)"),
 ("Modified", "適用範囲", "および表示"), ("Modified", "適用範囲", "PM-2 Ver3.0->3.1"), ("New", "適用範囲", "PM-6 row"),
 ("Modified", "条件信号", "[3DS]->[3D5]"), ("Modified", "条件信号", "*1 removed"), ("Deleted", "条件信号", "助手席 row"), ("New", "条件信号", "後席 row"),
 ("New", "診断機能", "new section 2.3 incl. table (one row)"),
 ("Modified", "エンディングモード制御", "strikethrough removed IG-OFF bullet"), ("Modified", "エンディングモード制御", "新たに"),
 ("New", "エンディングモード制御", "➢ フェード bullet"), ("Modified", "エンディングモード制御", "T_dis_RSR->RSP"), ("Modified", "エンディングモード制御", "bold メーター照明"),
 ("Modified", "ラベル定義", "Pw_Se-27d Max 3235->3300"),
 ("Modified", "外部表示", "fig2 rotated 180"),
 ("Modified", "警告音", "strikethrough added 0.5秒"), ("Deleted", "警告音", "音量 paragraph"), ("Modified", "警告音", "1.0s->0.8s"),
]
json.dump(EXPECTED, open("expected.json", "w"), ensure_ascii=False, indent=0)
print(len(EXPECTED), "expected items")
