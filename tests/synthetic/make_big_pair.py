"""Large spec-like test pair with ~50 known changes. Writes big_old.docx, big_new.docx, expected.json."""
import os
from pathlib import Path

# output folder: <repo>/output/synthetic/<script name>/   (git-ignored)
OUT = Path(__file__).resolve().parents[2] / "output" / "synthetic" / Path(__file__).stem
OUT.mkdir(parents=True, exist_ok=True)
os.chdir(OUT)
FONT = os.environ.get("CPA_TEST_FONT", "MS Gothic" if os.name == "nt" else "IPAGothic")
import json, random
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Mm
from PIL import Image, ImageDraw

R = random.Random(7)
SUBJ = ["メーター照明", "警告灯", "燃料計", "水温計", "シフト表示", "外気温表示", "時計表示", "平均燃費",
        "航続可能距離", "タイヤ空気圧", "ドア開警告", "シートベルト警告", "ブザー", "マルチインフォメーション",
        "オドメーター", "トリップメーター", "車速表示", "回転計", "エコ表示", "ヘッドランプ表示"]
COND = ["IG-ON中", "IG-OFF判定成立後", "車速5km/h以上", "エンジン始動後", "イルミ信号ON時", "診断要求受信時",
        "通信途絶検出時", "バッテリ電圧低下時", "シフトR位置", "パーキングブレーキ作動中"]
ACT = ["表示を更新すること。", "点灯すること。", "消灯すること。", "点滅表示すること。", "吹鳴すること。",
       "前回値を保持すること。", "初期表示とすること。", "表示を非表示とすること。", "調光すること。", "応答すること。"]
SIG = ["Hood Switch Status", "Trunk Switch Status", "Door Status D", "Door Status P", "IG Status", "Vehicle Speed",
       "Engine Speed", "Fuel Level", "Coolant Temp", "Illumination", "Shift Position", "Ambient Temp"]

def sentence(k):
    return f"{SUBJ[k % len(SUBJ)]}は{COND[(k * 3) % len(COND)]}に条件{k}が成立した場合、{ACT[(k * 7) % len(ACT)]}"

# ---------- base document model
sections = []
for s in range(1, 71):
    paras = [sentence(s * 100 + i) for i in range(R.randint(8, 13))]
    rows = [(f"{s}.{r}", f"[{R.randint(1,9)}{'ABCDEF'[R.randint(0,5)]}{R.randint(0,9)}] {SIG[(s + r) % len(SIG)]}",
             f"${R.randint(0,3)}") for r in range(1, R.randint(5, 12))]
    sections.append({"title": f"{SUBJ[s % len(SUBJ)]}制御{s}", "paras": paras, "rows": rows,
                     "img": s % 6 == 0, "bold": {}, "color": {}, "under": {}, "strike": set()})
REV = [("1.0", "2014.11.11", "-", "Ver1.0 Issued"), ("1.1", "2015.11.11", "-", "Ver1.1 Issued")]
for v in range(2, 9):
    REV.append((f"{v}.0", f"20{14 + v}.{v}.1{v}", f"{v}.1", f"誤記修正{v}"))
    for k in range(2):
        REV.append(("", "", f"{v}.{k + 2}", f"{SUBJ[v + k]}の説明を修正"))
cover = {"Version": "8.0", "Status": "Released", "Spec No.": "PM-1", "Issue Date": "2024.3.18"}

import copy
old, new = copy.deepcopy(sections), copy.deepcopy(sections)
newcover = dict(cover, **{"Version": "9.0", "Status": "DRAFT", "Issue Date": "2026.8.18"})
REV_NEW = REV + [("9.0", "2026.8.18", "3.2", "英語表記を日本語表記へ変更"), ("", "", "12.4", "新規Typeの条件を追加"),
                 ("", "", "25", "診断応答の説明を追加")]
E = []   # expected: (kind, token) token must appear in the change's text/details
E += [("Modified", "9"), ("Modified", "DRAFT"), ("Modified", "2026.8.18")]
E += [("New", "英語表記を日本語表記へ変更"), ("New", "新規Typeの条件を追加"), ("New", "診断応答の説明を追加")]

# text edits (10)
for s in (2, 4, 7, 9, 11, 14, 17, 21, 25, 28):
    p = new[s - 1]["paras"]; i = 1
    p[i] = p[i].replace("成立した場合", "成立し、かつ1秒経過した場合"); E.append(("Modified", f"条件{s * 100 + i}"))
# paragraph deleted (4) / added (4)
for s in (3, 13, 19, 27):
    t = new[s - 1]["paras"].pop(2); E.append(("Deleted", f"条件{s * 100 + 2}"))
for s in (5, 15, 22, 29):
    t = f"{SUBJ[s % len(SUBJ)]}の新規要件{s}として、表示遅延は100ms以内とすること。"; new[s - 1]["paras"].insert(3, t); E.append(("New", f"新規要件{s}"))
# paragraphs moved (3) - NOT reported
for a, b in ((6, 16), (10, 26), (18, 8)):
    t = new[a - 1]["paras"].pop(-1); new[b - 1]["paras"].append(t)
# table cell modified (6) / row inserted (4) / row deleted (4)
for s in (1, 6, 12, 16, 20, 25):
    r = list(new[s - 1]["rows"][1]); r[2] = "$9"; new[s - 1]["rows"][1] = tuple(r); E.append(("Modified", "$9"))
for s in (8, 14, 23, 30):
    new[s - 1]["rows"].insert(2, (f"{s}.N", "[7Z7] New Signal Added", "$5")); E.append(("New", "New Signal Added"))
for s in (9, 18, 26, 29):
    t = new[s - 1]["rows"].pop(3); E.append(("Deleted", t[0]))
# formatting (7)
def subj(sec, i): return SUBJ[(sec * 100 + i) % len(SUBJ)]
new[3]["bold"][0] = subj(4, 0); E.append(("Modified", "bold added 「" + subj(4, 0)))
new[10]["bold"][2] = subj(11, 2); E.append(("Modified", "bold added 「" + subj(11, 2)))
new[6]["color"][3] = subj(7, 3); E.append(("Modified", "font colour"))
new[19]["color"][0] = subj(20, 0); E.append(("Modified", "font colour"))
new[22]["under"][4] = subj(23, 4); E.append(("Modified", "underline added 「" + subj(23, 4)))
new[12]["strike"].add(0); E.append(("Modified", "strikethrough added"))
old[15]["strike"].add(1); E.append(("Modified", "strikethrough removed"))
# image rotated (1): section 12
ROT = {12}; E.append(("Modified", "image"))
# new section (after 20) / deleted section (24)
new.insert(20, {"title": "診断制御", "paras": ["診断要求を受信した場合は、診断仕様書に従い応答すること。",
                "診断セッション中はメーター表示を継続すること。"], "rows": [], "img": False,
                "bold": {}, "color": {}, "under": {}, "strike": set()}); E.append(("New", "診断制御"))
dead = new.pop(24); E.append(("Deleted", dead["title"]))   # section 24 (index shifted by the insert)

def pic(s, rotate):
    im = Image.new("RGB", (520, 200), "white"); d = ImageDraw.Draw(im)
    d.rectangle([20, 50, 170, 150], outline="black", width=3); d.text((40, 95), f"ECU-{s}", fill="black")
    d.rectangle([350, 30, 500, 170], outline="black", width=3); d.text((370, 95), f"Unit {s}", fill="black")
    d.line([170, 100, 350, 100], fill="black", width=3); d.polygon([(350, 100), (335, 90), (335, 110)], fill="black")
    d.text((20, 10), f"Fig {s}", fill="black")
    if rotate: im = im.rotate(180)
    path = f"fig{s}{'_r' if rotate else ''}.png"; im.save(path); return path

def table(doc, header, rows, merge=False):
    t = doc.add_table(rows=1 + len(rows), cols=len(header)); t.style = "Table Grid"
    for c, h in enumerate(header): t.cell(0, c).text = h
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row):
            if v or not merge or c > 1: t.cell(r, c).text = v
    if merge:
        starts = [r for r, row in enumerate(rows, start=1) if row[0]] + [len(rows) + 1]
        for a, b in zip(starts, starts[1:]):
            for c in (0, 1):
                if b - 1 > a:
                    m = t.cell(a, c).merge(t.cell(b - 1, c)); m.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for p in m.paragraphs[1:]: p._element.getparent().remove(p._element)

def build(path, secs, cov, rev, is_new):
    doc = Document(); st = doc.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    doc.add_paragraph("IPC POWER METER Requirements Specification")
    table(doc, ("項目", "内容"), list(cov.items()))
    if not is_new:
        doc.add_paragraph("改訂履歴"); table(doc, ("Ver.", "Date", "Section", "Description"), rev, merge=True)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    if is_new:
        doc.add_paragraph("改訂履歴"); table(doc, ("Ver.", "Date", "Section", "Description"), rev, merge=True)
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    for n, s in enumerate(secs, start=1):
        doc.add_paragraph(f"{n}. {s['title']}")
        for i, t in enumerate(s["paras"]):
            p = doc.add_paragraph()
            word = s["bold"].get(i) or s["color"].get(i) or s["under"].get(i)
            if word and t.startswith(word):
                r = p.add_run(word)
                if i in s["bold"]: r.bold = True
                if i in s["color"]: r.font.color.rgb = RGBColor(0, 0, 255)
                if i in s["under"]: r.underline = True
                p.add_run(t[len(word):])
            else:
                r = p.add_run(t); r.font.strike = i in s["strike"]
        if s["rows"]:
            doc.add_paragraph(f"表{n} 条件信号一覧"); table(doc, ("No.", "条件", "値"), s["rows"])
        if s["img"]:
            num = int(s["title"].rstrip("0123456789") and s["title"][len(s["title"].rstrip("0123456789")):] or 0)
            doc.add_picture(pic(num, is_new and num in ROT), width=Mm(110))
            doc.add_paragraph(f"図{n} 構成図")
    doc.save(path)

build("big_old.docx", old, cover, REV, False)
build("big_new.docx", new, newcover, REV_NEW, True)
json.dump(E, open("expected.json", "w"), ensure_ascii=False)
print(len(E), "expected changes")
