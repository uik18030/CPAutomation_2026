"""Step 2b test data: build an old/new .docx pair that differs ONLY in formatting.

Run (from the repo folder, with .venv active; needs `pip install python-docx` once):
    python tests/make_format_pair.py
Writes tests/data/Format_Old.docx and tests/data/Format_New.docx.
Then check:
    python tests/check_answer_key.py --old tests/data/Format_Old.docx --new tests/data/Format_New.docx
        --key tests/expected/answer_key_format.json --out output/check_format
Expected: 7 passed, 0 failed (the highlighted word must NOT be reported).
"""

from __future__ import annotations

import argparse
from pathlib import Path 

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]

PARAS = [
    "本書は新たに追加された機能の仕様を説明する。",
    "警告音は条件成立後に吹鳴すること。",
    "表示は運転者が容易に確認できること。",
    "下線のテストとして重要な語を示す。",
    "文字サイズの変更を確認するための文章。",
    "ハイライトは変更として扱わないこと。",
    "信号名はT_dis_RSPとすること。",
]
# paragraph index -> (word, formatting applied in the new file); keep in line with
# tests/expected/answer_key_format.json
EDITS = {
    0: ("新たに", "red"),
    1: ("警告音", "bold"),
    2: ("容易", "italic"),
    3: ("重要", "underline"),
    4: ("サイズ", "size"),
    5: ("ハイライト", "highlight"),   # must be ignored
    6: ("T_dis_RSP", "blue"),
}

def _style(run, how: str) -> None:
    if how == "red":
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    elif how == "blue":
        run.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)
    elif how == "bold":
        run.bold = True
    elif how == "italic":
        run.italic = True
    elif how == "underline":
        run.underline = True
    elif how == "size":
        run.font.size = Pt(14)
    elif how == "highlight":
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW

def build(path: Path, edited: bool, font: str) -> None:
    d = Document()
    st = d.styles["Normal"]
    st.font.name = font
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), font)   # Japanese text uses the same font

    d.add_paragraph("1. 概要")
    for i, text in enumerate(PARAS):
        p = d.add_paragraph()
        if edited and i in EDITS:
            word, how = EDITS[i]
            before, after = text.split(word, 1)
            p.add_run(before)
            _style(p.add_run(word), how)
            p.add_run(after)
        else:
            p.add_run(text)

    d.add_paragraph("2. 条件信号")
    t = d.add_table(rows=2, cols=2)
    t.style = "Table Grid"
    t.cell(0, 0).text = "信号"
    t.cell(0, 1).text = "値"
    t.cell(1, 0).text = "入力"
    run = t.cell(1, 1).paragraphs[0].add_run("[3D5]")
    if edited:
        run.bold = True                                  # bold inside a table cell
    path.parent.mkdir(parents=True, exist_ok=True)
    d.save(str(path))

def main() -> None:
    ap = argparse.ArgumentParser(description="Build the formatting test pair (Step 2b).")
    ap.add_argument("--font", default="MS Mincho", help='font for all text (default: "MS Mincho")')
    ap.add_argument("--out", default=str(ROOT / "tests/data"), help="output folder")
    args = ap.parse_args()
    out = Path(args.out)
    build(out / "Format_Old.docx", False, args.font)
    build(out / "Format_New.docx", True, args.font)
    print(f"Wrote {out / 'Format_Old.docx'} and {out / 'Format_New.docx'}")

if __name__ == "__main__":
    main()