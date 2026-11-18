# CPA Tool – Handover (continue in a new conversation)

> Upload this file + docs/requirements.md + docs/progress.md when starting a new conversation.
> Working rules: **no code until the user says so; ask questions one at a time;
> send long code in short pieces (one file per message) to avoid cut-offs.**

## 1. What the project is
- **CPA = Change Point Analysis.** Compare an ORIGINAL (old) and LATEST (new) spec, find every
  visible difference, and fill the CPA Excel template: heading, change type, old screenshot,
  new screenshot with RED boxes.
- Team does this by hand in Word today. This is a NEW tool (repo: CPAutomation_2026).
  Old CPAuto (github.com/uik07454/CPA-VIO) is not used, only ideas were borrowed.
- Documents are **Japanese**, mostly digital PDF; old file is sometimes .docx/.doc.
  Translation is out of scope (team does it later).

## 2. Agreed requirements (summary – full list in docs/requirements.md)
| Topic | Decision |
|---|---|
| Inputs | PDF / .docx / .doc, mixed pairs OK. .docx/.doc → PDF via Word with **fields locked** (no "Error! Reference source not found"). Never PDF → Word |
| Method | Compare rendered PDFs, not Word Compare. Whole document = one stream, so content pushed to another page is NOT a change. Paragraphs first, then characters (Japanese = character level) |
| Counts as change | Any visible difference: text, values, table cells/rows, images (incl. rotate/flip/resize), strikethrough added/removed, font colour/bold/size, table formatting, spacing |
| Ignored | Cyan/any highlight, page numbers, invisible changes |
| Reported | TOC (one row per TOC page), page references, header/footer (once, not per page) |
| Column C | New / Deleted / Modified |
| Spacing | Clear changes → CPA_JPN; small ones → separate sheet 「要確認_Minor」 |
| Rows | One row per change (paragraph / table row / image); a whole deleted/new section = one row |
| Column B | Nearest **numbered** heading above the change, without its number (1)条件信号 → 条件信号). Before first heading: 表紙. Header/footer: ヘッダー / フッター |
| D, E | Empty (keep template dropdowns) |
| F Old Spec | Screenshot. Red box only around deleted/removed content |
| G Latest Spec | Screenshot. Red box around new/changed content; red marker where content was removed |
| Screenshots | Full page width, cropped around the change (+40pt); table cells cropped with the whole table row |
| Template | templates/CPA_Template_noAutoCount.xlsx, sheet CPA_JPN, header row 1, data from row 2. Written **via Excel (pywin32)**, not openpyxl (openpyxl drops the dropdowns) |
| App | PySide6 (LGPL) window + command line, PyInstaller .exe later |
| Order of files | Old file first, new file second (command line); separate fields in the app |

## 3. Environment (user's PC)
- Windows, repo: `C:\Users\uik18030\uik18030_github\CPAutomation_2026`
- `python` is NOT on PATH → always activate the venv first:
  `.\.venv\Scripts\Activate.ps1` (prompt shows `(.venv)`)
- Python 3.12 (Microsoft Store) in `.venv`; 3.15 also installed (too new, not used)
- Installed: pdfplumber 0.11.10, pypdfium2 5.15.2, Pillow 13.0.0, ImageHash 4.3.2,
  openpyxl 3.1.5, docx2pdf 0.1.8, pywin32 312, PySide6_Essentials 6.12.0 (+ numpy, scipy…).
  Versions pinned with `pip freeze`
- Word + Excel installed (used through COM)
- Company network: pip works directly (slow → `--timeout 120 --retries 10`)
- PowerShell: file names with `[ ]` need `-LiteralPath`; use short copies like
  `data\doc2_old.docx` / `data\doc2_new.pdf`
- `.gitignore`: .venv/, *.pdf, *.docx, *.doc, data/, output/, __pycache__/, build/, dist/

## 4. Code structure (cpa/)
| File | Role |
|---|---|
| convert.py | .docx/.doc → PDF via Word (fields locked, read-only). `ensure_password` / `get_password`: asks for a PDF password in the terminal (hidden, never saved) |
| extract.py | pdfplumber: text lines + chars (font, size, colour), thin horizontal shapes, tables (`_logical_columns` merges double-border/shading strips), images (position + data hash) |
| clean.py | Blocks: heading / para / cell / hf. Strikethrough (thin line through the middle 35–65% of the char), cells placed by table grid, paragraphs joined, headings (`heading_title`), page numbers removed, text normalised (full/half width, punctuation). `CleanDoc.table_rows` = every table row's area (incl. empty rows); `CleanDoc.images` |
| compare.py | Block alignment (exact first, then similar ≥0.5), character diff (incl. strikethrough), red boxes (English words boxed whole), New/Deleted rows find the matching table row in the other file (`_partner_row`, `_row_area`), merges per table row / whole section. Numbers changes in document order (text + images) |
| images.py | Image changes. .docx: rotation/flip read from Word (`a:xfrm rot/flipH/flipV`, exact). Otherwise perceptual hash incl. rotated/flipped versions. Rotation preferred over flip |
| shots.py | Renders pages (150 dpi), crops, draws red boxes/markers → output/<run>/shots/NN_old.png / NN_new.png |
| excel.py | Copies template, fills CPA_JPN via Excel COM, pictures move/size with cells |
| tests/check_answer_key.py | Runs the test pair, compares with tests/expected/answer_key.json |

Commands:
```
python -m cpa.excel OLD NEW --out output\<name>        # full run → CPA_<new>.xlsx, changes.txt, shots\
python -m cpa.compare OLD NEW                           # changes.txt only
python -m cpa.extract OLD NEW                           # dump_*.txt (positions, tables, images)
python tests/check_answer_key.py                        # must show 6 passed, 0 failed
```

## 5. Test data
- **Answer key pair** (tests/data, not in git): File_Old_update.docx / File_New_update.docx, 6 changes:
  1 [3DS]→[3D5] (条件信号, table row 1) · 2 *1 removed (row 2) · 3 Trank image rotated 180° ·
  4 「新たに」 added · 5 strikethrough removed · 6 T_dis_RSR→T_dis_RSP.
  Current result: **6 passed, 0 failed**.
- **Real pair 1** (data, not in git): YGC CS 適用書, old .docx vs supplier PDF → 4 changes, all correct;
  crop for cells empty in one file fixed.
- **Real pair 2** (data, not in git): doc2, old .docx vs password-protected new PDF → 189 changes
  (76 New, 61 Deleted, 52 Modified). Problems listed in section 7.

## 6. Progress
- [x] Step 0 setup
- [x] Step 1a–1f (convert, extract, clean, compare, screenshots, Excel, answer-key check)
- [x] Password-protected PDF support
- [x] Step 2a images (verify it is committed: "Step 2a: image changes")
- [ ] **Fixes from the doc2 test (code given but probably NOT applied yet – check):**
  - excel.py: `write_excel` two-pass (pass 1 text + row heights incl. wrap text, pass 2 pictures,
    warning if a picture is not in its own row). Cause: rows grew after pictures were placed →
    pictures drifted up from ~row 111 (rows 111/117/175/182 had 2–5 pictures, rows 114/177/184–190 none)
  - clean.py `heading_title`: reject titles starting with `[ ［ ( （ < ＜ { ｛` and titles with
    fewer than 2 letters/kana/kanji (fixes headings like "② ③ ④", "⑤", "[km/h]")
- [ ] Investigate doc2 changes 45–64 (Deleted block followed by New block under 「単発(サイン音/音声案内)」):
      likely table cells not pairing between .docx-PDF and supplier PDF
- [ ] 2b formatting (colour, bold/italic from font name, underline, size)
- [ ] 2c table formatting · 2d spacing + 「要確認_Minor」 · 2e pixel safety net
- [ ] Step 3 Excel finish · Step 4 app

## 7. Known issues / limits
- PDF-only comparison: a rotated **symmetric** picture may be mislabelled (no Word rotation to read)
- Moved text shows as Deleted + New
- Tables without borders are read as paragraphs
- Diagrams drawn from shapes are not compared yet (2e)
- Password-protected .docx not handled (Word would block in background)
- Big real documents: many New/Deleted pairs may really be Modified (see section 6)

## 8. Agreed ideas for later
- **Review screen** (PySide6, becomes the main app, ~1 week): side-by-side old/new per change,
  Keep/Discard, recrop from the PDF page, add/delete red boxes, edit heading (column B) and type,
  add missed change by hand, export only kept rows. Review decisions saved to a review file.
- **Header info input**: fields for old/new ticket, variant, received date → written into F1/G1
  (typed values in red, rest black, same layout as the template).
- **VIO chat (Step 4b)**: slide-up "Chat with VIO" panel (translate / summarise / explain term,
  optional "include selected row"). Pending IT: user vs shared token, approval for CONFIDENTIAL
  customer documents, API details, limits. Off by default; tool works without it; answers never
  written into the sheet automatically; token never in code/GitHub.
- Optional later: warning if old/new look swapped (cover version/date).

## 9. Next action when resuming
1. Confirm the two fixes in section 6 are applied (excel.py two-pass, heading_title). If not, apply them.
2. Rerun doc2: `python -m cpa.excel data\doc2_old.docx data\doc2_new.pdf --out output\doc2`
   and `python tests/check_answer_key.py`.
3. Check rows 111–190 line up; report any "WARNING: picture(s) not in their own row".
4. Look at changes 45–64 in output\doc2\changes.txt (structure lines only, no confidential text).
5. Then decide: review screen next, or continue 2b.
