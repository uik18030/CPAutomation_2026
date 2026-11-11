# CPA Tool – Requirements & Decisions (master record)
Last updated: 2026-11-11. Status: requirements agreed.
**Rule: NO CODE until the user explicitly says to start.**

---

## 1. Background
- **CPA = Change Point Analysis.** Compare an ORIGINAL (old) and a LATEST document, find every difference, record the section heading, screenshot old and new, put a RED BOX on the change, and fill in the CPA Excel sheet.
- Today the team does CPA **by hand in Word**. The old CPAuto tool (github.com/uik07454/CPA-VIO) is **not in use**.
- **Goal:** build a NEW CPA tool.
  - Reuse from CPAuto: Excel template writing and image embedding (`output.py`), settings (`app_config.py`), GUI structure (`gui.py`), `ChangeRecord` (`models.py`).
  - The comparison engine is new.
- CPAuto's old engine reads Word Compare tracked changes from one compare `.docx`. It **doesn't need to be kept**.

## 2. Inputs
- Usually PDF. Sometimes the old document only exists as `.docx`.
- Mixed pairs allowed: PDF/PDF, docx/PDF, PDF/docx, docx/docx.
- Convert `.docx` → PDF through Microsoft Word (the team has Word). **Never convert PDF → Word.**
- PDFs are digital (text and tables can be selected), so **no OCR** is needed.
- Language: **Japanese**. Translation is done later by the team and is **out of scope**.

## 3. Comparison approach
- Compare **rendered PDF pages**, not Word XML. **Don't use Word Compare.**
- Treat each document as **one continuous stream**, so content that was only **pushed to another page is NOT a change**.
- Two levels: compare paragraphs first, then characters inside changed paragraphs.
- Japanese: compare **character by character**, but keep runs of letters and numbers (signal names, values) together.
- **Standardise text before comparing:**
  - full-width / half-width characters
  - Japanese / Western punctuation
  - full-width spaces
  - text split across line wraps and hyphenation
- Remove **page numbers** before comparing. Handle headers/footers and repeated table header rows separately.
- **Images:**
  - Match old and new images across the whole document by content: exact hash first, then perceptual hash.
  - Compare them **as they appear on the page**: rotation, flip, crop, size and position all count.
- **Diagrams** can be pictures or drawn shapes. Handle both: image fingerprint plus pixel comparison of the page area, and compare labels through the text diff.
- **Safety net:** a pixel comparison of matched blocks catches anything the other checks miss.

## 4. What counts as a change (rule: ANY visible difference)
**Counts as a change:**
- Text and value changes, additions, deletions.
- Tables: cell content, rows, cell shading, borders, merged/split cells.
- Images: added, deleted, content changed, **rotated** (e.g. 180°), flipped, resized, cropped, moved.
- **Strikethrough added or removed**, even when the text is the same.
- Font colour, bold, italic, underline, font size.
- Spacing (see section 6).
- Struck-through text is compared as normal text.

**Ignore:**
- Cyan highlight.
- Page numbers.
- Content that only moved to another page.
- Invisible changes, e.g. a field that became an automatic cross-reference but looks the same, or an empty paragraph at the end.

**Report:**
- Table of contents changes: **one row per TOC page**.
- Page references in the text (e.g. 12ページ参照 → 13ページ参照).
- Header/footer changes: **reported once**, not on every page.

## 5. Changes Description labels (column C)
| Label | Meaning |
|---|---|
| **New** | Content added |
| **Deleted** | Content removed |
| **Modified** | Text, value, image (including rotation), strikethrough, formatting or spacing changed. "Modify" and "update" also mean Modified |

## 6. Spacing sensitivity
**Goes in the `CPA_JPN` sheet (as Modified):**
- Blank line added or removed.
- Indentation changed by 1 character or more (about 10.5pt).
- Alignment changed.
- Space before/after a paragraph changed by half a line or more (about 6pt).

**Goes on a separate sheet 「要確認_Minor」** (same columns): smaller differences.

Thresholds are configurable.

## 7. Rows
- **One row per change** (paragraph, table row or image). The team merges or splits rows afterwards.
- Several edits in the same paragraph = **one row with several red boxes**.
- A change spanning a page break = one row, with the page crops stacked.
- Tables: one row per changed **table row**. *(Assumption; the user may switch to one row per cell.)*

## 8. Excel output: template `CPA_Template_noAutoCount.xlsx`
- Write into sheet **`CPA_JPN`**. Header is **row 1**; data starts at **row 2**.
- **Don't touch the header row.** The user fills in document numbers (YEX/YGC), ticket and received dates by hand.

| Column | Content |
|---|---|
| A Change Point (No.) | 1, 2, 3… |
| B Sheet/Header | Nearest numbered heading above the change, **without the number or brackets** (1)条件信号 → 条件信号; 14.4.1 リヤシートリマインダ → リヤシートリマインダ). Before the first heading: **表紙** |
| C Changes Description | New / Deleted / Modified |
| D DOORS ID | **Empty** (keep the dropdown) |
| E ID changes | **Empty** (keep the dropdown) |
| F Old Spec | Screenshot, **no boxes** |
| G Latest Spec | Screenshot with **RED boxes**. Deletions get a red marker where the content was removed |

- **Box colour:** everything red. The orange box in the example screenshot was just one reviewer's style.
- **Screenshot size:** full page width, cropped around the change (as in the user's example). Fit the column width (about 66 characters, about 460 px).
- **Keep** the autofilter A1:G200 (extend it if there are more rows) and the freeze pane at F2.
- **Don't touch** the sheet "List&guideline".
- **⚠ Important:** saving with openpyxl drops the template's dropdowns (the Data Validation extension). They must be **restored after saving**.

## 9. Heading detection
- Headings are **numbered**: 1) / 1. / (1) / ① / 14.4.1 / 第3章 / 第3節.
- **Not a heading:** lines ending in 「。」, or longer than about 40 characters.
- **Fallback:** the nearest short line without 「。」.

## 10. App
- Desktop window app built with **PySide6** (LGPL), not PyQt6 (GPL).
- Packaged as a **PyInstaller `.exe` folder**, plus a **command-line version**.
- **Window:**
  - Old Spec, Latest Spec, Template and Output fields, each with a Browse button
  - Collapsible Settings section
  - Run button, progress bar, Cancel button
- Settings live in a **config file**.
- **Licence-safe libraries:**
  - pdfplumber (MIT)
  - pypdfium2 (Apache/BSD)
  - Pillow
  - imagehash
  - scipy (optional)
  - **Avoid PyMuPDF (AGPL).**
- **Sharing with other teams:** not decided yet. Build for the team first, but keep it easy to share later.

## 11. Document size
- Must handle **both short and long** documents.
- Design so that time grows with the number of **changes**, not pages:
  - render only pages that contain a change
  - pair exact image matches first
  - progress bar and Cancel button
  - shrink screenshots before saving
  - warn if the output file gets very large

## 12. Test pair and answer key: File_Old_update.docx vs File_New_update.docx (6 changes)
| No. | Sheet/Header | C | Change | Red box in Latest Spec |
|---|---|---|---|---|
| 1 | 条件信号 | Modified | [3DS] → [3D5] (Hood row, 条件 column) | [3D5] |
| 2 | 条件信号 | Modified | 開いている*1 → 開いている (Trunk row, 説明) | Marker where *1 was |
| 3 | 条件信号 | Modified | **Trank image rotated 180°** (Trunk row, 意匠例; same image file, rot=10800000) | Whole image |
| 4 | エンディングモード制御 | Modified | 「新たに」 added, two paragraphs merged into one | 新たに |
| 5 | エンディングモード制御 | Modified | **Strikethrough removed** from 「リヤシートリマインダ表示は以下条件に従い制御すること。」 | Whole sentence |
| 6 | エンディングモード制御 | Modified | T_dis_RSR → T_dis_RSP | RSP |

- **Not changes:** the 14.4.1 line was turned into an automatic cross-reference but looks the same; an empty paragraph at the end.
- **Lesson:** identical image files can still look different on the page, so compare what's rendered.

## 13. Known bugs in the old CPAuto code (relevant when reusing it)
- **B1:** Delete Toggle column added twice (`output.py` and `gui.py`); in template mode the second pass may hit the wrong sheet.
- **B2:** HYPERLINK formula doesn't escape quotes, and Excel limits formula text to 255 characters.
- **B3:** `auto_delete_watcher` overwrites a file that Excel has open.
- **B4:** `is_caption` declared twice in `models.py`.
- **B5:** Mutable default arguments.
- **B6:** Header/footer extraction gets no image map.
- **Other issues:**
  - The word diff splits on spaces, which breaks Japanese.
  - `w:moveFrom`/`w:moveTo`, `w:hyperlink` and some other tags are skipped silently.
  - `build_image_map` reads only `document.xml.rels`, so header/footer images are missed.
  - `IMAGE_PX_MAX = 500` is too small for screenshots.

## 14. Build plan (about 3–3.5 weeks, 1 developer)
| Step | Content | Effort |
|---|---|---|
| 1. Proof of concept | Convert, extract, clean, compare, screenshot with red box, write `CPA_JPN`. Check against the 6-change answer key | 3–4 days |
| 2. Visual checks | Strikethrough, formatting, image rotation/size/crop, table formatting, spacing plus Minor sheet, pixel safety net | 4–5 days |
| 3. Excel output | Fill template, restore dropdowns, extend filter, add Minor sheet | 1–2 days |
| 4. App | PySide6 GUI, background thread, progress/Cancel, config file, command line | 2 days |
| 5. Real documents | 1 short and 1 long pair; tune thresholds and speed | 2–3 days |
| 6. Packaging | PyInstaller `.exe` | 1 day |

Testing has to run on the user's PC, because the assistant's environment has no Word and can't render PDFs.

## 15. Working rules
- **Don't write code** until the user explicitly says to start.
- Ask questions **one at a time** when information is needed.
