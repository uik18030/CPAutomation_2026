# CPA Tool – Handover (read this first in a new conversation)

> Last updated: 2026-11-27. Repo: `uik18030/CPAutomation_2026`, branch `main`
> (last checked commit `124baa8`). Also read: **docs/App_Design.md** (agreed app design),
> docs/requirements.md, tests/synthetic/README.md, reference/README.md.

---

## 0. Working rules (agreed with the user – always follow)

1. **Claude never pushes or commits to the user's repo.** (Pushing from Claude's environment is
   also blocked: GitHub 403.) Every code change is **pasted in the chat as find/replace steps**
   ("Find … / Replace with … / Add directly ABOVE/BELOW") and the **user types it in** on the
   company laptop, then pushes. Documents and small test tools may be copy-pasted.
2. **One file per message**, small numbered steps, exact text, indentation hints.
3. **Build and test everything in Claude's environment first** (Linux, LibreOffice-made PDFs,
   PySide6 offscreen screenshots), then send it.
4. **After every push, check the user's code against the tested version**: git fetch, compare
   structure (Python AST, ignoring docstrings/comments/annotations), pyflakes, and run the user's
   code on the synthetic test pairs. Typical typing mistakes found so far: missing call brackets
   (`_unnumbered` instead of `_unnumbered(b)`), wrong attribute (`t.box` vs `t.bbox`), missing
   terms (`- first`), `Celles`/`Clles`, `parent=` vs `parents=`, `{ }` vs `( )` – pyflakes does
   **not** catch most of these, so always compare + run.
5. Ask questions **one batch at a time** (AskUserQuestion), give a recommendation.
6. **No confidential content leaves the user's laptop** except to approved VIO. Ask only for
   structure lines / screenshots of non-confidential parts; never repeat e-mail addresses or names
   seen in screenshots.
7. VIO token: **never** in chat, code or git (environment variable only).

---

## 1. What the project is

**CPA = Change Point Analysis.** Compare an ORIGINAL (old) and LATEST (new) Japanese spec, find every
visible difference, and fill the company CPA Excel template (sheet `CPA_JPN`): heading, change
type, old screenshot, new screenshot with RED boxes. The team does this by hand in Word today.
Internship project; the user is the only user for now (team later).

### 1.1 Agreed requirements (unchanged)

| Topic | Decision |
|---|---|
| Inputs | PDF / .docx / .doc, mixed pairs OK. .docx/.doc → PDF via Word with **fields locked**. Never PDF → Word |
| Method | Compare rendered PDFs. Whole document = one stream; content pushed to another page is NOT a change. Paragraphs/rows first, then characters |
| Counts as change | Any visible difference: text, values, table cells/rows, images (rotate/flip/resize), strikethrough added/removed, font colour/bold/italic/underline/size, (table formatting, spacing: not yet) |
| Ignored | Cyan/any highlight, page numbers, invisible changes, **moved unchanged content**, **renumbering only** (heading / 表N / 図N numbers) |
| Reported | **TOC: one row per TOC page** (label 目次), page references, header/footer **once** |
| Column C | New / Deleted / Modified |
| Rows | One row per change (paragraph / table row / image); a whole deleted/new section (heading + text + tables + pictures) = one row |
| Column B | Nearest numbered heading without its number; before first heading 表紙; header/footer ヘッダー / フッター |
| D, E | Empty (template dropdowns kept) |
| F Old Spec | Screenshot; red box only around removed content; **N/A for New rows** (no old picture) |
| G Latest Spec | Screenshot; red box around new/changed content; red marker where content was removed (only for pure removals) |
| Template | templates/CPA_Template_noAutoCount.xlsx, sheet CPA_JPN, header row 1, data from row 2, written **via Excel COM** (openpyxl drops dropdowns) |
| Spacing | Clear changes → CPA_JPN; small → sheet 「要確認_Minor」 (Step 2d, not done) |

---

## 2. Environment (user's PC)

- Windows, repo `C:\Users\uik18030\uik18030_github\CPAutomation_2026`, VS Code.
- Always activate the venv: `.\.venv\Scripts\Activate.ps1` (prompt shows `(.venv)`).
- Python 3.12 venv. Installed: pdfplumber 0.11.10, pypdfium2 5.15.2, Pillow, ImageHash 4.3.2,
  openpyxl, docx2pdf, pywin32 312, **PySide6_Essentials 6.12.0**, pyflakes, python-docx,
  requests, truststore. Word + Excel installed (COM).
- pip on company network: `--timeout 120 --retries 10`.
- PowerShell: names with spaces → quote paths; names with `[ ]` → `-LiteralPath`; copy specs to
  short names (`data\med_old.docx`). `data/`, `output/`, *.pdf, *.docx, *.doc are git-ignored.
- `.gitignore` (current): *.pdf *.docx *.doc, data/, output/, .venv/, pyvenv.cfg, __pycache__/,
  *.pyc, build/, dist/, `~$*`. Note `data/` also matches `tests/data/`.
- **VIO token**: environment variable `VIO_API_TOKEN` (User). If Python can't see it in a running
  VS Code: `$env:VIO_API_TOKEN = [Environment]::GetEnvironmentVariable("VIO_API_TOKEN","User")`.

---

## 3. Engine – current state on GitHub (all typed in and verified)

All engine steps A–D and earlier fixes are on `main` and **identical to the tested reference**
(only message wording differs in excel.py: `info;`, `rw`, `thier`).

### 3.1 Files

| File | Role (current) |
|---|---|
| convert.py | .docx/.doc → PDF via Word (fields locked). Passwords: `ensure_password`/`get_password` (terminal); **V1.0 step 1 adds `needs_password`/`try_password` (app pop-up) – not typed yet** |
| extract.py | pdfplumber chars (font, size, colour, **bold** = font name or fake bold fill+stroke via pdfium, **italic** = font name or slant), thin lines, images, tables. **`_table_page`**: strikethrough/underline lines and coloured highlight boxes are removed before table detection (they split cells) |
| clean.py | Blocks (heading/para/cell/hf); strikethrough (35–65 % band) + **underline** (80–130 %, not table borders); **running header/footer by repetition** (`_running_ids`: same height near top/bottom 20 % on ≥ half the pages, digits ignored, not inside tables); table continuation across pages only if **no text between** and same columns; **rows split across pages joined** (`split_row`: each filled cell continues a cell above that ran to the right edge); paragraphs continue onto a new page only if the last line reaches the document's right margin (`text_right` = widest body line or table); heading check uses `text_right` |
| compare.py | **Units** = paragraph or whole table row; aligned with SequenceMatcher on (section heading, text) then DP pairing (`_pair_range` + `_make_unit_sim`: row text ratio or same-column cells); **moved** unchanged unique units paired (`_pair_moved`, MOVED_MIN_CHARS 4) + moved-and-edited pairs (MOVED_SIM 0.85, mutual best); cells inside paired rows lined up by text then position; **renumbering ignored** (`_unnumbered`, headings + 表/図/Table/Fig captions); completely different text in the same place = one replacement (`「Released」 → 「DRAFT」`); **formatting diff** (strike, bold, italic, underline, colour, size; systematic differences ignored with WARNING); **whole-word details** (`_word_span`, `_text_details`, `_txt_spaced`: 「3235」→「3300」, 「2023.10.16」→「2026.8.18」, 「4.1 5.0」); format notes inside changed words dropped; removal markers only for pure removals; deleted/new section = one row incl. pictures (`_fold_images`); **TOC grouped per page** (`_group_toc`, label 目次); numbering by `min(page, top)` |
| images.py | Pictures paired in document order by look (pHash incl. rotations/flips) **+ same-section bonus** (HEADING_BONUS 0.25); .docx rotation read from Word |
| shots.py | Renders crops with red boxes/markers; **no old picture for New rows** |
| excel.py | Two-pass write (rows/heights first, then pictures); **`_snap_to_row`** moves each picture until Excel puts it in its own row (fixes drift on long sheets – not yet confirmed on a real large file); **N/A** in column F for New rows |
| tests/check_answer_key.py | Supports `"text"` entries (formatting). Answer keys: answer_key.json (6), answer_key_format.json (7) |
| tests/make_format_pair.py | Builds the formatting pair |
| tools/vio_check.py, vio_chat_test.py, vio_batch_test.py | VIO connection / single / batch tests (no tokens inside) |

### 3.2 Engine settings (for the "engine block", App_Design 8.1)
compare: MATCH_MIN 0.5, COL_BONUS 0.1, MAX_PAIR_CELLS 250000, SIZE_TOL 0.3, COLOR_TOL 0.15,
SYSTEMATIC_RATIO 0.3, SYSTEMATIC_MIN 200, MOVED_MIN_CHARS 4, MOVED_SIM 0.85 ·
clean: HF_TOP_RATIO 0.075, HF_BOTTOM_RATIO 0.925, RUNNING_ZONE 0.20, RUNNING_Y_TOL 4.0,
STRIKE_BAND (0.35, 0.65), STRIKE_Y_TOL 1.5, STRIKE_X_GAP 1.0, STRIKE_MIN_COVER 0.5,
UNDERLINE_BAND (0.80, 1.30), BORDER_TOL 1.5, HEADING_MAX_LEN 40 ·
extract: MIN_IMAGE_PT 6.0, ITALIC_SKEW 0.1 · images: SAME_DIST 6, ROT_DIST 16, PAIR_DIST 20,
FLIP_MARGIN 4, SIZE_TOL 0.05, HEADING_BONUS 0.25 · shots: DPI 150, PAD_Y 40, BOX_PAD 2, BOX_WIDTH 3 ·
excel: PIC_MARGIN 3, MAX_ROW_HEIGHT 409, MIN_ROW_HEIGHT 20.

### 3.3 Commands
```
python -m cpa.excel OLD NEW --out output\<name>      # full run -> CPA_<new>.xlsx, changes.txt, shots\
python -m cpa.compare OLD NEW --out output\<name>    # changes.txt only
python tests/check_answer_key.py                      # 6 passed
python tests/check_answer_key.py --old tests/data/Format_Old.docx --new tests/data/Format_New.docx --key tests/expected/answer_key_format.json --out output/check_format   # 7 passed
```

---

## 4. Test results

### 4.1 Real files (user's laptop)
- Answer key pair: **6 passed**; formatting pair: **7 passed** (after every step).
- Medium file runs: "medium2" was the last real run **before** steps A–D. Problems seen then and
  fixed since (on synthetic copies): 264 pictures in wrong rows (snap fix, unconfirmed), revision
  table header rows falsely Deleted (running header + repeated header rows), cover table rows
  falsely Deleted (strikethrough/highlight splitting cells), identical-text Deleted+New (moved
  content), char fragments in descriptions.
- **Not yet done: rerun the real medium/large files with the final code** (the user chose to
  move on to the app first). Ask for: change count, `info:` lines, WARNINGs, screenshots of any
  wrong rows.

### 4.2 Synthetic tests (Claude's environment; generators in tests/synthetic/)
All pass with the current engine: spec 27 · big 48/48 found, 0 missed, 0 extra · medium 18 ·
revision 10 · header 4 · split rows 4 · cover 3 · identical rows 1/1/1 · strike+text 2 ·
formatting 7 · same file vs itself 0. History: big test went 17 false rows → 0; header test
13 → 4; split rows 33 → 4; cover 6 → 3.

### 4.3 Known limits
- Indentation-only changes are not reported (Step 2d).
- When one of several **identical rows** is added/deleted, the red box may sit on a neighbouring
  identical row (count and type are correct).
- Rotated picture in PDF-only input shows as "image content changed" (Word input reads rotation).
- Table formatting (2c), spacing/「要確認_Minor」 (2d), pixel safety net (2e), OCR: not done.
- Excel picture snap fix not yet confirmed on a real large file.

---

## 5. App – decisions and status

**Full agreed design: docs/App_Design.md (425 lines).** Mockups: docs/mockups/*.png.

Key decisions: modern styled UI, light + dark toggle, English UI, 1366×768 must fit, fonts
Segoe UI / Yu Gothic UI. Main page with A quick-start drop zone, B continue card, C system check,
D this-month stats. New session = one page (name, old/latest file boxes, **⇅ Swap**, Auto/Advanced
switch, more options). Quick start guesses old/latest by version numbers → old/new words → modified
date; nothing starts until Start. Review page = list (thumbnail focused on red boxes, chips) +
big Old/Latest pictures side by side (white margins trimmed in preview only), What changed,
Comment (app only), Approve/Reject/Follow-up, everything starts **Pending**, editable B/C/description,
batch actions (Ctrl/Shift/Ctrl+A, batch bar ✓ ✕ ⚑ ⋯), undo/redo surviving restart, shortcuts
A R F N P Ctrl+Z/Y Ctrl+F Ctrl+E (letters only when no text box has focus). Export: template check
→ pending/follow-up warning (go back / export approved only) → F1/G1 form → Excel (approved only,
N/A, snapped pictures). Sessions auto-saved in `output/sessions/<id>/` (session.json + shots),
spec files **linked not copied**. Advanced mode (V1.2): crop = the change, Continue/Start over in
Advanced, Source Auto/Manual/AutoEdited. Merge (V1.1) combines rows into one picture.
Stats SQLite (no document text), engine block with all settings, origin_ids, reject reasons (V1.1),
missed_check_done, review-based metrics (approval rate, estimated recall).

Releases: **V1.0** Auto + Review + Export + sessions + stats · V1.1 crop editor/merge/reject
reasons · V1.2 Advanced mode + smart sync · V1.3 settings/help · later VIO, MLflow, OCR.
Not planned: multi-user/roles/audit trail.

### 5.1 V1.0 build plan (one file per message; status)

| Step | File(s) | Status |
|---|---|---|
| **1** | convert.py (`needs_password`, `try_password` above `ensure_password`), new `cpa/app/__init__.py`, new `cpa/app/theme.py` | **Sent to the user, not typed yet.** Tested reference: reference/cpa/app/theme.py (screenshot check light/dark OK, readable text on accent buttons = `on_accent`) |
| 2 | cpa/app/worker.py (background jobs) | Reference prototype in reference/cpa/app/worker.py (run_compare / run_export / start) |
| 3 | cpa/app/session.py (model, auto-save, undo) | to build |
| 4 | cpa/app/stats.py (SQLite + engine block) | to build |
| 5 | cpa/app/home.py (Main page A–D) | to build (mockup ready) |
| 6 | cpa/app/new_session.py (quick start, swap) | to build |
| 7 | cpa/app/review.py (list, preview, actions, batch, filters) | to build |
| 8 | cpa/app/export.py (template check, warning, F1/G1, Excel) | to build |
| 9 | cpa/app/main.py + `__main__.py` (window, pages, theme toggle) | to build → `python -m cpa.app` |

A first plain prototype (start page + simple review + export, `python -m cpa.app`) was built and
tested earlier but **rejected by the user as too plain**; only its worker.py is reused.
Testing method in Claude's env: `QT_QPA_PLATFORM=offscreen`, drive the window in a script,
`widget.grab().save(png)`, stand-in for Excel COM; needs system libs libegl1 libgl1 libxkbcommon0
libfontconfig1 libdbus-1-3.

Step-1 exact text for convert.py:
```python
def needs_password(pdf_path) -> bool:
    """True if the PDF cannot be opened without a password that is not known yet."""
    import pypdfium2 as pdfium

    if _key(pdf_path) in _PASSWORDS:
        return False
    try:
        pdfium.PdfDocument(str(pdf_path)).close()
        return False
    except pdfium.PdfiumError:
        return True

def try_password(pdf_path, pw: str) -> bool:
    """Remember *pw* for this PDF if it opens it (for the app's password box)."""
    import pypdfium2 as pdfium

    try:
        pdfium.PdfDocument(str(pdf_path), password=pw).close()
    except pdfium.PdfiumError:
        return False
    _PASSWORDS[_key(pdf_path)] = pw
    return True
```

---

## 6. VIO (company AI) – status

See App_Design.md section 15. Summary: supervisor approved sending confidential spec content to VIO
(keep the message). API = **agents only** (direct LLM retired), OpenAI-compatible
`/chat/completions`, `/models`, Bearer token (expires, default 30 days), base URL
`https://vio.automotive-wan.com:446` (company network only; appears AWS-hosted), certificates via
`truststore`. Token visible agents: **Compare Expert** (not the user's – test only). Plan: user
creates own agent **CPA Translator** (instructions in the chat history: translate JP→EN, keep codes/
numbers/units, "shall", no extra text). Measured: cold 57 s, warm ~21 s, **10 lines in one request
20.6 s** → batching (~20 per request) in the background; numbered-lines format worked perfectly.
VIO comes **after V1.0**.

---

## 7. MLOps / LLMOps plan (agreed)

- **Engine (MLOps):** synthetic evaluation set (tests/synthetic + answer keys) → evaluation script
  (recall, precision, false rows, seconds) → MLflow runs per engine version/profile (params =
  all settings) → regression gate (later GitHub Actions) → real-world monitoring from Review
  decisions (stats SQLite, no text) → feedback loop (new synthetic case per real problem).
- **VIO (LLMOps):** prompts versioned in `prompts/`, per-request log (agent, prompt version,
  batch size, seconds, tokens, accepted/rejected/edited), compare prompt versions by acceptance.
- Real-world metrics are review-based (approval rate ≠ precision; estimated recall only with
  missed_check_done); profiles are compared on the evaluation set, not on real logs.

## 8. Career notes (discussed)

Honest assessment: strong intern project if finished, used by the team and measured. CV bullets
drafted in chat (Python/PySide6 tool, engine with row matching/moved/renumbering handling,
synthetic test suite 17→0 false rows with 100 % recall, MLflow + regression gate, SQLite review
logging, VIO integration with prompt versioning). Cautions: confidentiality (no customer names,
real spec content or screenshots; use synthetic specs for demos; ask supervisor what may be shown),
own the code (be able to explain each fix), measure impact (time per CPA before/after, users,
approval rate, VIO acceptance).

---

## 9. Next actions when resuming

1. Ask whether V1.0 step 1 has been typed and pushed; check it (theme.py must be identical to
   reference/cpa/app/theme.py; `{{ }}` braces are the usual mistake; convert.py helpers).
2. Continue V1.0 step 2 (worker.py), then steps 3–9 – build + screenshot-test each first.
3. When convenient: user reruns real medium/large files with the final engine and reports.
4. Optional in parallel: tools/evaluate.py + MLflow on tests/synthetic (MLOps Part A).
5. After V1.0: VIO (CPA Translator agent, batching) per App_Design 15.
