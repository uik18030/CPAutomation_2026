# CPA Tool – App Design (agreed)

> Status: agreed design for the desktop app. Engine details: docs/Handover.md.
> Mockups (V1.0): Main page, New session, Review page (single + batch), light and dark.
> Working rules: never push from Claude; every change is pasted as find/replace and typed in by the user.

---

## 1. Purpose

Compare an **old** and a **latest** spec (PDF / .docx / .doc), let a reviewer check every detected
change, and fill the company **CPA Excel template** (sheet `CPA_JPN`) with headings, change types and
**old/new screenshots with red boxes**.

```text
Old spec + Latest spec
        ↓
Auto mode (tool finds changes)   or   Advanced mode (reviewer marks changes)
        ↓
Review page  (approve / reject / edit – every row is checked here)
        ↓
Export → CPA_JPN sheet (approved rows only, screenshots in F/G)
```

Both modes always end on the **same Review page**.

---

## 2. Users and environment

| Item | Decision |
|---|---|
| Users | The author only for now (run with `python -m cpa.app`); `.exe` later |
| OS | Windows; Microsoft **Word** (for .docx/.doc) and **Excel** (for export) installed |
| UI language | English (spec text stays Japanese) |
| Screen | Must fit **1366 × 768** without horizontal scrolling |
| Fonts | Segoe UI (English), Yu Gothic UI / Meiryo (Japanese) |
| Theme | Modern styled **light** theme + **dark** toggle (top bar) |

---

## 3. Releases

| Release | Contents |
|---|---|
| **V1.0** | Main page (A–D), New session (quick start, swap), **Auto mode**, Review page (statuses, edits, comments, batch, undo, shortcuts, filters, stats), sessions with auto-save, Export with screenshots + F1/G1 form + template check, local stats recording |
| **V1.1** | **Crop editor** on the Review page: edit / re-crop / delete / duplicate / **merge** rows |
| **V1.2** | **Advanced mode**: side-by-side viewer with **smart sync scroll**, manual crops; **Switch to Advanced** from Review |
| **V1.3** | Settings page (engine profiles, template, theme, defaults), Help pages, "files look swapped" warning |
| Later | VIO (AI) features, MLflow / analytics dashboard, OCR (only if needed) |
| Not planned | Multi-user review, roles/permissions, audit trail (need a server) |

---

## 4. Main page

Top bar: logo + **CPA Tool**, Dark/Light toggle, Settings, Help.

| Area | Content |
|---|---|
| Hero | Title "Compare spec revisions and build the CPA sheet", one-line description, **New session**, **Open session…** |
| **A. Quick start** | Big drop zone: "Drop the old and new spec here" (see 5.1). Click = file picker (select 2 files) |
| **B. Continue** | Last unfinished session: name, pending / follow-up counts, progress bar, "at change #N", **Resume** |
| Recent sessions | Last 4 sessions: name, files, date, chip "Done" or "12/40"; **View all** |
| **C. System check** | Status lights: Excel template (found + valid), Microsoft Word, Microsoft Excel, output folder writable, VIO (off · company network only). **Re-check** button |
| **D. This month** | Sessions, changes reviewed, **approval rate** (review-based), average review time, each with a small trend line; **estimated recall** added in V1.2. Note: "engine vX · recorded locally, no document text" |

---

## 5. New session page (one page)

Fields: **Session name**, **Old spec** box, **⇅ Swap old / latest**, **Latest spec** box,
**Mode** switch (Auto / Advanced), "More options" (output folder, template), **Start comparison**.

- Each file box shows file name, pages, type, folder; 🔒 if password protected; **Change…** button.
- Files can be dropped on each box (replaces that one file).
- Session name is filled from the file names (e.g. "spec Ver5.0 → Ver6.0"); editable.
- **Mode**: default = last used mode (Settings can fix a default). Advanced is greyed out until V1.2.
- Nothing starts until **Start comparison** is pressed.

### 5.1 Quick start (drop on the main page)

- **2 files dropped** → guess old / latest (drag order is **not** used):
  1. version numbers in the name (`Ver5.0` < `Ver6.0`, `v3` < `v4`)
  2. words `old` / `new`, `旧` / `新`
  3. file modified date (older = old spec)

  → opens New session filled in; the user checks it (Swap if wrong) and presses Start.
- **1 file dropped** → fills Old spec; New session waits for the latest spec.
- Checks: only PDF / Word ("Only PDF or Word files can be compared."), at most 2 files,
  same file twice → warning, password → 🔒 and asked on Start (never saved).

### 5.2 Swap

- One click swaps old ↔ latest and updates the session name.
- After a session has results: "Swap and compare again? Your review decisions for this session will be cleared."
- V1.3: if the cover version/date of the latest spec is older than the old spec → "These look swapped. Swap them?"

### 5.3 Running the comparison

- Runs in the background (window stays responsive): reading old file → reading latest file →
  comparing → screenshots; progress text + bar; errors shown in a message box.
- Password-protected PDFs: password box, wrong password asks again.

---

## 6. Review page

### 6.1 Layout

```text
Top bar:  CPA Tool / <session name>                       Dark   ⚙   [Export to Excel]
Stats:    27 changes · Approved 9 · Pending 16 · Follow-up 1 · Rejected 1 · ▬▬▬ 37% reviewed
          [search text, heading, #no]  [All statuses ▾] [All types ▾] [All headings ▾]
┌ List (≈380 px) ───────────┐ ┌ Change #13  Modified  Pending   old p.4 → new p.4   Sheet/Header [条件信号] [Modified ▾] ┐
│ [thumb] #5 表紙            │ │ ┌ OLD SPEC ──────────────┐ ┌ LATEST SPEC ───────────┐                       │
│   Modified  Approved      │ │ │  picture               │ │  picture               │                       │
│   added 「and HUD」         │ │ └────────────────────────┘ └────────────────────────┘                       │
│ ...                       │ │ ┌ WHAT CHANGED ────────────┐ ┌ COMMENT (app only) ───┐                     │
│                           │ │ └──────────────────────────┘ └────────────────────────┘                     │
│ [3 selected ✓ ✕ ⚑ ⋯]      │ │ [↑ Previous (P)] [Next (N) ↓]          [✕ Reject (R)] [⚑ Follow-up (F)] [✓ Approve (A)] │
└───────────────────────────┘ └──────────────────────────────────────────────────────────────────────────────────┘
```

- **List row**: thumbnail (focused on the red boxes), #No., heading, type chip, status chip, one-line summary.
- **Pictures**: Old and Latest side by side; white margins trimmed **in the preview only**
  (Excel gets the full crop). New rows show **N/A** in Old Spec.
- **What changed**: the detail lines (editable text, app only).
- **Comment**: free text, **app only** (saved in the session, not exported).
- Page numbers shown as "old p.12 → new p.13".

### 6.2 Statuses

`Pending` (start) → `Approved` / `Rejected` / `Follow-up`. Everything starts **Pending**.

- V1.1: rejecting asks for a **reason** (one click, default *False alarm*): *False alarm* /
  *Not needed* / *Duplicate*. Batch reject applies one reason to all selected rows.
- Check box "**I checked for missed changes**" (sets `missed_check_done`, see 8.3).

### 6.3 Editable per row (V1.0)

| Field | Exported |
|---|---|
| Sheet/Header (column B) | yes |
| Type: New / Deleted / Modified (column C) | yes |
| Description text | no (app only) |
| Comment | no (app only) |

### 6.4 Batch actions

- Select: Ctrl+click, Shift+click, Ctrl+A (all **shown** rows – combine with filters).
- Batch bar: **✓ Approve**, **✕ Reject**, **⚑ Follow-up**, **⋯** (Set heading…, Set type…, Add comment…).
- "Approve all / Reject all" (shown rows) always asks: "Reject 48 changes?"
- **Undo / Redo**: Ctrl+Z / Ctrl+Y for every action incl. batch; survives restart (stored in the session).

### 6.5 Navigation, filters, shortcuts

| Key | Action |
|---|---|
| A | Approve |
| R | Reject |
| F | Follow-up |
| N or ↓ | Next change |
| P or ↑ | Previous change |
| Ctrl+Z / Ctrl+Y | Undo / Redo |
| Ctrl+A | Select all shown |
| Ctrl+F | Search |
| Ctrl+E | Export |

Single-letter keys work only when no text box has focus.
Filters: status, type, heading; search in text, heading and #No.

### 6.6 Switch to Advanced (V1.2)

Button on the Review page:

- **Continue in Advanced** (default): keep the auto rows, open the side-by-side viewer with the auto
  changes drawn as boxes, add missing changes as **Manual** rows.
- **Start over in Advanced**: all auto rows set to Rejected (undoable), mark everything by hand.

Same session; rows keep their source (**Auto** / **Manual**).

---

## 7. Export

1. **Template check**: sheet `CPA_JPN` exists, header row 1 = No. / Sheet/Header / Changes Description /
   DOORS ID / ID changes / Old Spec / Latest Spec. If not: "Template structure invalid" + what is missing.
2. **Pending / Follow-up warning**: "N changes still pending, M follow-up" → **Go back** or **Export approved only**.
3. **Header info form** (F1 / G1): old and latest – ticket, variant, received date (dd/mm/yyyy);
   typed values written in red, rest black, same layout as the template.
4. **Save as** (default `<output>/<session>/CPA_<latest file>.xlsx`).
5. Written through **Excel (COM)**, two-pass, each picture snapped into its own row.
6. Done → "Saved" with **Open folder** / **Open file**.

### 7.1 Column mapping (only **Approved** rows, in document order)

| Column | Content |
|---|---|
| A No. | 1, 2, 3 … (renumbered after filtering) |
| B Sheet/Header | Heading (edited value if changed) |
| C Changes Description | New / Deleted / Modified |
| D DOORS ID, E ID changes | empty (template dropdowns kept) |
| F Old Spec | Old screenshot with red boxes; **N/A** for New rows |
| G Latest Spec | Latest screenshot with red boxes / removal markers |

---

## 8. Sessions

- **Auto-save** after every action (no Save button).
- Stored in `output/sessions/<session id>/` (git-ignored):

| File | Content |
|---|---|
| `session.json` | name, mode, **paths** of old/latest files (+ size, modified time to detect changed files), **engine block** (see below), rows with status / edits / comments / **source / origin_ids**, `missed_check_done`, undo history, F1/G1 info |
| `shots/` | screenshots (NN_old.png / NN_new.png) |
| `changes.txt` | engine report |

- Spec files are **linked, not copied** (fewer copies of confidential documents).
- If a linked file is missing or changed: "spec_new.pdf was moved or changed – locate it / compare again".
- Converted PDFs of Word files are kept in the session folder so reopening does not need Word.

### 8.1 Engine block (exact configuration of every run)

All engine settings are collected **automatically** from the engine modules (every UPPER_CASE
setting: `MATCH_MIN`, `MOVED_SIM`, `SIZE_TOL`, `COLOR_TOL`, `MOVED_MIN_CHARS`, `COL_BONUS`,
`SYSTEMATIC_RATIO`, `STRIKE_BAND`, `UNDERLINE_BAND`, `RUNNING_ZONE`, `HEADING_BONUS`, …),
so a new or renamed setting is never forgotten.

```json
"engine": {
  "version": "0.9.0",
  "git_commit": "3753592",
  "profile": "Standard",
  "settings": { "MATCH_MIN": 0.5, "MOVED_SIM": 0.85, "SIZE_TOL": 0.3, "COLOR_TOL": 0.15, "...": "..." }
}
```

- Running from VS Code: `git_commit` read from git. Built `.exe`: version + commit written into
  the build (no git on other PCs).

### 8.2 Rows: source and origin

| Field | Values / meaning |
|---|---|
| `source` | **Auto** (from the engine) · **Manual** (drawn by the reviewer, V1.2) · **AutoEdited** (auto row changed with the crop editor, V1.1) |
| `edits` | what was changed on an auto row: `crop`, `heading`, `type`, `merged` (empty = untouched) |
| `origin_ids` | the engine change numbers a final row came from, e.g. merged row 12 = `[12, 13, 14]`; Manual rows = `[]` |
| `reject_reason` (V1.1) | **False alarm** (engine error) · **Not needed** (correct but not wanted in the CPA) · **Duplicate** |

### 8.3 Missed-change check

`missed_check_done` (session flag) = the reviewer actually looked for changes the tool missed:
set automatically when **Continue in Advanced** is used (V1.2), or by ticking
"I checked for missed changes" on the Review page. Without it, estimated recall is not computed.

---

## 9. Recorded statistics (for the main page and MLOps)

Local SQLite file `output/cpa_stats.sqlite`. **No document text is stored**
(only counts, types, headings, settings and decisions).

| Table | Fields |
|---|---|
| runs | run_id, session_id, time, engine_version, git_commit, profile, **settings_json**, pages_old, pages_new, seconds, changes_found |
| decisions | run_id, final_row_id, **origin_ids**, **source**, **edits**, type, heading, status, **reject_reason** (V1.1), review_seconds |
| sessions | session_id, created, finished, total_review_seconds, **missed_check_done** |

Settings are stored as **one JSON field** (`settings_json`) plus a few key columns, so adding or
renaming an engine setting never breaks the database.

### 9.1 Metrics (all labelled "review-based")

| Metric | Formula | Available |
|---|---|---|
| **Approval rate** | approved Auto/AutoEdited rows ÷ auto changes found | V1.0 |
| **False-alarm rate** | rows rejected as *False alarm* ÷ auto changes found | V1.1 (reject reasons) |
| **Manual addition rate** | Manual rows ÷ final rows | V1.2 |
| **Estimated recall** | auto rows kept ÷ (auto rows kept + Manual rows) – only sessions with `missed_check_done`, never Advanced-only sessions | V1.2 |
| **Review time** | total review seconds per session / per change | V1.0 |

- "Approval rate" is **not** called precision: a correct detection can be rejected for team policy
  ("Not needed"). Once reject reasons exist, only *False alarm* counts against the engine.
- Estimated recall is shown as **"estimated"**: the tool cannot know changes nobody found.
- Main page "This month" (D) shows approval rate (V1.0); estimated recall is added in V1.2.

### 9.2 Use for MLOps

| Purpose | Data | Where |
|---|---|---|
| **Monitoring** ("is it still good on real specs?") | runs + decisions from real sessions | stats DB → main page, later a dashboard |
| **Comparing engine versions / profiles** (Standard / Conservative / Aggressive) | the **same synthetic test documents with known answers** (true precision and recall) | evaluation script → **MLflow** |
| **False-positive analysis** | Auto rows rejected as *False alarm*, by type / heading | stats DB |
| **False-negative analysis** | Manual rows, by type / heading | stats DB |

Profiles are compared on the evaluation set, not on real-world logs: real sessions use different
documents, so differences there can come from document difficulty instead of the profile.
Real-world numbers show whether the chosen profile still works in practice.

### 9.3 Needed per release

| Release | Recorded |
|---|---|
| V1.0 | engine block (version, commit, profile, all settings), `source` = Auto, statuses, edits (heading/type), review time, approval rate |
| V1.1 | AutoEdited, `edits` incl. crop/merged, `origin_ids` for merges, **reject reasons**, false-alarm rate |
| V1.2 | Manual rows, `missed_check_done`, manual addition rate, **estimated recall** |
| Later | MLflow runs for engine versions/profiles on the evaluation set; dashboard |

---

## 10. V1.1 – Crop editor

On the Review page, per row: **Edit crop** (resize / move area, zoom, pan, Save / Cancel),
**Re-crop** (draw a new area), **Add / move / delete red boxes**, **Duplicate row**,
**Merge rows** (multi-select → one row; heading of the first row, type Modified unless all equal;
close changes → one crop with all boxes, far apart → stacked crops; warning if the picture would
shrink a lot because of Excel's 409 pt row limit), **Split** (undo merge).
Edited auto rows become **AutoEdited** (`edits` records crop / merged); merged rows keep the
original engine numbers in `origin_ids`.

## 11. V1.2 – Advanced mode

- Side-by-side viewer of the whole old and latest document.
- **Smart sync scroll**: the same paragraph stays lined up on both sides (uses the engine's matching),
  even when the latest spec has added pages.
- Draw an old crop and/or a latest crop → one **Manual** row (`source = Manual`, crop = the change; no detection).
- Rows go to the same Review page.
- Using **Continue in Advanced** after an Auto run sets `missed_check_done` (used for estimated recall).

## 12. V1.3 – Settings and Help

- Settings: default template, output folder, default mode, theme, engine profile
  (**Standard** default; Conservative / Aggressive only after they are tested on the evaluation set
  in MLflow), advanced engine values under "Advanced". The chosen profile and all values are saved
  with every run (8.1).
- Help: user guide, shortcuts, known limitations, version notes.

---

## 13. Look and feel

| Token | Light | Dark |
|---|---|---|
| background | #F4F6F9 | #11151B |
| surface | #FFFFFF | #1A2028 |
| border | #E2E7EE | #2C3542 |
| text / muted | #1A2230 / #5E6A7A | #E7EBF1 / #9AA6B6 |
| accent | #1F4E8C | #6C9CF0 |
| approved | #1E8A57 | #4CC38A |
| rejected | #C2362B | #F0766B |
| follow-up | #B7791F | #E3AE52 |
| pending | #6B7686 | #9AA6B6 |

Red is reserved for change boxes in screenshots and the Reject action. Rounded cards (10 px),
status chips, one primary (accent) button per screen.

---

## 14. Code structure

```text
cpa/                 engine (extract, clean, compare, images, shots, excel, convert)
cpa/app/
  __main__.py        python -m cpa.app
  main.py            window, page switching, top bar, theme
  theme.py           colour tokens + stylesheet (light / dark)
  home.py            Main page (A–D)
  new_session.py     New session page, quick start, swap
  review.py          Review page (list, preview, actions, batch, filters)
  session.py         session model, auto-save, undo/redo
  stats.py           SQLite recording + "This month" numbers
  export.py          template check, warning, F1/G1 form, Excel writing
  worker.py          background jobs
```

The engine stays independent of the app: engine fixes need no app changes.

---

## 15. AI (VIO) – after V1.0

### 15.1 Status

| Item | Status |
|---|---|
| Approval | Supervisor confirmed (2026-11-24) that confidential spec content may be sent to VIO (keep the written message) |
| API | VIO API, **agents only** (direct `VIO:` models retired). OpenAI-compatible `POST /chat/completions`, `GET /models` |
| Auth | Bearer token, created in VIO (🔌 VIO API) or via `/api/auth/authentication`; default lifetime 30 days (max 365) |
| Network | Company network only; VIO appears to run in the company's AWS environment |
| Certificates | Windows certificate store via `truststore` (`pip install truststore`) |
| Connection test | `tools/vio_check.py` → HTTP 200; agents available to this token: `Compare Expert` (not ours) |

### 15.2 Measured (Compare Expert, test sentences only)

| Test | Time | Tokens |
|---|---|---|
| 1 sentence, cold start | 57.4 s | 30 in / 19 out |
| 1 sentence, warm | ~21.5 s | – |
| **10 sentences in one request** | **20.6 s** | 316 in / 184 out (≈ 30–50 per change) |

→ ~20 s is fixed overhead per request: **batching** is required. Output format (numbered lines,
"shall", codes/numbers/units kept, nothing extra) was exactly as asked.

### 15.3 Design

- **Off by default**; the tool works fully without VIO. Shown in System check (C) as on/off.
- Token in `VIO_API_TOKEN` (also accepts `API_TOKEN`); base URL in settings / `VIO_BASE_URL`
  (not hard-coded). Never in code or git. Expired / invalid token (401) → clear message
  "VIO token expired – create a new one in VIO → 🔌 VIO API".
- **Agent**: `Compare Expert` for testing only. Before real use, create our own agent
  **`CPA Translator`** (we control and version its instructions).
- **Translate all**: changes sent in **batches of ~20** (or by text length), in the background with
  progress; the Review page stays usable. **Translate row**: single request with a spinner; rows
  clicked while waiting join the next batch. First request of the day: "Connecting to VIO
  (first request can take up to a minute)…".
- Format: numbered lines in, numbered lines out; answers are checked (missing / extra line →
  that batch is retried); leading blank lines trimmed.
- Results saved in the session (asked once per row); shown as **suggestions**, marked as AI,
  never written into Excel automatically. Where accepted translations go (app only / extra column /
  extra sheet) is decided later.
- **LLMOps log** per request: agent name, prompt version (`prompts/translate_vN.txt`), batch size,
  seconds, tokens in/out, result per row (accepted ✓ / rejected ✗ / edited). No document text in
  the statistics database.
