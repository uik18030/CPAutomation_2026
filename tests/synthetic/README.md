# Synthetic test pairs (no confidential content)

Each script builds an old/new **Word** pair with known changes into
`output/synthetic/<script name>/` (git-ignored). Needs `pip install python-docx`.
Font: `MS Gothic` on Windows (change with the `CPA_TEST_FONT` environment variable).

```
python tests/synthetic/make_spec_pair.py
python -m cpa.compare output\synthetic\make_spec_pair\spec_old.docx output\synthetic\make_spec_pair\spec_new.docx --out output\check_spec
```
(.docx inputs are converted with Word; the same commands work with the PDFs.)

| Script | Pair | Expected rows | What it tests |
|---|---|---|---|
| make_spec_pair.py | spec_old / spec_new (9 pages) | **27** | Realistic spec: page header with Ver/Page x of y, approval table, cover box with struck + highlighted values, 改訂履歴 moved to page 2, 目次, headings 1. / 1.1, bullets • ➢ ①, tables (one with rows split across pages), pictures, new section, renumbering |
| make_big_pair.py | big_old / big_new (45 pages) | **48** (writes `expected.json`) | 70 sections, 48 planted changes + moved paragraphs + a deleted section; score with `python tools/score_changes.py output\synthetic\make_big_pair\expected.json <out>\changes.txt` → 48 found, 0 missed, 0 extra |
| make_medium_pair.py | med_old / med_new (4 pages) | **18** | Cover table, revision table moved, refs table, picture rotated, signal table, bullets bold/colour/deleted/added, moved paragraph, new section |
| make_revision_pair.py | rev_old / rev_new | **10** | Revision table with merged Ver./Date cells moved after 序章, inserted/deleted rows, 3.2→3.3, Released→DRAFT |
| make_header_pair.py | hdr_old / hdr_new | **4** | Tall running page header (below 7.5 %), long tables with repeated header rows |
| make_split_rows_pair.py | rs_old / rs_new | **4** | Same, with table rows breaking across pages |
| make_cover_pair.py | cov_old / cov_new | **3** | Strikethrough + highlight inside cover table cells must not split cells |
| make_identical_rows_pair.py | dup_old vs dup_a / dup_b / dup_c | **1 each** | Identical rows: deleted / changed / added; heading recognised on a page with only short lines |
| make_strike_text_pair.py | reg_old / reg_new | **2** | Strikethrough removed + T_dis_RSR→RSP; shared underline not reported |

Formatting pair: `tests/make_format_pair.py` + `tests/expected/answer_key_format.json` → 7 passed.

Expected numbers were measured with LibreOffice-made PDFs. Word-made PDFs and the MS Gothic font
can lay out pages differently; small differences are worth a look, big ones are a regression.
