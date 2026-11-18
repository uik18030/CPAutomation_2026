"""Step 1e: write the changes into the CPA template (sheet CPA_JPN).

Excel itself (via COM) is used so the template stays exactly as it is:
dropdowns (DOORS ID / ID changes), the List&guideline sheet, filter and
freeze panes are kept. Only rows 2+ of CPA_JPN are filled.

Run:
    python -m cpa.excel tests/data/File_Old_update.docx tests/data/File_New_update.docx
Writes output/CPA_<new file name>.xlsx (+ output/changes.txt and output/shots/)
"""
from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from PIL import Image

from cpa.clean import load
from cpa.compare import Change, compare, write_changes
from cpa.shots import make_shots

# ------------------------------------------------------------------ settings
SHEET = "CPA_JPN"
FIRST_ROW = 2
COL_NO, COL_HEADER, COL_DESC, COL_OLD, COL_NEW = 1, 2, 3, 6, 7
FILTER_RANGE_PREFIX = "A1:G"
FILTER_LAST_ROW = 200       # template filter covers A1:G200
PIC_MARGIN = 3.0            # pt between cell border and picture
MAX_ROW_HEIGHT = 409.0      # Excel limit (pt)
MIN_ROW_HEIGHT = 20.0

XL_MOVE_AND_SIZE = 1        # picture moves/hides with its row (needed for filtering)
XL_TOP = -4160

def _fit(png: Path, max_w: float, max_h: float) -> tuple[float, float]:
    """Picture size in points that fits max_w x max_h, keeping the aspect ratio."""
    with Image.open(png) as im:
        w, h = im.size
    scale = min(max_w / w, max_h / h)
    return w * scale, h * scale

def write_excel(changes: list[Change], shots: dict[int, tuple[Path | None, Path | None]], template: str | Path, out_path: str | Path) -> Path:

    import pythoncom
    import win32com.client

    template = Path(template).resolve()
    out_path = Path(out_path).resolve()
    if not template.exists():
        raise FileNotFoundError(f"Template not found: {template}")
    out_path.parent.mkdir(parent=True, exist_ok=True)
    try:
        shutil.copyfile(template, out_path)
    except PermissionError:
        raise PermissionError(f"Close this file in Excel first: {out_path}") from None

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(str(out_path))
        ws = wb.Worksheets(SHEET)

        col_w = min(ws.Columns(COL_OLD).Width, ws.Columns(COL_NEW).Width)
        max_w = col_w - 2 * PIC_MARGIN
        max_h = MAX_ROW_HEIGHT - 2 * PIC_MARGIN

        #pass 1: text and row heights (nothing placed yet, so nothng can drift)
        plan: list[tuple[int, Change, dict]] = []
        row = FIRST_ROW
        for c in changes:
            old_png, new_png = shots.get(c.no, (None, None))
            pics: dict[int, tuple[Path, float, float]] = {}
            for col, png in ((COL_OLD, old_png), (COL_NEW, new_png)):
                if png is not None and Path(png).exists():
                    w, h = _fit(Path(png), max_w, max_h)
                    pics[col] = (Path(png).resolve(), w,h)
            
            ws.Cells(row, COL_NO).Value = c.no
            ws.Cells(row, COL_HEADER).Value = c.heading
            ws.Cells(row, COL_DESC).Value = c.kind
            text = ws.Range(ws.Cells(row, COL_NO), ws.Cells(row, COL_DESC))
            text.VerticalAlignment = XL_TOP
            text.WrapText = True
            pic_h = max((p[2] for p in pics.values()), default = 0.0)
            #set the height AFTER wrap text, so Excel cannot change it later
            ws.Rows(row).RowHeight = max (MIN_ROW_HEIGHT, pic_h + 2 * PIC_MARGIN)
            plan.append((row, c, pics))
            row += 1
        
        # pass 2: pictures, now that every row position is final
        misplaced = 0
        for r, c, pics in plan:
            for col, (png, w, h) in pics.items():
                cell = ws.Cells(r, col)
                shp = ws.Shapes.AddPicture(str(png), False, True, cell.Left + PIC_MARGIN, cell.Top + PIC_MARGIN, w, h)
                shp.Placement = XL_MOVE_AND_SIZE
                shp.Name = f"CP{c.no:03d}_{'old' if col == COL_OLD else 'new'}"
                if shp.TopLeftCell.Row != r:
                    misplaced += 1
        
        if misplaced:
            print(f"WARNING: {misplaced} picture(s) not in their own row")
        
        last_row = row - 1
        if last_row > FILTER_LAST_ROW:
            ws.AutoFilterMode = False
            ws.Range(f"{FILTER_RANGE_PREFIX}{last_row}").AutoFilter()
        
        wb.Save()
    
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return out_path



def main() -> None:
    ap = argparse.ArgumentParser(description="Write changes into the CPA template (Step 1e).")
    ap.add_argument("old", help="old .pdf or .docx")
    ap.add_argument("new", help="new .pdf or .docx")
    ap.add_argument("--template", default="templates/CPA_Template_noAutoCount.xlsx",
                    help="CPA template (default: templates/CPA_Template_noAutoCount.xlsx)")
    ap.add_argument("--out", default="output", help="output folder (default: output)")
    ap.add_argument("--name", default=None, help="output file name (default: CPA_<new file>.xlsx)")
    args = ap.parse_args()

    out_dir = Path(args.out)
    old = load(args.old, out_dir)
    new = load(args.new, out_dir)
    changes = compare(old, new)
    write_changes(changes, old, new, out_dir / "changes.txt")
    shots = make_shots(changes, old.pdf, new.pdf, out_dir, old, new)

    name = args.name or f"CPA_{Path(args.new).stem}.xlsx"
    path = write_excel(changes, shots, args.template, out_dir / name)
    print(f"{len(changes)} changes written -> {path}")

if __name__ == "__main__":
    main()
