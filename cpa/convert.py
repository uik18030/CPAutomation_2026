"""Turn an input file into a PDF. PDFs are used as-is; .docx/.doc go through Word.

Fields (cross-references, dates, TOC, ...) are LOCKED before export so the PDF
shows exactly the text saved in the document - Word must not refresh them
(refreshing can produce "Error! Reference source not found.").
The source document is opened read-only and never saved.
"""
from __future__ import annotations

from pathlib import Path

_WD_EXPORT_FORMAT_PDF = 17
_WD_DO_NOT_SAVE_CHANGES = 0
_WD_ALERTS_NONE = 0

def to_pdf(src: str | Path, out_dir: str | Path, prefix: str = "") -> Path:
    """Return the path of a PDF version of *src*.

    - .pdf          -> returned unchanged
    - .docx / .doc  -> exported with Microsoft Word into *out_dir*
    """
    src = Path(src).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not src.exists():
        raise FileNotFoundError(f"Input file not found: {src}")

    suffix = src.suffix.lower()
    if suffix == ".pdf":
        return src

    if suffix in (".docx", ".doc"):
        dst = out_dir / f"{prefix}{src.stem}.pdf"
        _word_to_pdf(src, dst)
        if not dst.exists():
            raise RuntimeError(f"Word did not produce a PDF for: {src}")
        return dst

    raise ValueError(f"Unsupported file type '{src.suffix}' (use .pdf or .docx): {src}")

def _lock_all_fields(doc) -> int:
    """Lock fields in every story (body, headers/footers, text boxes, notes)."""
    locked = 0
    for story in doc.StoryRanges:
        rng = story
        while rng is not None:
            try:
                count = rng.Fields.Count
                if count:
                    rng.Fields.Locked = True
                    locked += count
            except Exception:
                pass  # some story types have no Fields collection
            try:
                rng = rng.NextStoryRange
            except Exception:
                rng = None
    return locked

def _word_to_pdf(src: Path, dst: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    doc = None
    old_update_at_print = None
    try:
        # Separate, invisible Word instance: does not touch the user's open windows.
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = _WD_ALERTS_NONE

        try:
            old_update_at_print = word.Options.UpdateFieldsAtPrint
            word.Options.UpdateFieldsAtPrint = False
        except Exception:
            old_update_at_print = None

        # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        doc = word.Documents.Open(str(src), False, True, False)
        _lock_all_fields(doc)
        doc.ExportAsFixedFormat(str(dst), _WD_EXPORT_FORMAT_PDF)
    finally:
        if doc is not None:
            try:
                doc.Close(_WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        if word is not None:
            if old_update_at_print is not None:
                try:
                    word.Options.UpdateFieldsAtPrint = old_update_at_print
                except Exception:
                    pass
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

# Part to handle password 
_PASSWORDS: dict[str, str] = {}

def _key(pdf_path) -> str:
    return str(Path(pdf_path).resolve()).lower()

def get_password(pdf_path) -> str | None:
    """Password entered for this PDF in this run (never stored on disk)."""
    return _PASSWORDS.get(_key(pdf_path))

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

def ensure_password(pdf_path) -> None:
    """If the PDF needs a password to open, ask for it (input is hidden)."""
    import getpass
    import pypdfium2 as pdfium

    key = _key(pdf_path)
    if key in _PASSWORDS:
        return 
    try:   
        pdfium.PdfDocument(str(pdf_path)).close()
        return 
    except pdfium.PdfiumError:
        pass
    for _ in range(3):
        pw = getpass.getpass(f"Password for {Path(pdf_path).name}: ")
        try:
            pdfium.PdfDocument(str(pdf_path), password=pw).close()
            _PASSWORDS[key] = pw
            return
        except pdfium.PdfiumError:
            print("Wrong password, try again.")
    raise PermissionError(f"Could not open (wrong password): {pdf_path}")
