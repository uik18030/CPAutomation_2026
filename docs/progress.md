# CPA Tool – Progress Log

## Status
- Requirements agreed → see Project_details.md (or docs/requirements.md)
- Rule: NO tool code until the user says "start Step 1".

## Step 0 – Project setup
- [x] GitHub repo created and cloned into VS Code (CPAutomation_2026)
- [x] Folder structure: docs/, cpa/__init__.py, templates/, tests/data/, tests/expected/, output/
- [x] .gitignore: .venv/, pyvenv.cfg, *.pdf, *.docx, output/, data/, __pycache__/, *.pyc, build/, dist/
- [x] tests/expected/answer_key.json (6 expected changes)
- [x] Python: `python` not on PATH, so use `py`. Only 3.15.2 was installed (too new)
      → installed Python 3.12 (Microsoft Store)
- [x] Virtual environment: `py -3.12 -m venv .venv`, activate with `.venv\Scripts\Activate.ps1`
- [x] Network: pip connects directly (company PAC file routes pypi.org DIRECT).
      No proxy setting needed. Slow connection → use `--timeout 120 --retries 10`
- [x] requirements.txt: pdfplumber, pypdfium2, Pillow, imagehash, openpyxl, docx2pdf, PySide6-Essentials
      (PySide6-Essentials instead of PySide6: much smaller download, smaller .exe)
- [x] Installed: pdfplumber 0.11.10, pypdfium2 5.15.2, Pillow 13.0.0, imagehash 4.3.2,
      openpyxl 3.1.5, docx2pdf 0.1.8 (+ numpy, scipy, pdfminer.six, pywin32 …)
- [x] First commits pushed ("Project setup", "Use PySide6-Essentials")
- [x] PySide6-Essentials installed
- [x] PDF conversion with docx2pdf (output/old.pdf, output/new.pdf)
- [ ] PDF check results:
      - new.pdf: Trank image upside down?  → ?
      - old.pdf: 「リヤシートリマインダ表示は以下条件に従い制御すること。」 struck through? → ?
- [ ] Optional: pin versions with `python -m pip freeze > requirements.txt`

## Next
- Step 1: proof of concept (convert, extract, compare, red-box screenshots, write CPA_JPN)
  → check against tests/expected/answer_key.json (6 changes)

## When resuming with the assistant
- Upload Project_details.md (requirements) + docs/progress.md
- Say "start Step 1"
