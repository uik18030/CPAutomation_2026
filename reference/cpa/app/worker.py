"""Long jobs (Word conversion, comparing, screenshots, Excel) run in a background thread,
so the window stays responsive. Each job reports progress text and returns a result."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Signal

from cpa.clean import CleanDoc, load
from cpa.compare import Change, compare, write_changes
from cpa.shots import make_shots

@dataclass
class CompareResult:
    old: CleanDoc
    new: CleanDoc
    changes: list[Change]
    shots: dict                     # change no -> (old png | None, new png | None)
    out_dir: Path

def run_compare(report, old_path: str, new_path: str, out_dir: str) -> CompareResult:
    out = Path(out_dir)
    report("Reading the old file…")
    old = load(old_path, out)
    report("Reading the new file…")
    new = load(new_path, out)
    report("Comparing…")
    changes = compare(old, new)
    write_changes(changes, old, new, out / "changes.txt")
    report(f"Making screenshots for {len(changes)} changes…")
    shots = make_shots(changes, old.pdf, new.pdf, out, old, new)
    return CompareResult(old, new, changes, shots, out)

def run_export(report, changes: list[Change], shots: dict, template: str, xlsx: str) -> Path:
    from cpa.excel import write_excel             # needs Excel (Windows)
    report(f"Writing {len(changes)} rows to Excel…")
    return write_excel(changes, shots, template, xlsx)

class _Job(QObject):
    progress = Signal(str)
    done = Signal(object)
    failed = Signal(str)

    def __init__(self, fn, args):
        super().__init__()
        self.fn, self.args = fn, args

    def run(self) -> None:
        try:
            self.done.emit(self.fn(self.progress.emit, *self.args))
        except Exception as e:                     # shown to the user in a message box
            self.failed.emit(f"{type(e).__name__}: {e}")

def start(owner: QObject, fn, *args, on_done, on_failed, on_progress) -> None:
    """Run fn(report, *args) in a background thread; the callbacks run in the window's thread."""
    thread = QThread(owner)
    job = _Job(fn, args)
    job.moveToThread(thread)
    thread.started.connect(job.run)
    job.progress.connect(on_progress)
    job.done.connect(on_done)
    job.failed.connect(on_failed)
    for sig in (job.done, job.failed):
        sig.connect(thread.quit)
    thread.finished.connect(job.deleteLater)
    thread.finished.connect(thread.deleteLater)
    owner._jobs = getattr(owner, "_jobs", []) + [(thread, job)]   # keep them alive while running
    thread.start()
