"""Colours and styles of the app (light and dark), see docs/App_Design.md section 13.

Widgets get their look from an objectName, e.g.  button.setObjectName("primary").
"""
from __future__ import annotations

from PySide6.QtWidgets import QApplication, QLabel, QSizePolicy

FONTS = '"Segoe UI", "Yu Gothic UI", "Meiryo", "Inter", "IPAGothic", sans-serif'

THEMES = {
    "light": dict(bg="#F4F6F9", surface="#FFFFFF", raised="#F9FAFC", border="#E2E7EE",
                  text="#1A2230", muted="#5E6A7A", accent="#1F4E8C", accent_soft="#E6EEF9",
                  selected="#DCE7F7", ok="#1E8A57", ok_soft="#E3F4EB", bad="#C2362B",
                  bad_soft="#FBE7E5", warn="#B7791F", warn_soft="#FBF1DE", pend="#6B7686",
                  pend_soft="#ECEFF3", on_accent="#FFFFFF"),
    "dark": dict(bg="#11151B", surface="#1A2028", raised="#202733", border="#2C3542",
                 text="#E7EBF1", muted="#9AA6B6", accent="#6C9CF0", accent_soft="#1E2D46",
                 selected="#26385A", ok="#4CC38A", ok_soft="#173527", bad="#F0766B",
                 bad_soft="#3D1E1C", warn="#E3AE52", warn_soft="#3A2E17", pend="#9AA6B6",
                 pend_soft="#262E3A", on_accent="#0E1622"),
}

# review status -> (text colour, background colour) token names
STATUS_COLORS = {
    "Approved": ("ok", "ok_soft"),
    "Rejected": ("bad", "bad_soft"),
    "Follow-up": ("warn", "warn_soft"),
    "Pending": ("pend", "pend_soft"),
}

current = "light"

def colors() -> dict:
    return THEMES[current]

def stylesheet(t: dict) -> str:
    return f"""
* {{ font-family: {FONTS}; font-size: 13px; color: {t['text']}; }}
QMainWindow, QWidget#page {{ background: {t['bg']}; }}
QFrame#topbar {{ background: {t['surface']}; border-bottom: 1px solid {t['border']}; }}
QFrame#card {{ background: {t['surface']}; border: 1px solid {t['border']}; border-radius: 10px; }}
QFrame#soft {{ background: {t['raised']}; border: 1px solid {t['border']}; border-radius: 10px; }}
QFrame#drop {{ background: {t['raised']}; border: 2px dashed {t['border']}; border-radius: 10px; }}
QFrame#drop[hover="true"] {{ border-color: {t['accent']}; background: {t['accent_soft']}; }}
QFrame#line {{ background: {t['border']}; max-height: 1px; min-height: 1px; }}
QLabel#h1 {{ font-size: 26px; font-weight: 700; }}
QLabel#h2 {{ font-size: 16px; font-weight: 600; }}
QLabel#muted {{ color: {t['muted']}; }}
QLabel#eyebrow {{ color: {t['muted']}; font-size: 11px; font-weight: 600; letter-spacing: 1px; }}
QLabel#brand {{ font-size: 15px; font-weight: 700; color: {t['accent']}; }}
QPushButton {{ background: {t['surface']}; border: 1px solid {t['border']}; border-radius: 7px;
               padding: 7px 14px; }}
QPushButton:hover {{ border-color: {t['accent']}; }}
QPushButton:disabled {{ color: {t['muted']}; }}
QPushButton#primary {{ background: {t['accent']}; color: {t['on_accent']}; border: none; font-weight: 600;
                       padding: 9px 18px; }}
QPushButton#primary:disabled {{ background: {t['border']}; color: {t['muted']}; }}
QPushButton#ghost {{ background: transparent; border: none; color: {t['muted']}; padding: 6px 10px; }}
QPushButton#ghost:hover {{ color: {t['accent']}; }}
QPushButton#approve {{ background: {t['ok']}; color: {t['on_accent']}; border: none; font-weight: 600;
                       padding: 9px 20px; }}
QPushButton#reject {{ color: {t['bad']}; border: 1px solid {t['bad']}; font-weight: 600; padding: 8px 16px; }}
QPushButton#follow {{ color: {t['warn']}; border: 1px solid {t['warn']}; font-weight: 600; padding: 8px 16px; }}
QPushButton#segment {{ border-radius: 0; padding: 9px 18px; }}
QPushButton#segment:checked {{ background: {t['accent_soft']}; border: 1px solid {t['accent']};
                               color: {t['accent']}; font-weight: 600; }}
QLineEdit, QComboBox, QPlainTextEdit {{ background: {t['surface']}; border: 1px solid {t['border']};
                                       border-radius: 7px; padding: 6px 9px; }}
QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus {{ border-color: {t['accent']}; }}
QComboBox {{ padding-right: 22px; }}
QComboBox::drop-down {{ border: none; width: 20px; }}
QListWidget {{ background: {t['surface']}; border: 1px solid {t['border']}; border-radius: 10px;
               outline: none; }}
QListWidget::item {{ border-bottom: 1px solid {t['border']}; }}
QListWidget::item:selected {{ background: {t['selected']}; }}
QFrame#batch {{ background: {t['accent']}; border-radius: 10px; }}
QFrame#batch QLabel, QFrame#batch QPushButton {{ color: {t['on_accent']}; font-weight: 600; }}
QFrame#batch QPushButton {{ background: transparent; border: 1px solid {t['on_accent']};
                           padding: 4px 7px; font-size: 12px; }}
QProgressBar {{ background: {t['border']}; border: none; border-radius: 3px; max-height: 6px; }}
QProgressBar::chunk {{ background: {t['accent']}; border-radius: 3px; }}
QScrollArea {{ border: none; background: transparent; }}
QToolTip {{ background: {t['surface']}; color: {t['text']}; border: 1px solid {t['border']}; }}
"""

def apply(name: str) -> None:
    """Switch the whole app to the light or dark theme."""
    global current
    current = name if name in THEMES else "light"
    QApplication.instance().setStyleSheet(stylesheet(THEMES[current]))

def chip(text: str, fg: str, bg: str) -> QLabel:
    """Small rounded label, e.g. a status or type. fg/bg are token names ("ok", "ok_soft")."""
    t = colors()
    label = QLabel(text)
    label.setStyleSheet(f"color:{t[fg]}; background:{t[bg]}; border-radius:9px; "
                        f"padding:2px 9px; font-size:11px; font-weight:600;")
    label.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
    return label

def status_chip(status: str) -> QLabel:
    fg, bg = STATUS_COLORS.get(status, STATUS_COLORS["Pending"])
    return chip(status, fg, bg)
