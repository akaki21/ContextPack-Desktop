"""Colors and ttk styles used by the Desktop window."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


COLORS = {
    "background": "#0b1220",
    "card": "#162033",
    "card_alt": "#111a2b",
    "border": "#2b3a55",
    "text": "#f8fafc",
    "muted": "#9aa9bf",
    "accent": "#38bdf8",
    "accent_active": "#0ea5e9",
    "success": "#34d399",
    "warning": "#fbbf24",
    "danger": "#fb7185",
}


def configure_gui_styles(root: tk.Tk) -> None:
    """Configure every named ttk style in one visual-only module."""

    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TFrame", background=COLORS["background"])
    style.configure("Card.TFrame", background=COLORS["card"])
    style.configure("AltCard.TFrame", background=COLORS["card_alt"])
    style.configure("TLabel", background=COLORS["background"], foreground=COLORS["text"], font=("Segoe UI", 10))
    style.configure("Title.TLabel", font=("Segoe UI Semibold", 24), foreground=COLORS["text"])
    style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground=COLORS["muted"])
    style.configure("Section.TLabel", background=COLORS["card"], font=("Segoe UI Semibold", 12), foreground=COLORS["text"])
    style.configure("Card.TLabel", background=COLORS["card"], foreground=COLORS["text"])
    style.configure("Hint.TLabel", background=COLORS["card"], foreground=COLORS["muted"], font=("Segoe UI", 9))
    style.configure("Status.TLabel", background=COLORS["card_alt"], foreground=COLORS["muted"], font=("Segoe UI", 9))
    style.configure("TButton", font=("Segoe UI Semibold", 10), padding=(12, 8))
    style.configure("Accent.TButton", background=COLORS["accent"], foreground="#062033", borderwidth=0, padding=(18, 10))
    style.map("Accent.TButton", background=[("active", COLORS["accent_active"]), ("disabled", COLORS["border"])])
    style.configure("Secondary.TButton", background=COLORS["border"], foreground=COLORS["text"], borderwidth=0)
    style.map("Secondary.TButton", background=[("active", "#3b4d6b")])
    style.configure("Danger.TButton", background=COLORS["danger"], foreground="#3f0a18", borderwidth=0)
    style.configure("TCombobox", fieldbackground=COLORS["card_alt"], background=COLORS["border"], foreground=COLORS["text"], arrowcolor=COLORS["text"], padding=7)
    style.map("TCombobox", fieldbackground=[("readonly", COLORS["card_alt"]), ("disabled", COLORS["card"])], foreground=[("readonly", COLORS["text"]), ("disabled", COLORS["muted"])])
    style.configure("TSpinbox", fieldbackground=COLORS["card_alt"], foreground=COLORS["text"], arrowcolor=COLORS["text"], padding=7)
    style.configure("Horizontal.TProgressbar", troughcolor=COLORS["card_alt"], background=COLORS["accent"], bordercolor=COLORS["card_alt"], lightcolor=COLORS["accent"], darkcolor=COLORS["accent"])

