"""Widget CustomTkinter riutilizzabili: ScrollableFrame, styled_btn, card."""
from __future__ import annotations
import customtkinter as ctk
from tkinter import ttk

from turni.constants import (
    BG, PANEL_BG, ACCENT, BORDER, MUTED, FONT_BOLD,
)


def configure_ttk_style() -> None:
    """Configura aspetto globale CTk e stili ttk residui (Treeview)."""
    ctk.set_appearance_mode("light")
    s = ttk.Style()
    s.theme_use("clam")
    s.configure("Treeview",
                background=PANEL_BG, foreground="#1E293B",
                rowheight=28, fieldbackground=PANEL_BG)
    s.configure("Treeview.Heading",
                background="#E2E8F0", foreground="#1E293B",
                font=("", 10, "bold"))
    s.map("Treeview", background=[("selected", ACCENT)])


def styled_btn(parent, text: str, command,
               bg: str = ACCENT, fg: str = "white", **kw) -> ctk.CTkButton:
    """Pulsante con hover automatico e angoli arrotondati."""
    for k in ("cursor", "padx", "pady", "relief"):
        kw.pop(k, None)
    return ctk.CTkButton(
        parent, text=text, command=command,
        fg_color=bg, text_color=fg,
        font=FONT_BOLD,
        corner_radius=6, height=34,
        **kw,
    )


def card(parent, **kw) -> ctk.CTkFrame:
    """Card con bordi arrotondati e ombra leggera."""
    return ctk.CTkFrame(
        parent, fg_color=PANEL_BG,
        corner_radius=8,
        border_width=1, border_color=BORDER,
        **kw,
    )


class ScrollableFrame(ctk.CTkScrollableFrame):
    """Frame scrollabile — wrapper CTkScrollableFrame con attributo inner."""

    def __init__(self, parent, **kw):
        kw.setdefault("fg_color", BG)
        super().__init__(parent, **kw)
        self.inner = self
