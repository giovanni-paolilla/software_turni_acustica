"""Widget Tkinter riutilizzabili: ScrollableFrame, styled_btn, card."""
from __future__ import annotations
import tkinter as tk
from tkinter import ttk

from turni.constants import (
    BG, PANEL_BG, ACCENT, BORDER, MUTED, _FONT_FAMILY, FONT_BOLD,
)


def configure_ttk_style() -> None:
    s = ttk.Style()
    s.theme_use("clam")
    s.configure("Vertical.TScrollbar",
                background=BORDER, troughcolor=BG,
                arrowcolor=MUTED, bordercolor=BG)
    s.configure("Horizontal.TScrollbar",
                background=BORDER, troughcolor=BG,
                arrowcolor=MUTED, bordercolor=BG)
    s.configure("TProgressbar",
                troughcolor=BG, background=ACCENT,
                darkcolor=ACCENT, lightcolor=ACCENT)


def _darken(hex_color: str, factor: float = 0.85) -> str:
    """Scurisce un colore esadecimale del (1-factor)*100 %.

    Raises:
        ValueError: se hex_color non e' nel formato #RRGGBB.
    """
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"_darken: colore non valido '{hex_color}' (atteso #RRGGBB)")
    r, g, b = (int(h[i:i+2], 16) for i in (0, 2, 4))
    r, g, b = (max(0, int(c * factor)) for c in (r, g, b))
    return f"#{r:02x}{g:02x}{b:02x}"


def styled_btn(parent: tk.Widget, text: str, command,
               bg: str = ACCENT, fg: str = "white", **kw) -> tk.Button:
    hover = _darken(bg)
    btn = tk.Button(parent, text=text, command=command,
                    bg=bg, fg=fg, font=FONT_BOLD,
                    relief="flat", cursor="hand2",
                    padx=14, pady=6, **kw)
    btn.bind("<Enter>", lambda e: btn.config(bg=hover))
    btn.bind("<Leave>", lambda e: btn.config(bg=bg))
    return btn


def card(parent: tk.Widget, **kw) -> tk.Frame:
    return tk.Frame(parent, bg=PANEL_BG, relief="flat",
                    highlightbackground=BORDER, highlightthickness=1, **kw)


class ScrollableFrame(tk.Frame):
    """Frame con scrollbar verticale; lo scroll e' attivo solo quando il mouse e' dentro."""

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=BG, **kw)
        self._canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical",
                           command=self._canvas.yview)
        self.inner = tk.Frame(self._canvas, bg=BG)
        self.inner.bind("<Configure>",
            lambda e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")))
        self._canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self._canvas.configure(yscrollcommand=sb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.inner.bind("<Enter>", self._bind_scroll)
        self.inner.bind("<Leave>", self._unbind_scroll)

    def _on_scroll(self, e):
        # Linux:   usa e.num (Button-4 / Button-5), e.delta e' sempre 0
        # macOS:   e.delta e' float (es. 3.5), //120 dava 0 -> scroll non funzionava
        # Windows: e.delta e' multiplo di 120, solo il segno conta
        # e.delta==0 su Windows (trackpad in stato intermedio) -> no-op
        if e.num == 4:
            delta = -1
        elif e.num == 5:
            delta = 1
        elif e.delta > 0:
            delta = -1
        elif e.delta < 0:
            delta = 1
        else:
            return
        self._canvas.yview_scroll(delta, "units")

    def _bind_scroll(self, _=None):
        self._canvas.bind_all("<MouseWheel>", self._on_scroll)
        self._canvas.bind_all("<Button-4>",   self._on_scroll)
        self._canvas.bind_all("<Button-5>",   self._on_scroll)

    def _unbind_scroll(self, event=None):
        # Evita unbind se il cursore e' ancora sopra il canvas
        # (accade quando si entra in un widget figlio di self.inner)
        if event is not None:
            w = self._canvas.winfo_containing(event.x_root, event.y_root)
            if w and (w == self._canvas or str(w).startswith(str(self._canvas))):
                return
        self._canvas.unbind_all("<MouseWheel>")
        self._canvas.unbind_all("<Button-4>")
        self._canvas.unbind_all("<Button-5>")
