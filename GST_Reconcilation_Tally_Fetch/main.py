"""
GST Reconciliation — Landing Page
Four arrow-card chooser. Each card will launch its own reconciliation module.
"""

from __future__ import annotations

import tkinter as tk
import customtkinter as ctk


CARDS = [
    ("Sale Vs GSTR-1\nReconciliation",       "#f59e0b"),
    ("Purchase Vs. GSTR-2A\nReconciliation",  "#10b981"),
    ("Purchase Vs. GSTR-2B\nReconciliation",  "#e11d48"),
    ("2B Vs. 2A\nReconciliation",             "#7c3aed"),
]


def _make_reco_card(parent: tk.Widget, label: str, color: str,
                    area_bg: str, dark: bool) -> tk.Canvas:
    W, H = 490, 165
    arr  = 40
    cut  = 22
    bw   = 3
    m    = bw + 2

    card_bg  = "#1a2535" if dark else "#ffffff"
    text_col = "#e2e8f0" if dark else "#1e293b"
    sep_col  = "#334155" if dark else "#e5e7eb"
    dot_col  = "#4b5563" if dark else "#cbd5e1"

    cv = tk.Canvas(parent, width=W, height=H,
                   bg=area_bg, highlightthickness=0, bd=0, cursor="hand2")

    # Outer polygon — accent border
    cv.create_polygon(
        bw+cut, bw,   W-arr-bw, bw,   W-bw, H//2,
        W-arr-bw, H-bw,   bw+cut, H-bw,   bw, H-cut-bw,   bw, cut+bw,
        fill=color, outline="", smooth=False)

    # Inner polygon — card background
    cv.create_polygon(
        m+cut, m,   W-arr-m, m,   W-m-bw, H//2,
        W-arr-m, H-m,   m+cut, H-m,   m, H-cut-m,   m, cut+m,
        fill=card_bg, outline="", smooth=False)

    # Dashed circle (icon ring)
    cx, cy, r = 88, H//2, 46
    segs = 14
    for i in range(segs):
        start = (360 / segs) * i + 90
        cv.create_arc(cx-r, cy-r, cx+r, cy+r,
                      start=start, extent=(360/segs)*0.55,
                      style="arc", outline=color, width=2)

    cv.create_text(cx, cy, text="⚖", font=("Segoe UI Emoji", 26), fill=color)

    # Vertical separator
    sx = cx + r + 20
    cv.create_line(sx, 18, sx, H-18, fill=sep_col, width=1)

    # Title
    tx = sx + 20
    parts = label.split("\n")
    cv.create_text(tx, H//2 - (15 if len(parts) > 1 else 0),
                   text=parts[0], font=("Segoe UI", 14, "bold"),
                   fill=text_col, anchor="w")
    if len(parts) > 1:
        cv.create_text(tx, H//2 + 15, text=parts[1],
                       font=("Segoe UI", 14, "bold"),
                       fill=text_col, anchor="w")

    # Decorative footer dots + bar
    dy, dx = H - 22, sx + 14
    for _ in range(3):
        cv.create_oval(dx-4, dy-4, dx+4, dy+4, fill=dot_col, outline="")
        dx += 14
    cv.create_rectangle(dx+5, dy-4, dx+28, dy+4, fill=color, outline="")

    return cv


class RecoLandingApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("GST Reconciliation")
        self.geometry("1120x560")
        self.resizable(True, True)
        self._build_ui()

    def _build_ui(self):
        dark    = ctk.get_appearance_mode().lower() == "dark"
        area_bg = "#111827" if dark else "#f8fafc"

        outer = ctk.CTkFrame(self, fg_color=area_bg, corner_radius=0)
        outer.pack(fill="both", expand=True)
        outer.grid_rowconfigure(0, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        holder = ctk.CTkFrame(outer, fg_color="transparent")
        holder.place(relx=0.5, rely=0.5, anchor="center")

        for i, (lbl, color) in enumerate(CARDS):
            r, c = divmod(i, 2)
            cv = _make_reco_card(holder, lbl, color, area_bg, dark)
            cv.grid(row=r, column=c, padx=18, pady=18)

            def _click(_, label=lbl):
                self._on_card_click(label)
            cv.bind("<Button-1>", _click)

    def _on_card_click(self, label: str):
        import tkinter.messagebox as _mb
        _mb.showinfo(
            "Coming Soon",
            f"'{label.replace(chr(10), ' ')}'\n\nThis reconciliation module will be available soon.",
        )


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    app = RecoLandingApp()
    app.mainloop()
