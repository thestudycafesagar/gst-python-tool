"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         GST & INCOME TAX AUTOMATION SUITE  —  Unified Launcher               ║
║                                                                              ║
║  Landing page → select category → tabbed tool view → ← Home to go back.      ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import sys
import os
import importlib.util
import traceback
import customtkinter as ctk
from datetime import datetime

# ── Appearance (must run before any tool import) ─────────────────────────────
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

_RealCTk           = ctk.CTk
_RealSetAppearance = ctk.set_appearance_mode

# ── Base paths ────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    _BASE = sys._MEIPASS
else:
    _BASE = os.path.dirname(os.path.abspath(__file__))

_GST_BASE = os.path.join(_BASE, "GST")
_IT_BASE  = os.path.join(_BASE, "Income Tax")


# ══════════════════════════════════════════════════════════════════════════════
#  _EmbeddedFrame  – swaps in for ctk.CTk during tool import
# ══════════════════════════════════════════════════════════════════════════════
class _EmbeddedFrame(ctk.CTkFrame):
    _host: ctk.CTkFrame = None

    def __init__(self, *args, **kwargs):
        kwargs.pop("fg_color", None)
        super().__init__(_EmbeddedFrame._host)
        self.grid(row=0, column=0, sticky="nsew")

    def title(self, t=None):       return ""
    def geometry(self, g=None):    return ""
    def resizable(self, *a, **k):  pass
    def mainloop(self):            pass
    def lift(self):                pass
    def iconbitmap(self, *a, **k): pass
    def wm_title(self, t=None):    return ""
    def protocol(self, *a, **k):   pass
    def attributes(self, *a, **k): pass


# ══════════════════════════════════════════════════════════════════════════════
#  DESIGN TOKENS
# ══════════════════════════════════════════════════════════════════════════════
_C = {
    # Page / surface
    "banner_bg":   ("#1e293b", "#060c18"),
    "status_bg":   ("#1e293b", "#060c18"),
    "surface":     ("#ffffff", "#1e293b"),
    "surface2":    ("#f8fafc", "#111827"),
    "border":      ("#e2e8f0", "#334155"),
    # Text
    "text_hi":     ("#0f172a", "#f1f5f9"),
    "text_mid":    ("#475569", "#94a3b8"),
    "text_lo":     ("#94a3b8", "#475569"),
    # Brand
    "primary":     ("#4f46e5", "#6366f1"),
    "primary_hov": ("#4338ca", "#4f46e5"),
    # GST category
    "gst_acc":     ("#4f46e5", "#818cf8"),
    "gst_bg":      ("#ede9fe", "#12103a"),
    "gst_hover":   ("#ddd6fe", "#1c1856"),
    # Income Tax category
    "it_acc":      ("#059669", "#34d399"),
    "it_bg":       ("#d1fae5", "#062818"),
    "it_hover":    ("#a7f3d0", "#0a3d22"),
}

# Per-tool accent colours (light, dark)
_GST_ACCENTS = [
    ("#4f46e5", "#818cf8"),   # GSTR-2B      Indigo
    ("#7c3aed", "#a78bfa"),   # GSTR-3B      Violet
    ("#0891b2", "#22d3ee"),   # 3B→Excel     Cyan
    ("#059669", "#34d399"),   # GST Verifier Emerald
    ("#d97706", "#fbbf24"),   # Challan      Amber
    ("#0284c7", "#38bdf8"),   # R1 JSON      Sky
    ("#db2777", "#f472b6"),   # JSON→Excel   Pink
    ("#dc2626", "#f87171"),   # R1 PDF       Red
]
_IT_ACCENTS = [
    ("#0891b2", "#22d3ee"),   # 26AS         Cyan
    ("#059669", "#34d399"),   # IT Challan   Emerald
    ("#7c3aed", "#a78bfa"),   # ITR Bot      Violet
]


# ══════════════════════════════════════════════════════════════════════════════
#  TOOL REGISTRY
# ══════════════════════════════════════════════════════════════════════════════
GST_TOOLS = [
    {"tab": "📥  GSTR-2B",      "module": os.path.join(_GST_BASE, "GST 2B Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-2B returns via Selenium automation."},
    {"tab": "📥  GSTR-3B",      "module": os.path.join(_GST_BASE, "GST 3B Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-3B returns via Selenium automation."},
    {"tab": "📊  3B → Excel",   "module": os.path.join(_GST_BASE, "GST 3B to Excel",       "main.py"),        "class": "GSTR3BConverterPro", "desc": "Convert GSTR-3B PDF files to formatted Excel sheets."},
    {"tab": "🤖  GST Verifier", "module": os.path.join(_GST_BASE, "GST Bot",               "gst_pro_app.py"), "class": "GSTApp",             "desc": "Verify bulk GSTINs and extract filing history & details."},
    {"tab": "💰  Challan",      "module": os.path.join(_GST_BASE, "GST Challan Downloader","main.py"),        "class": "App",                "desc": "Download GST Challan PDFs in bulk (Monthly / Quarterly)."},
    {"tab": "📑  R1 JSON",      "module": os.path.join(_GST_BASE, "GST R1 Downloader",     "mai.py"),         "class": "App",                "desc": "Request or download GSTR-1 JSON files for multiple users."},
    {"tab": "📊  JSON → Excel", "module": os.path.join(_GST_BASE, "JSON to Excel",          "main.py"),        "class": "App",                "desc": "Convert GSTR-1 JSON exports to multi-sheet Excel reports."},
    {"tab": "🖨️  R1 PDF",       "module": os.path.join(_GST_BASE, "R1 PDF Downloader",     "main.py"),        "class": "App",                "desc": "Bulk download GSTR-1 PDF filed returns from the GST portal."},
]

IT_TOOLS = [
    {"tab": "📄  26 AS",       "module": os.path.join(_IT_BASE, "26 AS Downlaoder",  "main.py"),         "class": "App", "desc": "Download 26AS / AIS / TIS and filed return reports in bulk."},
    {"tab": "💰  IT Challan",  "module": os.path.join(_IT_BASE, "Challan Downloader","main.py"),         "class": "App", "desc": "Download Income Tax Challan PDFs in bulk."},
    {"tab": "🤖  ITR Bot",     "module": os.path.join(_IT_BASE, "ITR - Bot",         "GUI_based_app.py"),"class": "App", "desc": "Automate ITR filing workflows with the ITR bot."},
]


# ══════════════════════════════════════════════════════════════════════════════
#  TOOL LOADER
# ══════════════════════════════════════════════════════════════════════════════
def _load_tool(tab_frame: ctk.CTkFrame, module_path: str, class_name: str):
    tab_frame.grid_rowconfigure(0, weight=1)
    tab_frame.grid_columnconfigure(0, weight=1)

    _EmbeddedFrame._host = tab_frame
    ctk.CTk = _EmbeddedFrame
    _saved_mode = ctk.get_appearance_mode()
    ctk.set_appearance_mode = lambda *a, **k: None

    tool_dir = os.path.dirname(os.path.abspath(module_path))
    try:
        if tool_dir not in sys.path:
            sys.path.insert(0, tool_dir)
        uid  = f"_suite_{class_name}_{os.path.basename(tool_dir)}"
        spec = importlib.util.spec_from_file_location(uid, module_path)
        mod  = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        AppClass = getattr(mod, class_name)
        return AppClass()

    except Exception:
        err = traceback.format_exc()
        ef = ctk.CTkFrame(tab_frame, fg_color="transparent")
        ef.grid(row=0, column=0, sticky="nsew", padx=24, pady=24)

        hdr = ctk.CTkFrame(ef, fg_color=("#fef2f2", "#1c0a0a"), corner_radius=10,
                           border_color=("#fca5a5", "#7f1d1d"), border_width=1)
        hdr.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(hdr, text="  ⚠️   Failed to load tool — traceback below",
                     font=("Segoe UI", 13, "bold"),
                     text_color=("#ef4444", "#f87171")).pack(anchor="w", padx=18, pady=12)

        tb = ctk.CTkTextbox(ef, font=("Consolas", 11),
                            text_color=("#ef4444", "#f87171"),
                            fg_color=("#1a1a2e", "#0d0d1a"),
                            corner_radius=10, border_width=1,
                            border_color=("#7f1d1d", "#450a0a"), height=320)
        tb.pack(fill="both", expand=True)
        tb.insert("1.0", err)
        tb.configure(state="disabled")
        return None

    finally:
        ctk.CTk = _RealCTk
        ctk.set_appearance_mode = _RealSetAppearance
        _RealSetAppearance(_saved_mode)


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════
class GSTSuite(_RealCTk):

    def __init__(self):
        super().__init__()
        self.title("GST & Income Tax Automation Suite")
        self.geometry("1200x920")
        self.minsize(980, 760)
        self.resizable(True, True)

        self._current_theme  = "Dark"
        self._theme_btn      = None
        self._active_cat     = None       # "gst" | "it" | None
        self._active_poll    = None       # matches _active_cat when polling
        self._loaded         = {}         # tab_name -> bool
        self._instances      = {}         # tab_name -> tool instance
        self._tabviews       = {}         # cat_key  -> CTkTabview
        self._last_tabs      = {}         # cat_key  -> last active tab
        self._cat_frames     = {}         # cat_key  -> CTkFrame

        self._build_header()
        self._content = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self._content.pack(fill="both", expand=True)
        self._build_statusbar()

        self._landing = self._build_landing()
        self._show_landing()


    # ══════════════════════════════════════════════════════════════════════════
    #  HEADER
    # ══════════════════════════════════════════════════════════════════════════
    def _build_header(self):
        bar = ctk.CTkFrame(self, fg_color=_C["banner_bg"], corner_radius=0, height=68)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)

        # Accent stripe
        ctk.CTkFrame(bar, height=3, corner_radius=0,
                     fg_color=_C["primary"]).pack(fill="x", side="top")

        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=20)

        # Left area — rebuilt on each navigation
        self._hdr_left = ctk.CTkFrame(inner, fg_color="transparent")
        self._hdr_left.pack(side="left", fill="y")

        # Right area — theme toggle + clock
        h_right = ctk.CTkFrame(inner, fg_color="transparent")
        h_right.pack(side="right", fill="y")

        try:
            self._theme_btn = ctk.CTkSegmentedButton(
                h_right,
                values=["🌙  Dark", "☀️  Light"],
                command=self._set_theme,
                width=170, height=32,
                font=("Segoe UI", 11, "bold"),
                selected_color=_C["primary"],
                selected_hover_color=_C["primary_hov"],
                unselected_color=("#334155", "#1e293b"),
                unselected_hover_color=("#475569", "#2d3a4a"),
                text_color=("#f1f5f9", "#f1f5f9"),
            )
        except TypeError:
            self._theme_btn = ctk.CTkSegmentedButton(
                h_right, values=["Dark", "Light"],
                command=self._set_theme, width=160, height=32)

        self._theme_btn.set("🌙  Dark")
        self._theme_btn.pack(side="right", padx=(10, 0))

        self._clock_lbl = ctk.CTkLabel(
            h_right, text="",
            font=("Segoe UI Mono", 11),
            text_color=_C["text_mid"])
        self._clock_lbl.pack(side="right", padx=(0, 18))
        self._tick_clock()

    def _tick_clock(self):
        self._clock_lbl.configure(
            text=datetime.now().strftime("%d %b %Y   %H:%M:%S"))
        self.after(1000, self._tick_clock)

    def _set_theme(self, label: str):
        mode = "Light" if "Light" in label else "Dark"
        self._current_theme = mode
        _RealSetAppearance(mode)
        if self._theme_btn:
            self._theme_btn.set("☀️  Light" if mode == "Light" else "🌙  Dark")

    def _refresh_header_left(self, mode: str, cat_key: str = None):
        """Wipe and rebuild the left side of the header."""
        for w in self._hdr_left.winfo_children():
            w.destroy()

        if mode == "landing":
            pill = ctk.CTkFrame(self._hdr_left, fg_color=_C["primary"],
                                corner_radius=8, width=40, height=40)
            pill.pack(side="left", padx=(0, 14))
            pill.pack_propagate(False)
            ctk.CTkLabel(pill, text="G", font=("Segoe UI", 20, "bold"),
                         text_color="#ffffff").place(relx=0.5, rely=0.5, anchor="center")

            t = ctk.CTkFrame(self._hdr_left, fg_color="transparent")
            t.pack(side="left")
            ctk.CTkLabel(t, text="GST & Income Tax Automation",
                         font=("Segoe UI", 16, "bold"),
                         text_color=("#f1f5f9", "#f1f5f9")).pack(anchor="w")
            ctk.CTkLabel(t, text="Unified Suite  ·  All Tools  ·  One Window",
                         font=("Segoe UI", 10),
                         text_color=_C["text_mid"]).pack(anchor="w")

        else:
            # Back / Home button
            ctk.CTkButton(
                self._hdr_left,
                text="⬅  Home",
                font=("Segoe UI", 12, "bold"),
                width=108, height=36,
                fg_color=("#334155", "#1e293b"),
                hover_color=("#475569", "#2d3a4a"),
                text_color=("#f1f5f9", "#f1f5f9"),
                corner_radius=8,
                command=self._go_home,
            ).pack(side="left", padx=(0, 20))

            acc   = _C["gst_acc"]  if cat_key == "gst" else _C["it_acc"]
            icon  = "🏛"           if cat_key == "gst" else "💼"
            label = "GST Tools"    if cat_key == "gst" else "Income Tax Automation Suite"
            tools = GST_TOOLS      if cat_key == "gst" else IT_TOOLS

            pill = ctk.CTkFrame(self._hdr_left, fg_color=acc,
                                corner_radius=8, width=38, height=38)
            pill.pack(side="left", padx=(0, 12))
            pill.pack_propagate(False)
            ctk.CTkLabel(pill, text=icon, font=("Segoe UI Emoji", 18)
                         ).place(relx=0.5, rely=0.5, anchor="center")

            t = ctk.CTkFrame(self._hdr_left, fg_color="transparent")
            t.pack(side="left")
            ctk.CTkLabel(t, text=label,
                         font=("Segoe UI", 16, "bold"),
                         text_color=acc).pack(anchor="w")
            ctk.CTkLabel(t, text=f"{len(tools)} tools available",
                         font=("Segoe UI", 10),
                         text_color=_C["text_mid"]).pack(anchor="w")


    # ══════════════════════════════════════════════════════════════════════════
    #  STATUS BAR
    # ══════════════════════════════════════════════════════════════════════════
    def _build_statusbar(self):
        bar = ctk.CTkFrame(self, fg_color=_C["status_bg"],
                           corner_radius=0, height=28)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        ctk.CTkFrame(bar, height=1, corner_radius=0,
                     fg_color=_C["border"]).pack(fill="x", side="top")

        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=20)

        ctk.CTkLabel(inner, text="●  Ready",
                     font=("Segoe UI", 9, "bold"),
                     text_color=("#10b981", "#10b981")).pack(side="left")
        ctk.CTkLabel(inner,
                     text=f"  ·  {len(GST_TOOLS)} GST tools  ·  {len(IT_TOOLS)} Income Tax tools",
                     font=("Segoe UI", 9),
                     text_color=_C["text_mid"]).pack(side="left")
        ctk.CTkLabel(inner, text="Automation Suite  v2.0",
                     font=("Segoe UI", 9),
                     text_color=_C["text_lo"]).pack(side="right")


    # ══════════════════════════════════════════════════════════════════════════
    #  LANDING PAGE
    # ══════════════════════════════════════════════════════════════════════════
    def _build_landing(self) -> ctk.CTkFrame:
        page = ctk.CTkFrame(self._content, fg_color="transparent")

        # Everything is placed in a centered wrapper
        wrapper = ctk.CTkFrame(page, fg_color="transparent")
        wrapper.place(relx=0.5, rely=0.5, anchor="center")

        # ── Welcome text ──────────────────────────────────────────────────────
        ctk.CTkLabel(
            wrapper,
            text="Welcome — Select a Category",
            font=("Segoe UI", 26, "bold"),
            text_color=_C["text_hi"],
        ).pack(pady=(0, 8))

        ctk.CTkLabel(
            wrapper,
            text="All automation tools, organized by category. "
                 "Click a card to get started.",
            font=("Segoe UI", 13),
            text_color=_C["text_mid"],
        ).pack(pady=(0, 48))

        # ── Cards row ─────────────────────────────────────────────────────────
        cards_row = ctk.CTkFrame(wrapper, fg_color="transparent")
        cards_row.pack()

        gst_card = self._make_category_card(
            parent=cards_row,
            icon="🏛",
            label="GST Tools",
            sub_text=f"{len(GST_TOOLS)} modules",
            desc="Downloads, converters & verifiers\nfor GST portal automation.",
            acc=_C["gst_acc"],
            normal_fg=_C["gst_bg"],
            hover_fg=_C["gst_hover"],
            callback=lambda: self._show_category("gst"),
        )
        gst_card.pack(side="left", padx=22)

        it_card = self._make_category_card(
            parent=cards_row,
            icon="💼",
            label="Income Tax Automation Suite",
            sub_text=f"{len(IT_TOOLS)} modules",
            desc="26AS, Challan & ITR filing\nautomation tools.",
            acc=_C["it_acc"],
            normal_fg=_C["it_bg"],
            hover_fg=_C["it_hover"],
            callback=lambda: self._show_category("it"),
        )
        it_card.pack(side="left", padx=22)

        return page

    def _make_category_card(self, parent, icon, label, sub_text, desc,
                             acc, normal_fg, hover_fg, callback):
        """Build a large clickable category card."""
        card = ctk.CTkFrame(
            parent,
            fg_color=normal_fg,
            corner_radius=20,
            border_width=2,
            border_color=acc,
            width=310,
            height=340,
        )
        card.pack_propagate(False)

        # Top accent stripe
        strip = ctk.CTkFrame(card, height=6, corner_radius=0, fg_color=acc)
        strip.pack(fill="x")
        strip.pack_propagate(False)

        # Large icon
        ctk.CTkLabel(card, text=icon,
                     font=("Segoe UI Emoji", 58)).pack(pady=(28, 4))

        # Category name
        ctk.CTkLabel(card, text=label,
                     font=("Segoe UI", 22, "bold"),
                     text_color=acc).pack()

        # Module count badge
        badge = ctk.CTkFrame(card, fg_color=acc, corner_radius=20)
        badge.pack(pady=(8, 0))
        ctk.CTkLabel(badge, text=f"  {sub_text}  ",
                     font=("Segoe UI", 10, "bold"),
                     text_color="#ffffff").pack(padx=6, pady=4)

        # Description text
        ctk.CTkLabel(card, text=desc,
                     font=("Segoe UI", 13),
                     text_color=("#1e293b", "#e2e8f0"),
                     justify="center").pack(pady=(16, 6))

        # CTA hint
        ctk.CTkLabel(card, text="Click to explore  →",
                     font=("Segoe UI", 12, "bold"),
                     text_color=acc).pack(pady=(6, 0))

        # Bind hover + click to the entire card (including all children)
        self._bind_card(card, callback, normal_fg, hover_fg)
        return card

    def _bind_card(self, card, callback, normal_fg, hover_fg):
        """Recursively bind hover & click to card and every descendant."""
        def enter(_=None):
            try:
                card.configure(fg_color=hover_fg)
            except Exception:
                pass

        def leave(_=None):
            # Only un-hover if pointer has actually left the card bounds
            try:
                px, py = card.winfo_pointerxy()
                cx = card.winfo_rootx()
                cy = card.winfo_rooty()
                cw = card.winfo_width()
                ch = card.winfo_height()
                if not (cx <= px <= cx + cw and cy <= py <= cy + ch):
                    card.configure(fg_color=normal_fg)
            except Exception:
                try:
                    card.configure(fg_color=normal_fg)
                except Exception:
                    pass

        def click(_=None):
            callback()

        def attach(w):
            try:
                w.configure(cursor="hand2")
            except Exception:
                pass
            w.bind("<Enter>",    lambda _: enter(), add="+")
            w.bind("<Leave>",    lambda _: leave(), add="+")
            w.bind("<Button-1>", lambda _: click(), add="+")
            for child in w.winfo_children():
                attach(child)

        attach(card)


    # ══════════════════════════════════════════════════════════════════════════
    #  CATEGORY PAGE  (built lazily on first visit)
    # ══════════════════════════════════════════════════════════════════════════
    def _get_or_build_category(self, key: str) -> ctk.CTkFrame:
        if key in self._cat_frames:
            return self._cat_frames[key]

        tools   = GST_TOOLS   if key == "gst" else IT_TOOLS
        accents = _GST_ACCENTS if key == "gst" else _IT_ACCENTS

        frame = ctk.CTkFrame(self._content, fg_color="transparent",
                             corner_radius=0)

        # Build tabview
        try:
            tv = ctk.CTkTabview(
                frame, anchor="nw",
                segmented_button_selected_color=_C["primary"],
                segmented_button_selected_hover_color=_C["primary_hov"],
                segmented_button_unselected_hover_color=("#334155", "#1e293b"),
                text_color=_C["text_hi"],
            )
        except (TypeError, ValueError):
            tv = ctk.CTkTabview(frame, anchor="nw")

        # Increase tab label font via the internal segmented button
        try:
            tv._segmented_button.configure(font=("Segoe UI", 13, "bold"))
        except Exception:
            pass

        # Overview tab
        tv.add("🏠  Overview")
        self._build_category_overview(tv.tab("🏠  Overview"), key, tools, accents, tv)

        # Tool tabs (lazy)
        for t in tools:
            tv.add(t["tab"])
            self._loaded[t["tab"]] = False

        tv.pack(fill="both", expand=True, padx=10, pady=(4, 4))

        self._tabviews[key]   = tv
        self._last_tabs[key]  = ""
        self._cat_frames[key] = frame
        return frame

    def _build_category_overview(self, frame, key, tools, accents, tv=None):
        COLS  = 3 if len(tools) <= 6 else 4
        acc   = _C["gst_acc"] if key == "gst" else _C["it_acc"]
        icon  = "🏛"           if key == "gst" else "💼"
        label = "GST Tools"    if key == "gst" else "Income Tax Automation Suite"

        scroll = ctk.CTkScrollableFrame(frame, fg_color="transparent",
                                        corner_radius=0)
        scroll.pack(fill="both", expand=True)
        for c in range(COLS):
            scroll.grid_columnconfigure(c, weight=1)

        # Hero
        hero = ctk.CTkFrame(scroll, fg_color=_C["surface"], corner_radius=14,
                            border_width=1, border_color=_C["border"])
        hero.grid(row=0, column=0, columnspan=COLS,
                  padx=16, pady=(18, 12), sticky="ew")

        ctk.CTkFrame(hero, height=4, corner_radius=0,
                     fg_color=acc).pack(fill="x")

        hb = ctk.CTkFrame(hero, fg_color="transparent")
        hb.pack(fill="x", padx=24, pady=(16, 18))

        tr = ctk.CTkFrame(hb, fg_color="transparent")
        tr.pack(fill="x")
        ctk.CTkLabel(tr, text=f"{icon}  {label}",
                     font=("Segoe UI", 21, "bold"),
                     text_color=acc).pack(side="left")
        badge = ctk.CTkFrame(tr, fg_color=acc, corner_radius=20)
        badge.pack(side="left", padx=(12, 0))
        ctk.CTkLabel(badge, text=f"  {len(tools)} Tools  ",
                     font=("Segoe UI", 10, "bold"),
                     text_color="#ffffff").pack(padx=4, pady=4)

        ctk.CTkLabel(hb,
                     text="Click any tab to load a tool — "
                          "each loads once and stays active.",
                     font=("Segoe UI", 13),
                     text_color=("#1e293b", "#e2e8f0")).pack(anchor="w", pady=(8, 0))

        # Section label
        ctk.CTkLabel(scroll, text="Available Tools",
                     font=("Segoe UI", 14, "bold"),
                     text_color=("#0f172a", "#f1f5f9")).grid(
            row=1, column=0, columnspan=COLS,
            padx=18, pady=(6, 8), sticky="w")

        # Tool cards
        for idx, tool in enumerate(tools):
            ac    = accents[idx] if idx < len(accents) else acc
            rn, c = divmod(idx, COLS)

            card = ctk.CTkFrame(scroll, fg_color=_C["surface2"],
                                border_color=_C["border"], border_width=1,
                                corner_radius=12)
            card.grid(row=rn + 2, column=c, padx=9, pady=8, sticky="nsew")
            scroll.grid_rowconfigure(rn + 2, weight=0)

            ctk.CTkFrame(card, height=4, corner_radius=0,
                         fg_color=ac).pack(fill="x")

            body = ctk.CTkFrame(card, fg_color="transparent")
            body.pack(fill="both", expand=True, padx=14, pady=(10, 14))

            ctk.CTkLabel(body, text=tool["tab"],
                         font=("Segoe UI", 14, "bold"),
                         text_color=ac).pack(anchor="w")
            ctk.CTkFrame(body, height=1, corner_radius=0,
                         fg_color=_C["border"]).pack(fill="x", pady=(5, 8))
            ctk.CTkLabel(body, text=tool["desc"],
                         font=("Segoe UI", 12),
                         text_color=("#1e293b", "#e2e8f0"),
                         wraplength=230, justify="left").pack(anchor="w")
            ctk.CTkLabel(body, text="→  Click to open",
                         font=("Segoe UI", 11, "bold"),
                         text_color=ac).pack(anchor="w", pady=(10, 0))

            # Make card clickable — navigate to the tool's tab
            if tv is not None:
                tab_name = tool["tab"]
                normal_card_fg = _C["surface2"]
                hover_card_fg  = _C["surface"]

                def _enter(_, w=card, f=hover_card_fg):
                    try: w.configure(fg_color=f)
                    except Exception: pass

                def _leave(_, w=card, f=normal_card_fg):
                    # Only reset colour when pointer truly leaves the card bounds
                    try:
                        px, py = w.winfo_pointerxy()
                        cx, cy = w.winfo_rootx(), w.winfo_rooty()
                        cw, ch = w.winfo_width(), w.winfo_height()
                        if not (cx <= px <= cx + cw and cy <= py <= cy + ch):
                            w.configure(fg_color=f)
                    except Exception:
                        try: w.configure(fg_color=f)
                        except Exception: pass

                def _click(_, tv=tv, name=tab_name):
                    try: tv.set(name)
                    except Exception: pass

                def _attach(w):
                    try: w.configure(cursor="hand2")
                    except Exception: pass
                    w.bind("<Enter>",    _enter, add="+")
                    w.bind("<Leave>",    _leave, add="+")
                    w.bind("<Button-1>", _click, add="+")
                    # CTkLabel uses an internal _canvas that absorbs mouse events —
                    # bind directly to it so clicks on text labels work too
                    try:
                        if hasattr(w, "_canvas"):
                            w._canvas.bind("<Enter>",    _enter, add="+")
                            w._canvas.bind("<Leave>",    _leave, add="+")
                            w._canvas.bind("<Button-1>", _click, add="+")
                            try: w._canvas.configure(cursor="hand2")
                            except Exception: pass
                    except Exception: pass
                    for ch in w.winfo_children():
                        _attach(ch)

                _attach(card)


    # ══════════════════════════════════════════════════════════════════════════
    #  NAVIGATION
    # ══════════════════════════════════════════════════════════════════════════
    def _show_landing(self):
        self._active_poll = None
        for f in self._cat_frames.values():
            f.pack_forget()
        self._landing.pack(fill="both", expand=True)
        self._refresh_header_left("landing")

    def _show_category(self, key: str):
        self._landing.pack_forget()
        # Hide any other category that might be visible
        for k, f in self._cat_frames.items():
            if k != key:
                f.pack_forget()

        cat_frame = self._get_or_build_category(key)
        cat_frame.pack(fill="both", expand=True)
        self._active_cat  = key
        self._active_poll = key
        self._refresh_header_left("category", key)
        self.after(300, lambda: self._poll_tab(key))

    def _go_home(self):
        if self._active_cat and self._active_cat in self._cat_frames:
            self._cat_frames[self._active_cat].pack_forget()
        self._active_cat  = None
        self._active_poll = None
        self._show_landing()


    # ══════════════════════════════════════════════════════════════════════════
    #  LAZY TAB LOADER
    # ══════════════════════════════════════════════════════════════════════════
    def _poll_tab(self, key: str):
        if self._active_poll != key:
            return   # category no longer active
        tv = self._tabviews.get(key)
        if not tv:
            self.after(300, lambda: self._poll_tab(key))
            return
        try:
            current = tv.get()
        except Exception:
            self.after(300, lambda: self._poll_tab(key))
            return

        last = self._last_tabs.get(key, "")
        if current != last:
            self._last_tabs[key] = current
            if current != "🏠  Overview" and not self._loaded.get(current, True):
                self._activate_tab(key, current)

        self.after(300, lambda: self._poll_tab(key))

    def _activate_tab(self, cat_key: str, name: str):
        tools   = GST_TOOLS    if cat_key == "gst" else IT_TOOLS
        accents = _GST_ACCENTS if cat_key == "gst" else _IT_ACCENTS
        tool    = next((t for t in tools if t["tab"] == name), None)
        if tool is None:
            return
        idx    = next((i for i, t in enumerate(tools) if t["tab"] == name), 0)
        accent = accents[idx] if idx < len(accents) else _C["primary"]

        tab_frame = self._tabviews[cat_key].tab(name)

        # Loading overlay
        overlay = ctk.CTkFrame(tab_frame, fg_color=_C["surface"],
                               corner_radius=16, border_width=1,
                               border_color=accent)
        overlay.place(relx=0.5, rely=0.5, anchor="center",
                      relwidth=0.30, relheight=0.18)

        ctk.CTkLabel(overlay, text="⏳",
                     font=("Segoe UI", 28)).place(
            relx=0.15, rely=0.5, anchor="center")
        ctk.CTkLabel(overlay, text="Loading tool…",
                     font=("Segoe UI", 15, "bold"),
                     text_color=accent).place(
            relx=0.58, rely=0.36, anchor="center")
        ctk.CTkLabel(overlay, text="Please wait a moment",
                     font=("Segoe UI", 10),
                     text_color=_C["text_lo"]).place(
            relx=0.58, rely=0.65, anchor="center")

        self.update_idletasks()
        overlay.destroy()

        inst = _load_tool(tab_frame, tool["module"], tool["class"])
        self._instances[name] = inst
        self._loaded[name]    = True


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = GSTSuite()
    app.mainloop()
