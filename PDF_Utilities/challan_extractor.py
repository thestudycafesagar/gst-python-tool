"""
=============================================================
  PDF Redactor - PDF Redaction Tool
  By Studycafe | For Chartered Accountants
=============================================================
  Libraries Required:
      pip install: PyMuPDF, Pillow  (tkinter & re are built-in)

  Install Command (run in CMD):
      py -m pip install PyMuPDF Pillow

  Run: python challan_extractor.py
=============================================================
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import re
import os
import io
import threading

try:
    import fitz  # PyMuPDF
except ImportError:
    _r = tk.Tk(); _r.withdraw()
    messagebox.showerror("Missing Library",
        "PyMuPDF is not installed.\n\nRun in CMD:\n  py -m pip install PyMuPDF Pillow")
    raise SystemExit

try:
    from PIL import Image, ImageTk
except ImportError:
    _r = tk.Tk(); _r.withdraw()
    messagebox.showerror("Missing Library",
        "Pillow is not installed.\n\nRun in CMD:\n  py -m pip install PyMuPDF Pillow")
    raise SystemExit


# ─────────────────────────────────────────────────────────────────────────────
#  Colour Palettes
# ─────────────────────────────────────────────────────────────────────────────

C_DARK = {
    "bg":       "#0d1117",
    "panel":    "#161b22",
    "header":   "#1f2937",
    "accent1":  "#e94560",
    "accent2":  "#f59e0b",
    "accent3":  "#10b981",
    "accent4":  "#6366f1",
    "accent5":  "#ec4899",
    "accent6":  "#06b6d4",
    "txt":      "#f0f6fc",
    "sub":      "#8b949e",
    "border":   "#30363d",
    "canvas":   "#21262d",
}

C_LIGHT = {
    "bg":       "#f0f4f8",
    "panel":    "#ffffff",
    "header":   "#e2e8f0",
    "accent1":  "#e94560",
    "accent2":  "#d97706",
    "accent3":  "#059669",
    "accent4":  "#4f46e5",
    "accent5":  "#db2777",
    "accent6":  "#0891b2",
    "txt":      "#111827",
    "sub":      "#6b7280",
    "border":   "#cbd5e1",
    "canvas":   "#dde3ea",
}

C = dict(C_DARK)


# ─────────────────────────────────────────────────────────────────────────────
#  Page-range resolver  (from PDF Utilities main.py)
# ─────────────────────────────────────────────────────────────────────────────

def _resolve_pages(spec: str, n: int) -> list:
    if spec.strip().lower() == "all":
        return list(range(n))
    pages = set()
    for p in spec.split(","):
        p = p.strip()
        if "-" in p:
            a, b = p.split("-", 1)
            try:
                pages.update(range(int(a) - 1, int(b)))
            except ValueError:
                pass
        elif p:
            try:
                pages.add(int(p) - 1)
            except ValueError:
                pass
    return sorted(x for x in pages if 0 <= x < n)


# ─────────────────────────────────────────────────────────────────────────────
#  Main Application
# ─────────────────────────────────────────────────────────────────────────────

class ChallanExtractor:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF Redactor  |  PDF Redaction Tool")
        self.root.geometry("1280x800")
        self.root.minsize(960, 620)
        self.root.configure(bg=C["bg"])
        try:
            self.root.state("zoomed")
        except Exception:
            pass

        # ── PDF state ────────────────────────────────────────────────────────
        self.pdf_path:    str | None = None
        self.pdf_doc:     fitz.Document | None = None
        self.current_page = 0
        self.total_pages  = 0
        self.zoom         = 1.5
        self._tk_img      = None
        self.img_x0       = 10
        self.img_y0       = 10

        # ── Redaction state  (main.py engine) ────────────────────────────────
        # Format: {page_idx: [(fitz.Rect, hex_color, kind), ...]}
        # kind = "rect" | "brush"
        self._draw_rects: dict[int, list] = {}
        self._undo_stack: list = []
        self._redo_stack: list = []
        self._preview_doc:  fitz.Document | None = None
        self._preview_mode: bool = False

        # ── Drawing tool state ───────────────────────────────────────────────
        self._tool_var     = tk.StringVar(value="rect")
        self._brush_size   = 20          # half-size in PDF points
        self._brush_pts:   list = []
        self._drag_start   = None
        self._drag_item    = None

        # ── Fill colour ──────────────────────────────────────────────────────
        self._fill             = tk.StringVar(value="black")
        self._custom_color_hex = "#000000"
        self._custom_color_rgb = (0.0, 0.0, 0.0)

        # ── Keyword settings  (main.py engine) ───────────────────────────────
        self._case   = tk.BooleanVar(value=False)
        self._kpages = tk.StringVar(value="all")

        # ── Auto-Redact Fields ───────────────────────────────────────────────
        self._ar_gstn   = tk.BooleanVar(value=False)
        self._ar_pan    = tk.BooleanVar(value=False)
        self._ar_name   = tk.BooleanVar(value=False)
        self._ar_notice = tk.BooleanVar(value=False)
        self._ar_date   = tk.BooleanVar(value=False)

        self._theme = "dark"

        self._build_ui()
        self._set_status("Welcome to PDF Redactor  |  Open a PDF file to begin")

    # ─────────────────────────────────────────────────────────────────────────
    #  UI BUILD
    # ─────────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header bar
        hdr = tk.Frame(self.root, bg=C["accent1"], height=54)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="PDF Redaction Tool  |  Studycafe",
                 font=("Segoe UI", 10),
                 bg=C["accent1"], fg="#ffe0e6").pack(side="left", padx=4)

        self._theme_btn = tk.Button(
            hdr, text="☀ Light Mode",
            command=self._toggle_theme,
            bg="#c73652", fg="white", relief="flat",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2", padx=10, pady=4,
            activebackground="#a02a42", activeforeground="white"
        )
        self._theme_btn.pack(side="right", padx=12, pady=8)

        # Body
        body = tk.Frame(self.root, bg=C["bg"])
        body.pack(fill="both", expand=True)

        # Sidebar
        sidebar = tk.Frame(body, bg=C["panel"], width=260)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        self._build_sidebar(sidebar)

        # Divider
        tk.Frame(body, bg=C["border"], width=1).pack(side="left", fill="y")

        # Viewer
        viewer = tk.Frame(body, bg=C["canvas"])
        viewer.pack(side="left", fill="both", expand=True)
        self._build_viewer(viewer)

        # Status bar
        sb = tk.Frame(self.root, bg=C["header"], height=26)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)
        self._status_var    = tk.StringVar()
        self._page_info_var = tk.StringVar(value="")
        tk.Label(sb, textvariable=self._status_var,
                 font=("Segoe UI", 9), bg=C["header"], fg=C["sub"]
                 ).pack(side="left", padx=10)
        tk.Label(sb, textvariable=self._page_info_var,
                 font=("Segoe UI", 9, "bold"), bg=C["header"], fg=C["txt"]
                 ).pack(side="right", padx=10)

    # ─────────────────────────────────────────────────────────────────────────
    #  THEME TOGGLE
    # ─────────────────────────────────────────────────────────────────────────

    def _toggle_theme(self):
        old = dict(C_DARK if self._theme == "dark" else C_LIGHT)
        if self._theme == "dark":
            self._theme = "light"
            C.update(C_LIGHT)
            self._theme_btn.config(text="🌙 Dark Mode")
        else:
            self._theme = "dark"
            C.update(C_DARK)
            self._theme_btn.config(text="☀ Light Mode")
        color_map = {old[k].lower(): C[k] for k in old}
        self._recolor_widget(self.root, color_map)
        # update ttk scrollbar style
        style = ttk.Style()
        style.configure("Vertical.TScrollbar",
                        background=C["header"], troughcolor=C["bg"],
                        bordercolor=C["border"], arrowcolor=C["sub"])
        style.configure("Horizontal.TScrollbar",
                        background=C["header"], troughcolor=C["bg"],
                        bordercolor=C["border"], arrowcolor=C["sub"])
        if self.pdf_doc:
            self.canvas.configure(bg=C["canvas"])
            self.render_page()

    def _recolor_widget(self, widget, color_map):
        for opt in ("bg", "background"):
            try:
                val = widget.cget(opt)
                if isinstance(val, str) and val.lower() in color_map:
                    widget.configure(**{opt: color_map[val.lower()]})
                    break
            except Exception:
                pass
        for opt in ("fg", "foreground"):
            try:
                val = widget.cget(opt)
                if isinstance(val, str) and val.lower() in color_map:
                    widget.configure(**{opt: color_map[val.lower()]})
                    break
            except Exception:
                pass
        try:
            val = widget.cget("selectcolor")
            if isinstance(val, str) and val.lower() in color_map:
                widget.configure(selectcolor=color_map[val.lower()])
        except Exception:
            pass
        try:
            val = widget.cget("insertbackground")
            if isinstance(val, str) and val.lower() in color_map:
                widget.configure(insertbackground=color_map[val.lower()])
        except Exception:
            pass
        try:
            val = widget.cget("troughcolor")
            if isinstance(val, str) and val.lower() in color_map:
                widget.configure(troughcolor=color_map[val.lower()])
        except Exception:
            pass
        try:
            val = widget.cget("buttonbackground")
            if isinstance(val, str) and val.lower() in color_map:
                widget.configure(buttonbackground=color_map[val.lower()])
        except Exception:
            pass
        for child in widget.winfo_children():
            self._recolor_widget(child, color_map)

    # ── Sidebar ──────────────────────────────────────────────────────────────

    def _build_sidebar(self, parent):
        cv = tk.Canvas(parent, bg=C["panel"], highlightthickness=0)
        sc = ttk.Scrollbar(parent, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=sc.set)
        sc.pack(side="right", fill="y")
        cv.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(cv, bg=C["panel"])
        win_id = cv.create_window((0, 0), window=inner, anchor="nw")

        cv.bind("<Configure>",      lambda e: cv.itemconfig(win_id, width=e.width))
        inner.bind("<Configure>",   lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(-1*(e.delta//120), "units"))

        self._fill_sidebar(inner)

    def _fill_sidebar(self, p):
        pad = dict(padx=12, pady=3)

        # ── File Operations ──────────────────────────────────────────────────
        self._sec(p, "FILE OPERATIONS")
        self._btn(p, "Open PDF",           self.open_pdf,  C["accent3"], pad=pad)
        self._btn(p, "Save Redacted PDF",  self._save_redacted, C["accent1"], pad=pad)

        # ── Navigation ───────────────────────────────────────────────────────
        self._sec(p, "NAVIGATION")
        nav = tk.Frame(p, bg=C["panel"])
        nav.pack(fill="x", **pad)
        self._nbtn(nav, "Prev", self.prev_page, C["accent4"]).pack(
            side="left", expand=True, fill="x", padx=(0, 3))
        self._nbtn(nav, "Next", self.next_page, C["accent4"]).pack(
            side="left", expand=True, fill="x", padx=(3, 0))

        self._page_lbl = tk.Label(p, text="Page  -  /  -",
                                  bg=C["panel"], fg=C["txt"],
                                  font=("Segoe UI", 10, "bold"))
        self._page_lbl.pack(pady=3)

        zrow = tk.Frame(p, bg=C["panel"])
        zrow.pack(fill="x", padx=12, pady=2)
        tk.Label(zrow, text="Zoom:", bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 9)).pack(side="left")
        self._zoom_lbl = tk.Label(zrow, text=f"{int(self.zoom*100)}%",
                                  bg=C["panel"], fg=C["accent2"],
                                  font=("Segoe UI", 9, "bold"))
        self._zoom_lbl.pack(side="left", padx=6)
        self._nbtn(zrow, " + ", self.zoom_in,  C["accent2"], w=3).pack(side="right", padx=2)
        self._nbtn(zrow, " - ", self.zoom_out, C["accent2"], w=3).pack(side="right", padx=2)

        # ── Drawing Tool ─────────────────────────────────────────────────────
        self._sec(p, "DRAWING TOOL")
        tool_row = tk.Frame(p, bg=C["panel"])
        tool_row.pack(fill="x", padx=12, pady=2)

        for lbl, val, col in [("Rect Box", "rect", C["accent4"]),
                               ("Brush",    "brush", C["accent5"]),
                               ("Off",      "none",  C["sub"])]:
            rb = tk.Radiobutton(tool_row, text=lbl, variable=self._tool_var,
                                value=val, bg=C["panel"], fg=col,
                                selectcolor=C["bg"],
                                activebackground=C["panel"], activeforeground=col,
                                font=("Segoe UI", 9, "bold"),
                                cursor="hand2", relief="flat",
                                highlightthickness=0)
            rb.pack(side="left", padx=(0, 6))

        # Brush size
        brow = tk.Frame(p, bg=C["panel"])
        brow.pack(fill="x", padx=12, pady=2)
        tk.Label(brow, text="Brush size:", bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 9)).pack(side="left")
        self._brush_size_var = tk.IntVar(value=self._brush_size)
        tk.Spinbox(brow, from_=4, to=80, textvariable=self._brush_size_var,
                   width=4, bg=C["header"], fg=C["txt"],
                   font=("Segoe UI", 9), buttonbackground=C["border"],
                   relief="flat",
                   command=lambda: setattr(self, "_brush_size", self._brush_size_var.get())
                   ).pack(side="left", padx=6)

        # ── Fill Colour ──────────────────────────────────────────────────────
        self._sec(p, "REDACTION COLOUR")
        col_row = tk.Frame(p, bg=C["panel"])
        col_row.pack(fill="x", padx=12, pady=2)
        for lbl, val, col in [("Black", "black", "#333333"),
                               ("White", "white", "#cccccc"),
                               ("Red",   "red",   C["accent1"])]:
            rb = tk.Radiobutton(col_row, text=lbl, variable=self._fill,
                                value=val, bg=C["panel"], fg=col,
                                selectcolor=C["bg"],
                                activebackground=C["panel"], activeforeground=col,
                                font=("Segoe UI", 9, "bold"),
                                cursor="hand2", relief="flat",
                                highlightthickness=0,
                                command=self._invalidate_preview)
            rb.pack(side="left", padx=(0, 4))

        custom_row = tk.Frame(p, bg=C["panel"])
        custom_row.pack(fill="x", padx=12, pady=2)
        tk.Radiobutton(custom_row, text="Custom", variable=self._fill,
                       value="custom", bg=C["panel"], fg=C["accent6"],
                       selectcolor=C["bg"],
                       activebackground=C["panel"], activeforeground=C["accent6"],
                       font=("Segoe UI", 9, "bold"),
                       cursor="hand2", relief="flat",
                       highlightthickness=0,
                       command=self._invalidate_preview
                       ).pack(side="left")
        self._custom_swatch = tk.Label(custom_row, bg=self._custom_color_hex,
                                       width=3, relief="solid", cursor="hand2")
        self._custom_swatch.pack(side="left", padx=4)
        self._custom_swatch.bind("<Button-1>", lambda e: self._pick_custom_color())

        # ── Auto-Redact Fields ───────────────────────────────────────────────
        self._sec(p, "AUTO-REDACT FIELDS")
        tk.Label(p, text="Tick fields to auto-detect & blackout\nacross ALL pages:",
                 bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 8), justify="left").pack(anchor="w", padx=12, pady=(0, 4))

        ar_fields = [
            ("GSTN Number",   self._ar_gstn,   C["accent1"]),
            ("PAN Number",    self._ar_pan,     C["accent5"]),
            ("Name",          self._ar_name,    C["accent3"]),
            ("Notice Number", self._ar_notice,  C["accent2"]),
            ("Date",          self._ar_date,    C["accent4"]),
        ]
        for lbl, var, col in ar_fields:
            row = tk.Frame(p, bg=C["panel"])
            row.pack(fill="x", padx=12, pady=1)
            tk.Frame(row, bg=col, width=4).pack(side="left", fill="y", padx=(0, 6))
            tk.Checkbutton(row, text=lbl, variable=var,
                           bg=C["panel"], fg=C["txt"],
                           selectcolor=C["bg"],
                           activebackground=C["panel"], activeforeground=C["txt"],
                           font=("Segoe UI", 9),
                           cursor="hand2", relief="flat",
                           highlightthickness=0
                           ).pack(side="left", anchor="w")

        self._btn(p, "\u26a1 Apply Auto-Redact", self._apply_auto_redact,
                  C["accent4"], pad=pad)

        # ── Keyword Redact ───────────────────────────────────────────────────
        self._sec(p, "KEYWORD REDACT")
        tk.Label(p, text="One keyword / phrase per line.\nSupports regex patterns.",
                 bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 8), justify="left").pack(anchor="w", padx=12, pady=(0, 2))

        kw_frame = tk.Frame(p, bg=C["border"])
        kw_frame.pack(fill="x", padx=12, pady=(0, 4))
        self._kw = tk.Text(kw_frame, height=6, font=("Consolas", 9),
                           bg=C["header"], fg=C["txt"],
                           insertbackground=C["txt"],
                           relief="flat", wrap="word", bd=4)
        self._kw.pack(fill="x", padx=1, pady=1)
        self._kw.bind("<<Modified>>", self._on_kw_change)

        # Case sensitivity
        tk.Checkbutton(p, text="Case-sensitive matching",
                       variable=self._case,
                       bg=C["panel"], fg=C["txt"],
                       selectcolor=C["bg"],
                       activebackground=C["panel"], activeforeground=C["txt"],
                       font=("Segoe UI", 9),
                       cursor="hand2",
                       command=self._invalidate_preview
                       ).pack(anchor="w", padx=12, pady=1)

        # Page range
        prow = tk.Frame(p, bg=C["panel"])
        prow.pack(fill="x", padx=12, pady=2)
        tk.Label(prow, text="Pages:", bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 9)).pack(side="left")
        tk.Entry(prow, textvariable=self._kpages, width=10,
                 bg=C["header"], fg=C["txt"],
                 insertbackground=C["txt"], relief="flat", bd=4,
                 font=("Segoe UI", 9)).pack(side="left", padx=6)
        tk.Label(prow, text='e.g. "all" or "1,3-5"',
                 bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 8)).pack(side="left")

        self._btn(p, "Preview Redactions", self._preview_redactions,
                  C["accent6"], pad=pad)

        # ── Undo / Redo ──────────────────────────────────────────────────────
        self._sec(p, "UNDO / REDO")
        ur_row = tk.Frame(p, bg=C["panel"])
        ur_row.pack(fill="x", padx=12, pady=2)
        self._nbtn(ur_row, "Undo", self._undo, C["accent2"]).pack(
            side="left", expand=True, fill="x", padx=(0, 3))
        self._nbtn(ur_row, "Redo", self._redo, C["accent3"]).pack(
            side="left", expand=True, fill="x", padx=(3, 0))

        # ── Clear ────────────────────────────────────────────────────────────
        self._sec(p, "CLEAR MARKS")
        self._btn(p, "Clear This Page", self._clear_page, C["accent5"], pad=pad)
        self._btn(p, "Clear All Pages", self._clear_all,  "#7f1d1d",   pad=pad)

        self._boxes_lbl = tk.Label(p, text="No marks yet.",
                                   bg=C["panel"], fg=C["sub"],
                                   font=("Segoe UI", 8), justify="left")
        self._boxes_lbl.pack(anchor="w", padx=14, pady=2)

        # ── Help ─────────────────────────────────────────────────────────────
        self._sec(p, "HOW TO USE")
        tk.Label(p,
                 text=(
                     "1. Open a PDF file.\n"
                     "2. Draw boxes (Rect) or freehand\n"
                     "   (Brush) on the viewer.\n"
                     "3. OR enter keywords and click\n"
                     "   Preview Redactions.\n"
                     "4. Click Save Redacted PDF."
                 ),
                 bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 8), justify="left"
                 ).pack(anchor="w", padx=12, pady=4)

        tk.Label(p, text="PDF Redactor  v2.0\n(c) 2025 Studycafe",
                 bg=C["panel"], fg=C["border"],
                 font=("Segoe UI", 8), justify="center"
                 ).pack(pady=(16, 8))

    # ── PDF Viewer ────────────────────────────────────────────────────────────

    def _build_viewer(self, parent):
        tbar = tk.Frame(parent, bg=C["header"], height=36)
        tbar.pack(fill="x")
        tbar.pack_propagate(False)
        tk.Label(tbar, text="PDF Viewer  |  Draw rectangles or use brush to redact",
                 bg=C["header"], fg=C["sub"], font=("Segoe UI", 9)
                 ).pack(side="left", padx=12, pady=8)

        self._preview_badge = tk.Label(tbar, text="  PREVIEW MODE  ",
                                       bg=C["accent2"], fg="#000",
                                       font=("Segoe UI", 8, "bold"))
        # (packed dynamically when entering preview mode)

        cf = tk.Frame(parent, bg=C["canvas"])
        cf.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(cf, bg=C["canvas"],
                                cursor="crosshair", highlightthickness=0)
        vs = ttk.Scrollbar(cf, orient="vertical",   command=self.canvas.yview)
        hs = ttk.Scrollbar(cf, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)

        vs.pack(side="right",  fill="y")
        hs.pack(side="bottom", fill="x")
        self.canvas.pack(fill="both", expand=True)

        self.canvas.bind("<ButtonPress-1>",   self._on_press)
        self.canvas.bind("<B1-Motion>",       self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<MouseWheel>",
                         lambda e: self.canvas.yview_scroll(-1*(e.delta//120), "units"))

        self._welcome = tk.Label(
            self.canvas,
            text="Open a PDF file to begin\n\nUse the sidebar to navigate and redact",
            bg=C["canvas"], fg=C["sub"],
            font=("Segoe UI", 14), justify="center"
        )
        self._welcome.place(relx=0.5, rely=0.45, anchor="center")

    # ─────────────────────────────────────────────────────────────────────────
    #  RENDERING
    # ─────────────────────────────────────────────────────────────────────────

    def render_page(self):
        active_doc = self._preview_doc if self._preview_mode else self.pdf_doc
        if not active_doc:
            return

        page = active_doc[self.current_page]
        mat  = fitz.Matrix(self.zoom, self.zoom)
        pix  = page.get_pixmap(matrix=mat, alpha=False)

        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self._tk_img = ImageTk.PhotoImage(img)

        self.canvas.delete("all")
        if hasattr(self, "_welcome"):
            self._welcome.place_forget()

        self.root.update_idletasks()
        cw = self.canvas.winfo_width()
        self.img_x0 = max(10, (cw - pix.width) // 2)
        self.img_y0 = 10

        sr_w = max(pix.width  + 2 * self.img_x0, cw)
        sr_h = max(pix.height + 2 * self.img_y0, self.canvas.winfo_height())
        self.canvas.configure(scrollregion=(0, 0, sr_w, sr_h))
        self.canvas.create_image(self.img_x0, self.img_y0,
                                 anchor="nw", image=self._tk_img)

        # Re-draw committed marks (edit mode only)
        if not self._preview_mode:
            for r, hx, kd in self._draw_rects.get(self.current_page, []):
                self._paint_box(r, hx, kd)

        n_pages = active_doc.page_count
        self._page_lbl.config(text=f"Page  {self.current_page+1}  /  {n_pages}")
        total_marks = sum(len(v) for v in self._draw_rects.values())
        self._page_info_var.set(
            f"Page {self.current_page+1} of {n_pages}  |  "
            f"Zoom {int(self.zoom*100)}%  |  Marks: {total_marks}"
        )

    def _paint_box(self, pdf_rect: fitz.Rect, hex_col: str = "#111111",
                   kind: str = "rect"):
        s  = self.zoom
        ox, oy = self.img_x0, self.img_y0
        x0, y0 = pdf_rect.x0 * s + ox, pdf_rect.y0 * s + oy
        x1, y1 = pdf_rect.x1 * s + ox, pdf_rect.y1 * s + oy
        if kind == "brush":
            self.canvas.create_oval(x0, y0, x1, y1,
                                    fill=hex_col, outline="", tags="mark")
        else:
            self.canvas.create_rectangle(x0, y0, x1, y1,
                                         fill=hex_col, outline="", width=0, tags="mark")

    # ─────────────────────────────────────────────────────────────────────────
    #  MOUSE DRAWING  (main.py engine: rect + brush with undo/redo)
    # ─────────────────────────────────────────────────────────────────────────

    def _on_press(self, event):
        tool = self._tool_var.get()
        if tool == "none" or not self.pdf_doc:
            return
        if self._preview_mode:
            self._exit_preview_mode()

        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)

        if tool == "rect":
            self._drag_start = (cx, cy)
            self._drag_item  = self.canvas.create_rectangle(
                cx, cy, cx, cy, outline=C["accent2"], fill="", width=2, dash=(4, 2))

        elif tool == "brush":
            self._brush_size = self._brush_size_var.get()
            self._brush_pts  = [(cx, cy)]
            br  = max(4, self._brush_size * self.zoom / 2)
            col = self._fill_hex_preview()
            self.canvas.create_oval(cx - br, cy - br, cx + br, cy + br,
                                    fill=col, outline="", tags="brush_live")

    def _on_drag(self, event):
        tool = self._tool_var.get()
        cx   = self.canvas.canvasx(event.x)
        cy   = self.canvas.canvasy(event.y)

        if tool == "rect":
            if self._drag_start and self._drag_item:
                self.canvas.coords(self._drag_item, *self._drag_start, cx, cy)

        elif tool == "brush":
            if not self._brush_pts:
                return
            lx, ly = self._brush_pts[-1]
            if ((cx - lx)**2 + (cy - ly)**2)**0.5 > 5:
                self._brush_pts.append((cx, cy))
                br  = max(4, self._brush_size * self.zoom / 2)
                col = self._fill_hex_preview()
                self.canvas.create_oval(cx - br, cy - br, cx + br, cy + br,
                                        fill=col, outline="", tags="brush_live")

    def _on_release(self, event):
        tool = self._tool_var.get()
        if tool == "none" or not self.pdf_doc:
            return

        cx  = self.canvas.canvasx(event.x)
        cy  = self.canvas.canvasy(event.y)
        s   = self.zoom
        ox  = self.img_x0
        oy  = self.img_y0
        hx  = self._fill_hex()
        drew = False
        action_rects = []

        if tool == "rect":
            if not self._drag_start or not self._drag_item:
                return
            x0c, y0c = self._drag_start
            self.canvas.delete(self._drag_item)
            self._drag_start = None
            self._drag_item  = None

            px0 = (min(x0c, cx) - ox) / s
            py0 = (min(y0c, cy) - oy) / s
            px1 = (max(x0c, cx) - ox) / s
            py1 = (max(y0c, cy) - oy) / s

            if abs(px1 - px0) < 4 or abs(py1 - py0) < 4:
                return

            r = fitz.Rect(px0, py0, px1, py1)
            self._draw_rects.setdefault(self.current_page, []).append((r, hx, "rect"))
            self._paint_box(r, hx, "rect")
            action_rects = [(r, hx, "rect")]
            drew = True

        elif tool == "brush":
            if not self._brush_pts:
                return
            self.canvas.delete("brush_live")
            half    = self._brush_size / 2
            pg_rect = self.pdf_doc[self.current_page].rect

            sampled = self._brush_pts[::2] if len(self._brush_pts) > 1 else self._brush_pts
            for bx, by in sampled:
                px = (bx - ox) / s
                py = (by - oy) / s
                r  = fitz.Rect(px - half, py - half, px + half, py + half)
                r  = r & pg_rect
                if r.is_empty or r.width < 1 or r.height < 1:
                    continue
                self._draw_rects.setdefault(self.current_page, []).append((r, hx, "brush"))
                self._paint_box(r, hx, "brush")
                action_rects.append((r, hx, "brush"))
                drew = True
            self._brush_pts = []

        if not drew:
            return

        self._undo_stack.append((self.current_page, action_rects))
        self._redo_stack.clear()
        self._invalidate_preview()
        self._update_boxes_label()
        total = sum(len(v) for v in self._draw_rects.values())
        self._set_status(
            f"Mark added on page {self.current_page+1}  |  Total marks: {total}")

    # ─────────────────────────────────────────────────────────────────────────
    #  PREVIEW MODE
    # ─────────────────────────────────────────────────────────────────────────

    def _enter_preview_mode(self):
        self._preview_mode = True
        self.canvas.configure(bg="#1a1200" if self._theme == "dark" else "#fffbea")
        self._preview_badge.pack(side="right", padx=8, pady=4)
        self.render_page()

    def _exit_preview_mode(self):
        self._preview_mode = False
        self.canvas.configure(bg=C["canvas"])
        self._preview_badge.pack_forget()
        self.render_page()

    # ─────────────────────────────────────────────────────────────────────────
    #  KEYWORD REDACT ENGINE  (ported from main.py)
    # ─────────────────────────────────────────────────────────────────────────

    def _apply_auto_redact(self):
        if not self.pdf_doc:
            messagebox.showwarning("No PDF", "Please open a PDF file first.")
            return

        # Regex patterns for each field
        patterns = []
        if self._ar_gstn.get():
            patterns.append(r"\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b")
        if self._ar_pan.get():
            patterns.append(r"\b[A-Z]{5}[0-9]{4}[A-Z]{1}\b")
        if self._ar_name.get():
            patterns.append(r"\bName\s*[:\-]?\s*[A-Za-z\s]{2,40}")
        if self._ar_notice.get():
            patterns.append(r"\bNotice\s*No\.?\s*[:\-]?\s*[\w\-/]+")
        if self._ar_date.get():
            patterns.append(
                r"\b\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4}\b"
                r"|\b\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4}\b"
            )

        if not patterns:
            messagebox.showwarning("Nothing Selected",
                                   "Please tick at least one field to auto-redact.")
            return

        fill = self._fill_color()
        src  = self.pdf_path

        self._set_status("Applying auto-redact — please wait...")
        self.root.update_idletasks()

        def run():
            try:
                doc   = fitz.open(src)
                total = 0
                for pi in range(doc.page_count):
                    pg  = doc[pi]
                    txt = pg.get_text("text")
                    for pat in patterns:
                        for m in re.finditer(pat, txt, re.IGNORECASE):
                            rects = pg.search_for(m.group(0))
                            for r in rects:
                                pg.add_redact_annot(r, fill=fill)
                                total += 1
                    pg.apply_redactions()

                pdf_bytes = doc.tobytes(garbage=4, deflate=True)
                doc.close()

                if self._preview_doc:
                    self._preview_doc.close()
                self._preview_doc = fitz.open(
                    stream=io.BytesIO(pdf_bytes), filetype="pdf")

                self.root.after(0, lambda: self._on_preview_ready(total))
            except Exception as exc:
                self.root.after(0, lambda e=exc: (
                    self._set_status(f"Auto-redact error: {e}"),
                    messagebox.showerror("Auto-Redact Error", str(e))
                ))

        threading.Thread(target=run, daemon=True).start()

    def _preview_redactions(self):
        if not self.pdf_doc:
            messagebox.showwarning("No PDF", "Please open a PDF file first.")
            return

        kws = [ln.strip()
               for ln in self._kw.get("1.0", "end").splitlines() if ln.strip()]
        manual_total = sum(len(v) for v in self._draw_rects.values())

        if not kws and manual_total == 0:
            messagebox.showwarning("Nothing to Redact",
                                   "Enter keywords and/or draw redaction marks.")
            return

        fill       = self._fill_color()
        case       = self._case.get()
        kpg        = self._kpages.get()
        draw_rects = {k: list(v) for k, v in self._draw_rects.items()}
        src        = self.pdf_path

        self._set_status("Building preview — please wait...")
        self.root.update_idletasks()

        _RE_SPECIAL = set(r'\.^$*+?{}[]|()')

        def _is_regex(s):
            return any(c in _RE_SPECIAL for c in s)

        def _redact_kw(pg, kw, case_sensitive):
            count   = 0
            flags   = 0 if case_sensitive else re.IGNORECASE
            pattern = kw if _is_regex(kw) else re.escape(kw)
            txt     = pg.get_text("text")
            matches = [m.group(0) for m in re.finditer(pattern, txt, flags)
                       if m.group(0).strip()]
            if not matches:
                return 0
            words = pg.get_text("words")
            for matched in set(matches):
                tokens = matched.split()
                n_tok  = len(tokens)
                if n_tok == 0:
                    continue
                if n_tok == 1:
                    for wd in words:
                        ws = wd[4]
                        ok = (ws == matched) if case_sensitive \
                             else (ws.lower() == matched.lower())
                        if ok:
                            pg.add_redact_annot(fitz.Rect(wd[:4]), fill=fill)
                            count += 1
                else:
                    for i in range(len(words) - n_tok + 1):
                        chunk     = words[i:i + n_tok]
                        chunk_str = " ".join(w[4] for w in chunk)
                        ok = (chunk_str == matched) if case_sensitive \
                             else (chunk_str.lower() == matched.lower())
                        if ok:
                            pg.add_redact_annot(
                                fitz.Rect(min(w[0] for w in chunk),
                                          min(w[1] for w in chunk),
                                          max(w[2] for w in chunk),
                                          max(w[3] for w in chunk)),
                                fill=fill)
                            count += 1
            return count

        def run():
            try:
                doc   = fitz.open(src)
                n     = doc.page_count
                total = 0
                for pi in _resolve_pages(kpg, n):
                    pg = doc[pi]
                    for kw in kws:
                        total += _redact_kw(pg, kw, case)
                    page_marks = draw_rects.get(pi, [])
                    for r, hx, kd in page_marks:
                        h  = hx.lstrip("#")
                        rc = (int(h[0:2], 16)/255,
                              int(h[2:4], 16)/255,
                              int(h[4:6], 16)/255)
                        pg.add_redact_annot(r, fill=(1, 1, 1) if kd == "brush" else rc)
                        total += 1
                    pg.apply_redactions()
                    # Second pass: draw smooth ovals for brush strokes
                    for r, hx, kd in page_marks:
                        if kd == "brush":
                            h  = hx.lstrip("#")
                            rc = (int(h[0:2], 16)/255,
                                  int(h[2:4], 16)/255,
                                  int(h[4:6], 16)/255)
                            pg.draw_oval(r, color=rc, fill=rc, width=0)

                pdf_bytes = doc.tobytes(garbage=4, deflate=True)
                doc.close()

                if self._preview_doc:
                    self._preview_doc.close()
                self._preview_doc = fitz.open(
                    stream=io.BytesIO(pdf_bytes), filetype="pdf")

                self.root.after(0, lambda: self._on_preview_ready(total))
            except Exception as exc:
                self.root.after(0, lambda e=exc: (
                    self._set_status(f"Preview error: {e}"),
                    messagebox.showerror("Preview Error", str(e))
                ))

        threading.Thread(target=run, daemon=True).start()

    def _on_preview_ready(self, total: int):
        self._set_status(
            f"Preview ready — {total} area(s) redacted. "
            "Inspect, then click Save Redacted PDF.")
        self._enter_preview_mode()

    # ─────────────────────────────────────────────────────────────────────────
    #  FILE OPERATIONS
    # ─────────────────────────────────────────────────────────────────────────

    def open_pdf(self):
        path = filedialog.askopenfilename(
            title="Open PDF File",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
        if not path:
            return
        try:
            if self.pdf_doc:
                self.pdf_doc.close()
            self._invalidate_preview()
            self.pdf_doc      = fitz.open(path)
            self.pdf_path     = path
            self.current_page = 0
            self.total_pages  = len(self.pdf_doc)
            self._draw_rects.clear()
            self._undo_stack.clear()
            self._redo_stack.clear()
            self._update_boxes_label()
            self.root.update_idletasks()
            self.render_page()
            self._set_status(
                f"Opened: {os.path.basename(path)}  |  {self.total_pages} page(s)")
        except Exception as exc:
            messagebox.showerror("Open Error", f"Could not open PDF:\n{exc}")

    def _save_redacted(self):
        if not self.pdf_doc:
            messagebox.showwarning("No PDF", "Please open a PDF file first.")
            return

        # If no preview yet, build one first, then save
        if self._preview_doc is None:
            manual_total = sum(len(v) for v in self._draw_rects.values())
            kws = [ln.strip()
                   for ln in self._kw.get("1.0", "end").splitlines() if ln.strip()]
            if not kws and manual_total == 0:
                messagebox.showwarning("Nothing to Redact",
                                       "Enter keywords and/or draw redaction marks first.")
                return
            self._preview_redactions()
            self.root.after(300, self._save_redacted)
            return

        default_name = "redacted_" + os.path.basename(self.pdf_path)
        save_path = filedialog.asksaveasfilename(
            title="Save Redacted PDF",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=default_name)
        if not save_path:
            return

        self._set_status("Saving...")

        def run():
            try:
                self._preview_doc.save(save_path, garbage=4, deflate=True)
                self.root.after(0, lambda: (
                    self._set_status(f"Saved: {os.path.basename(save_path)}"),
                    messagebox.showinfo("Saved Successfully",
                                        f"Redacted PDF saved!\n\nFile: {save_path}")
                ))
            except Exception as exc:
                self.root.after(0, lambda e=exc: (
                    self._set_status(f"Save error: {e}"),
                    messagebox.showerror("Save Error", str(e))
                ))

        threading.Thread(target=run, daemon=True).start()

    # ─────────────────────────────────────────────────────────────────────────
    #  NAVIGATION
    # ─────────────────────────────────────────────────────────────────────────

    def prev_page(self):
        doc = self._preview_doc if self._preview_mode else self.pdf_doc
        if doc and self.current_page > 0:
            self.current_page -= 1
            self.render_page()

    def next_page(self):
        doc = self._preview_doc if self._preview_mode else self.pdf_doc
        if doc and self.current_page < doc.page_count - 1:
            self.current_page += 1
            self.render_page()

    def zoom_in(self):
        self.zoom = min(4.0, round(self.zoom + 0.25, 2))
        self._zoom_lbl.config(text=f"{int(self.zoom*100)}%")
        self.render_page()

    def zoom_out(self):
        self.zoom = max(0.5, round(self.zoom - 0.25, 2))
        self._zoom_lbl.config(text=f"{int(self.zoom*100)}%")
        self.render_page()

    # ─────────────────────────────────────────────────────────────────────────
    #  CLEAR / UNDO / REDO
    # ─────────────────────────────────────────────────────────────────────────

    def _clear_page(self):
        if not self.pdf_doc:
            return
        self._draw_rects.pop(self.current_page, None)
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._invalidate_preview()
        self.render_page()
        self._update_boxes_label()
        self._set_status(f"Cleared all marks on page {self.current_page+1}")

    def _clear_all(self):
        if not self.pdf_doc:
            return
        if not messagebox.askyesno("Clear All Marks",
                                    "Remove ALL redaction marks from ALL pages?\n"
                                    "This cannot be undone."):
            return
        self._draw_rects.clear()
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._invalidate_preview()
        self.render_page()
        self._update_boxes_label()
        self._set_status("All redaction marks cleared")

    def _undo(self):
        if not self._undo_stack:
            self._set_status("Nothing to undo")
            return
        page_idx, added = self._undo_stack.pop()
        page_list = self._draw_rects.get(page_idx, [])
        n = len(added)
        if n and len(page_list) >= n:
            del page_list[-n:]
        if not page_list:
            self._draw_rects.pop(page_idx, None)
        self._redo_stack.append((page_idx, added))
        self._invalidate_preview()
        self.render_page()
        self._update_boxes_label()
        self._set_status(f"Undo — {n} mark(s) removed")

    def _redo(self):
        if not self._redo_stack:
            self._set_status("Nothing to redo")
            return
        page_idx, added = self._redo_stack.pop()
        self._draw_rects.setdefault(page_idx, []).extend(added)
        self._undo_stack.append((page_idx, added))
        self._invalidate_preview()
        self.render_page()
        self._update_boxes_label()
        self._set_status(f"Redo — {len(added)} mark(s) restored")

    # ─────────────────────────────────────────────────────────────────────────
    #  COLOUR HELPERS
    # ─────────────────────────────────────────────────────────────────────────

    def _fill_color(self):
        c = self._fill.get()
        if c == "black": return (0, 0, 0)
        if c == "white": return (1, 1, 1)
        if c == "red":   return (0.8, 0, 0)
        return self._custom_color_rgb

    def _fill_hex(self):
        c = self._fill.get()
        if c == "black": return "#111111"
        if c == "white": return "#eeeeee"
        if c == "red":   return "#cc2222"
        return self._custom_color_hex

    def _fill_hex_preview(self):
        """Slightly transparent colour for live brush preview on canvas."""
        c = self._fill.get()
        if c == "black": return "#555555"
        if c == "white": return "#dddddd"
        if c == "red":   return "#ff4444"
        return self._custom_color_hex

    def _pick_custom_color(self):
        result = colorchooser.askcolor(
            color=self._custom_color_hex,
            title="Pick Redaction Colour",
            parent=self.root)
        if result and result[0]:
            r, g, b = result[0]
            self._custom_color_rgb = (r/255, g/255, b/255)
            self._custom_color_hex = result[1]
            self._fill.set("custom")
            self._custom_swatch.configure(bg=self._custom_color_hex)
            self._invalidate_preview()

    # ─────────────────────────────────────────────────────────────────────────
    #  INVALIDATION / HELPERS
    # ─────────────────────────────────────────────────────────────────────────

    def _invalidate_preview(self):
        if self._preview_doc is not None:
            self._preview_doc.close()
            self._preview_doc = None
        if self._preview_mode:
            self._exit_preview_mode()

    def _on_kw_change(self, _=None):
        self._kw.edit_modified(False)
        self._invalidate_preview()

    def _update_boxes_label(self):
        total = sum(len(v) for v in self._draw_rects.values())
        if total == 0:
            self._boxes_lbl.configure(text="No marks yet.")
        else:
            parts = [f"p{p+1}: {len(self._draw_rects[p])}"
                     for p in sorted(self._draw_rects)]
            self._boxes_lbl.configure(
                text=f"{total} mark(s)\n" + "  ".join(parts))

    def _set_status(self, msg: str):
        self._status_var.set(msg)

    # ─────────────────────────────────────────────────────────────────────────
    #  UI HELPERS
    # ─────────────────────────────────────────────────────────────────────────

    @staticmethod
    def _btn(parent, text, cmd, bg, pad=None):
        if pad is None:
            pad = dict(padx=10, pady=3)
        import colorsys
        b = tk.Button(parent, text=text, command=cmd,
                      bg=bg, fg="white", relief="flat",
                      font=("Segoe UI", 9, "bold"),
                      cursor="hand2", pady=6, padx=8,
                      activeforeground="white")
        b.pack(fill="x", **pad)

        def _darken(col):
            col = col.lstrip("#")
            r, g, bv = int(col[0:2],16)/255, int(col[2:4],16)/255, int(col[4:6],16)/255
            h, s, v  = colorsys.rgb_to_hsv(r, g, bv)
            r2,g2,b2 = colorsys.hsv_to_rgb(h, s, max(0, v-0.12))
            return "#{:02x}{:02x}{:02x}".format(int(r2*255),int(g2*255),int(b2*255))

        hover = _darken(bg)
        b.bind("<Enter>", lambda e: b.config(bg=hover))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    @staticmethod
    def _nbtn(parent, text, cmd, bg, w=None):
        kw = dict(width=w) if w else {}
        return tk.Button(parent, text=text, command=cmd,
                         bg=bg, fg="white", relief="flat",
                         font=("Segoe UI", 9, "bold"),
                         cursor="hand2", pady=5, **kw)

    @staticmethod
    def _sec(parent, label):
        tk.Frame(parent, bg=C["panel"]).pack(pady=(10, 0))
        f = tk.Frame(parent, bg=C["panel"])
        f.pack(fill="x", padx=0, pady=(2, 0))
        tk.Label(f, text=label, bg=C["panel"], fg=C["accent2"],
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=10)
        tk.Frame(parent, bg=C["border"], height=1).pack(fill="x", padx=10)


# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()

    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("Vertical.TScrollbar",
                    background=C["header"], troughcolor=C["bg"],
                    bordercolor=C["border"], arrowcolor=C["sub"])
    style.configure("Horizontal.TScrollbar",
                    background=C["header"], troughcolor=C["bg"],
                    bordercolor=C["border"], arrowcolor=C["sub"])

    app = ChallanExtractor(root)

    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    w, h   = 1280, 800
    root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    root.mainloop()


if __name__ == "__main__":
    main()
