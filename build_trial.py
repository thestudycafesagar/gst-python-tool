"""
build_trial.py  —  GUI Trial EXE Builder
=========================================
Run:  python build_trial.py
"""

import os
import re
import shutil
import subprocess
import sys
import threading
from datetime import datetime, timedelta

import customtkinter as ctk

# ── Theme ──────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ── Paths ──────────────────────────────────────────────────────────────────────
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
SOURCE_FILE  = os.path.join(SCRIPT_DIR, "GST_Suite.py")
SPEC_FILE    = os.path.join(SCRIPT_DIR, "GST_Suite.spec")
DIST_DIR     = os.path.join(SCRIPT_DIR, "dist")
ORIGINAL_EXE = os.path.join(DIST_DIR, "GST_Suite.exe")
TRIAL_PATTERN = re.compile(r"^TRIAL_EXPIRY\s*=\s*None\s*$", re.MULTILINE)

# ── Colour tokens ──────────────────────────────────────────────────────────────
_BG      = ("#f8fafc", "#0f172a")
_SURFACE = ("#ffffff", "#1e293b")
_BORDER  = ("#e2e8f0", "#334155")
_PRIMARY = ("#4f46e5", "#6366f1")
_PRI_HOV = ("#4338ca", "#4f46e5")
_TEXT_HI = ("#0f172a", "#f1f5f9")
_TEXT_MID= ("#475569", "#94a3b8")
_GREEN   = ("#059669", "#34d399")
_RED     = ("#dc2626", "#f87171")
_AMBER   = ("#d97706", "#fbbf24")


class BuildTrialApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Trial EXE Builder  —  GST & IT Automation Suite")
        self.geometry("860x680")
        self.minsize(760, 600)
        self.resizable(True, True)
        self._building   = False
        self._original   = None      # original source content saved before patch

        self._build_ui()
        self._refresh_preview()

    # ══════════════════════════════════════════════════════════════════════════
    #  UI
    # ══════════════════════════════════════════════════════════════════════════
    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color=("#1e293b", "#060c18"), corner_radius=0, height=54)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        ctk.CTkFrame(hdr, height=3, corner_radius=0, fg_color=_PRIMARY).pack(fill="x")
        ctk.CTkLabel(hdr, text="🛠   Trial EXE Builder",
                     font=("Segoe UI", 16, "bold"),
                     text_color="#f1f5f9").pack(side="left", padx=20, pady=8)
        ctk.CTkLabel(hdr, text="GST & Income Tax Automation Suite",
                     font=("Segoe UI", 11),
                     text_color="#64748b").pack(side="right", padx=20)

        # ── Log panel — docked bottom ─────────────────────────────────────────
        log_panel = ctk.CTkFrame(self, fg_color=("#0f172a", "#020617"),
                                 corner_radius=0)
        log_panel.pack(fill="both", side="bottom", expand=True)

        log_hdr = ctk.CTkFrame(log_panel, fg_color=("#1e293b", "#0f172a"),
                               corner_radius=0, height=30)
        log_hdr.pack(fill="x")
        log_hdr.pack_propagate(False)
        ctk.CTkFrame(log_hdr, height=2, corner_radius=0,
                     fg_color=_PRIMARY).pack(fill="x", side="top")
        ctk.CTkLabel(log_hdr, text="📋  Build Log",
                     font=("Segoe UI", 11, "bold"),
                     text_color=_TEXT_HI).pack(side="left", padx=14)

        self._log = ctk.CTkTextbox(log_panel,
                                   font=("Consolas", 11),
                                   fg_color=("#0f172a", "#020617"),
                                   text_color=("#94a3b8", "#94a3b8"),
                                   corner_radius=0, border_width=0)
        self._log.pack(fill="both", expand=True, padx=6, pady=4)
        self._log.configure(state="disabled")

        # ── Config area (fixed, no scroll) ────────────────────────────────────
        body = ctk.CTkFrame(self, fg_color=_BG, corner_radius=0)
        body.pack(fill="x", side="top", padx=0, pady=0)

        # ── Row 1: Expiry + Preview side by side ──────────────────────────────
        row1 = ctk.CTkFrame(body, fg_color="transparent")
        row1.pack(fill="x", padx=16, pady=(14, 8))
        row1.columnconfigure(0, weight=3)
        row1.columnconfigure(1, weight=2)

        # Left — Set Expiry
        expiry_outer = ctk.CTkFrame(row1, fg_color=_SURFACE, corner_radius=12,
                                    border_width=1, border_color=_BORDER)
        expiry_outer.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ctk.CTkFrame(expiry_outer, height=3, corner_radius=12,
                     fg_color=_PRIMARY).pack(fill="x")
        ctk.CTkLabel(expiry_outer, text="⏰  Set Trial Expiry",
                     font=("Segoe UI", 12, "bold"),
                     text_color=_TEXT_HI).pack(anchor="w", padx=14, pady=(8, 4))
        ctk.CTkFrame(expiry_outer, height=1, fg_color=_BORDER,
                     corner_radius=0).pack(fill="x", padx=14)
        expiry_inner = ctk.CTkFrame(expiry_outer, fg_color="transparent")
        expiry_inner.pack(fill="x", padx=14, pady=10)

        # Mode toggle
        mode_row = ctk.CTkFrame(expiry_inner, fg_color="transparent")
        mode_row.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(mode_row, text="Mode:",
                     font=("Segoe UI", 11),
                     text_color=_TEXT_MID).pack(side="left", padx=(0, 10))
        self._mode_var = ctk.StringVar(value="📅  By Date")
        ctk.CTkSegmentedButton(
            mode_row,
            values=["📅  By Date", "⏱  By Hours"],
            variable=self._mode_var,
            command=self._on_mode_change,
            font=("Segoe UI", 11, "bold"),
            selected_color=_PRIMARY,
            selected_hover_color=_PRI_HOV,
            width=240, height=30,
        ).pack(side="left")

        # Date inputs
        self._date_frame = ctk.CTkFrame(expiry_inner, fg_color="transparent")
        self._date_frame.pack(fill="x")
        date_row = ctk.CTkFrame(self._date_frame, fg_color="transparent")
        date_row.pack(fill="x", pady=(0, 2))

        self._day_var   = ctk.StringVar(value=str(datetime.now().day))
        self._month_var = ctk.StringVar(value=str(datetime.now().month))
        self._year_var  = ctk.StringVar(value=str(datetime.now().year + 1))
        self._hour_var  = ctk.StringVar(value="23")
        self._min_var   = ctk.StringVar(value="59")

        for label, var, vals, w in [
            ("Day",   self._day_var,   [str(d) for d in range(1, 32)],   62),
            ("Month", self._month_var, [str(m) for m in range(1, 13)],   62),
            ("Year",  self._year_var,  [str(y) for y in range(datetime.now().year, datetime.now().year + 6)], 82),
            ("Hour",  self._hour_var,  [f"{h:02d}" for h in range(0, 24)], 62),
            ("Min",   self._min_var,   [f"{m:02d}" for m in range(0, 60)], 62),
        ]:
            col = ctk.CTkFrame(date_row, fg_color="transparent")
            col.pack(side="left", padx=(0, 8))
            ctk.CTkLabel(col, text=label, font=("Segoe UI", 10),
                         text_color=_TEXT_MID).pack(anchor="w")
            ctk.CTkOptionMenu(col, variable=var, values=vals,
                              width=w, height=30,
                              font=("Segoe UI", 11, "bold"),
                              fg_color=_SURFACE,
                              button_color=_PRIMARY,
                              button_hover_color=_PRI_HOV,
                              command=lambda _=None: self._refresh_preview()).pack()
            var.trace_add("write", lambda *_: self._refresh_preview())

        # Hours input
        self._hours_frame = ctk.CTkFrame(expiry_inner, fg_color="transparent")

        hours_row = ctk.CTkFrame(self._hours_frame, fg_color="transparent")
        hours_row.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(hours_row, text="From now:",
                     font=("Segoe UI", 11), text_color=_TEXT_MID).pack(side="left", padx=(0, 10))
        self._hours_var = ctk.StringVar(value="72")
        ctk.CTkEntry(hours_row, textvariable=self._hours_var,
                     width=72, height=30, font=("Segoe UI", 13, "bold"),
                     justify="center").pack(side="left")
        ctk.CTkLabel(hours_row, text="hrs",
                     font=("Segoe UI", 11), text_color=_TEXT_MID).pack(side="left", padx=(5, 14))
        self._hmin_var = ctk.StringVar(value="0")
        ctk.CTkEntry(hours_row, textvariable=self._hmin_var,
                     width=62, height=30, font=("Segoe UI", 13, "bold"),
                     justify="center").pack(side="left")
        ctk.CTkLabel(hours_row, text="min",
                     font=("Segoe UI", 11), text_color=_TEXT_MID).pack(side="left", padx=(5, 0))
        self._hours_var.trace_add("write", lambda *_: self._refresh_preview())
        self._hmin_var.trace_add("write",  lambda *_: self._refresh_preview())

        quick_row = ctk.CTkFrame(self._hours_frame, fg_color="transparent")
        quick_row.pack(fill="x", pady=(4, 0))
        for h in [24, 48, 72, 168]:
            lbl = f"{h}h" if h < 168 else "7 days"
            ctk.CTkButton(quick_row, text=lbl, width=66, height=26,
                          font=("Segoe UI", 10),
                          fg_color=_SURFACE, hover_color=("#e2e8f0", "#334155"),
                          text_color=_TEXT_HI, border_width=1, border_color=_BORDER,
                          command=lambda v=h: self._set_hours(v)).pack(side="left", padx=(0, 6))

        # Right — Preview + Output
        right_col = ctk.CTkFrame(row1, fg_color="transparent")
        right_col.grid(row=0, column=1, sticky="nsew")

        # Preview box
        prev_outer = ctk.CTkFrame(right_col, fg_color=_SURFACE, corner_radius=12,
                                  border_width=1, border_color=_BORDER)
        prev_outer.pack(fill="x", pady=(0, 8))
        ctk.CTkFrame(prev_outer, height=3, corner_radius=12,
                     fg_color=_PRIMARY).pack(fill="x")
        ctk.CTkLabel(prev_outer, text="👁  Expiry Preview",
                     font=("Segoe UI", 12, "bold"),
                     text_color=_TEXT_HI).pack(anchor="w", padx=14, pady=(8, 4))
        ctk.CTkFrame(prev_outer, height=1, fg_color=_BORDER,
                     corner_radius=0).pack(fill="x", padx=14)
        prev_inner = ctk.CTkFrame(prev_outer, fg_color="transparent")
        prev_inner.pack(fill="x", padx=14, pady=10)

        self._preview_lbl = ctk.CTkLabel(prev_inner, text="",
                                          font=("Segoe UI", 16, "bold"),
                                          text_color=_PRIMARY)
        self._preview_lbl.pack()
        self._preview_sub = ctk.CTkLabel(prev_inner, text="",
                                          font=("Segoe UI", 10),
                                          text_color=_TEXT_MID)
        self._preview_sub.pack(pady=(2, 0))

        # Output file box
        out_outer = ctk.CTkFrame(right_col, fg_color=_SURFACE, corner_radius=12,
                                 border_width=1, border_color=_BORDER)
        out_outer.pack(fill="x")
        ctk.CTkFrame(out_outer, height=3, corner_radius=12,
                     fg_color=_GREEN).pack(fill="x")
        ctk.CTkLabel(out_outer, text="📁  Output File",
                     font=("Segoe UI", 12, "bold"),
                     text_color=_TEXT_HI).pack(anchor="w", padx=14, pady=(8, 4))
        ctk.CTkFrame(out_outer, height=1, fg_color=_BORDER,
                     corner_radius=0).pack(fill="x", padx=14)
        out_inner = ctk.CTkFrame(out_outer, fg_color="transparent")
        out_inner.pack(fill="x", padx=14, pady=10)
        self._out_lbl = ctk.CTkLabel(out_inner, text="",
                                      font=("Consolas", 11, "bold"),
                                      text_color=_GREEN, wraplength=220,
                                      justify="left")
        self._out_lbl.pack(anchor="w")
        ctk.CTkLabel(out_inner, text=f"→  {DIST_DIR}",
                     font=("Segoe UI", 9), text_color=_TEXT_MID,
                     wraplength=220, justify="left").pack(anchor="w", pady=(2, 0))

        # ── Build button ──────────────────────────────────────────────────────
        self._build_btn = ctk.CTkButton(
            body,
            text="🚀   Build Trial EXE",
            font=("Segoe UI", 14, "bold"),
            height=44,
            fg_color=_PRIMARY,
            hover_color=_PRI_HOV,
            corner_radius=10,
            command=self._on_build,
        )
        self._build_btn.pack(fill="x", padx=16, pady=(0, 14))

        self._log_append("Ready. Configure expiry above and press  Build Trial EXE.\n")

    # ── Helpers ────────────────────────────────────────────────────────────────
    def _card(self, parent, title: str, pady=(0, 10)) -> ctk.CTkFrame:
        """Titled surface card — packs itself into parent, returns inner content frame."""
        outer = ctk.CTkFrame(parent, fg_color=_SURFACE, corner_radius=14,
                             border_width=1, border_color=_BORDER)
        outer.pack(fill="x", padx=20, pady=pady)
        ctk.CTkFrame(outer, height=4, corner_radius=14, fg_color=_PRIMARY).pack(fill="x")
        ctk.CTkLabel(outer, text=title,
                     font=("Segoe UI", 13, "bold"),
                     text_color=_TEXT_HI).pack(anchor="w", padx=18, pady=(10, 8))
        ctk.CTkFrame(outer, height=1, fg_color=_BORDER, corner_radius=0).pack(fill="x", padx=18)
        inner = ctk.CTkFrame(outer, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=18, pady=12)
        return inner

    def _on_mode_change(self, _=None):
        mode = self._mode_var.get()
        if "Date" in mode:
            self._hours_frame.pack_forget()
            self._date_frame.pack(fill="x")
        else:
            self._date_frame.pack_forget()
            self._hours_frame.pack(fill="x")
        self._refresh_preview()

    def _set_hours(self, h: int):
        self._hours_var.set(str(h))
        self._hmin_var.set("0")
        self._refresh_preview()

    def _get_expiry(self):
        """Parse current inputs → datetime or None on error."""
        try:
            if "Date" in self._mode_var.get():
                return datetime(
                    int(self._year_var.get()),
                    int(self._month_var.get()),
                    int(self._day_var.get()),
                    int(self._hour_var.get()),
                    int(self._min_var.get()),
                )
            else:
                hours   = int(self._hours_var.get())
                minutes = int(self._hmin_var.get())
                return datetime.now() + timedelta(hours=hours, minutes=minutes)
        except Exception:
            return None

    def _refresh_preview(self, *_):
        expiry = self._get_expiry()
        if expiry is None:
            self._preview_lbl.configure(text="Invalid input", text_color=_RED)
            self._preview_sub.configure(text="")
            self._out_lbl.configure(text="")
            return

        now   = datetime.now()
        delta = expiry - now
        hours = int(delta.total_seconds() // 3600)
        days  = delta.days

        if expiry <= now:
            self._preview_lbl.configure(text="⚠  Date is in the past!", text_color=_RED)
            self._preview_sub.configure(text="Choose a future date/time.")
        else:
            self._preview_lbl.configure(
                text=expiry.strftime("%d  %B  %Y   %H:%M"),
                text_color=_PRIMARY)
            if days == 0:
                sub = f"Expires in  {hours} hours"
            elif days == 1:
                sub = "Expires in  1 day"
            else:
                sub = f"Expires in  {days} days  ({hours} hours)"
            self._preview_sub.configure(text=sub)

        tag  = expiry.strftime("%d%b%Y")
        self._out_lbl.configure(text=f"GST_Suite_Trial_{tag}.exe")

    def _log_append(self, text: str):
        """Thread-safe log append."""
        def _do():
            self._log.configure(state="normal")
            self._log.insert("end", text)
            self._log.see("end")
            self._log.configure(state="disabled")
        self.after(0, _do)

    def _log_clear(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _set_building(self, building: bool):
        self._building = building
        self.after(0, lambda: self._build_btn.configure(
            state="disabled" if building else "normal",
            text="⏳  Building…" if building else "🚀   Build Trial EXE",
        ))

    # ══════════════════════════════════════════════════════════════════════════
    #  BUILD LOGIC  (runs in background thread)
    # ══════════════════════════════════════════════════════════════════════════
    def _on_build(self):
        if self._building:
            return

        expiry = self._get_expiry()
        if expiry is None:
            self._log_append("[ERROR] Invalid expiry — fix inputs and try again.\n")
            return
        if expiry <= datetime.now():
            self._log_append("[ERROR] Expiry date is in the past!\n")
            return

        # Check files
        for path, label in [(SOURCE_FILE, "GST_Suite.py"), (SPEC_FILE, "GST_Suite.spec")]:
            if not os.path.exists(path):
                self._log_append(f"[ERROR] {label} not found:\n  {path}\n")
                return

        self._set_building(True)
        self.after(0, self._log_clear)
        thread = threading.Thread(target=self._build_thread, args=(expiry,), daemon=True)
        thread.start()

    def _build_thread(self, expiry: datetime):
        original = None
        success  = False
        try:
            # 1. Patch source
            self._log_append("[ 1 / 4 ]  Patching GST_Suite.py …\n")
            original = self._patch_source(expiry)
            self._log_append(f"          TRIAL_EXPIRY = datetime("
                             f"{expiry.year},{expiry.month},{expiry.day},"
                             f"{expiry.hour},{expiry.minute})\n\n")

            # 2. Run PyInstaller — stream output line by line
            self._log_append("[ 2 / 4 ]  Running PyInstaller …\n")
            self._log_append("─" * 56 + "\n")
            cmd = [sys.executable, "-m", "PyInstaller", SPEC_FILE,
                   "--noconfirm", "--clean"]
            proc = subprocess.Popen(
                cmd,
                cwd=SCRIPT_DIR,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )
            for line in proc.stdout:
                self._log_append(line)
            proc.wait()
            self._log_append("─" * 56 + "\n\n")

            if proc.returncode != 0:
                raise RuntimeError(f"PyInstaller exited with code {proc.returncode}")
            self._log_append("          PyInstaller finished  ✓\n\n")

            # 3. Rename exe
            self._log_append("[ 3 / 4 ]  Renaming output EXE …\n")
            trial_path = self._rename_exe(expiry)
            self._log_append(f"          → {os.path.basename(trial_path)}  ✓\n\n")

            success = True

            # 4. Done
            self._log_append("[ 4 / 4 ]  Restoring GST_Suite.py …\n")

        except Exception as e:
            self._log_append(f"\n[ERROR]  {e}\n")

        finally:
            if original is not None:
                self._restore_source(original)
                self._log_append("          TRIAL_EXPIRY = None  (restored)  ✓\n\n")

            if success:
                tag = expiry.strftime("%d%b%Y")
                self._log_append("═" * 56 + "\n")
                self._log_append(f"  BUILD COMPLETE\n")
                self._log_append(f"  File    :  GST_Suite_Trial_{tag}.exe\n")
                self._log_append(f"  Expires :  {expiry.strftime('%d %b %Y  %H:%M')}\n")
                self._log_append(f"  Folder  :  {DIST_DIR}\n")
                self._log_append("═" * 56 + "\n")
                self.after(0, lambda: self._show_success_popup(expiry))
            else:
                self._log_append("  BUILD FAILED — see errors above.\n")
                self.after(0, self._show_fail_popup)

            self._set_building(False)

    # ── Popup dialogs ──────────────────────────────────────────────────────────
    def _show_success_popup(self, expiry: datetime):
        tag = expiry.strftime("%d%b%Y")
        exe_name = f"GST_Suite_Trial_{tag}.exe"

        popup = ctk.CTkToplevel(self)
        popup.title("Build Complete")
        popup.geometry("460x300")
        popup.resizable(False, False)
        popup.grab_set()          # modal
        popup.lift()
        popup.attributes("-topmost", True)

        # Centre on parent
        self.update_idletasks()
        px = self.winfo_x() + (self.winfo_width()  - 460) // 2
        py = self.winfo_y() + (self.winfo_height() - 300) // 2
        popup.geometry(f"460x300+{px}+{py}")

        # Green accent stripe
        ctk.CTkFrame(popup, height=5, corner_radius=0,
                     fg_color=_GREEN).pack(fill="x")

        # Icon + title
        ctk.CTkLabel(popup, text="✅",
                     font=("Segoe UI Emoji", 44)).pack(pady=(18, 0))
        ctk.CTkLabel(popup, text="Build Complete!",
                     font=("Segoe UI", 18, "bold"),
                     text_color=_GREEN).pack(pady=(4, 0))

        # Details
        details = ctk.CTkFrame(popup, fg_color=_SURFACE, corner_radius=10,
                               border_width=1, border_color=_BORDER)
        details.pack(fill="x", padx=24, pady=(14, 0))

        for label, value in [
            ("File",    exe_name),
            ("Expires", expiry.strftime("%d %b %Y   %H:%M")),
            ("Folder",  DIST_DIR),
        ]:
            row = ctk.CTkFrame(details, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=3)
            ctk.CTkLabel(row, text=f"{label}:",
                         font=("Segoe UI", 11),
                         text_color=_TEXT_MID, width=56,
                         anchor="w").pack(side="left")
            ctk.CTkLabel(row, text=value,
                         font=("Segoe UI", 11, "bold"),
                         text_color=_TEXT_HI,
                         anchor="w").pack(side="left")

        ctk.CTkButton(popup, text="OK",
                      width=120, height=36,
                      font=("Segoe UI", 13, "bold"),
                      fg_color=_GREEN,
                      hover_color=("#047857", "#059669"),
                      command=popup.destroy).pack(pady=18)

    def _show_fail_popup(self):
        popup = ctk.CTkToplevel(self)
        popup.title("Build Failed")
        popup.geometry("400x200")
        popup.resizable(False, False)
        popup.grab_set()
        popup.lift()
        popup.attributes("-topmost", True)

        self.update_idletasks()
        px = self.winfo_x() + (self.winfo_width()  - 400) // 2
        py = self.winfo_y() + (self.winfo_height() - 200) // 2
        popup.geometry(f"400x200+{px}+{py}")

        ctk.CTkFrame(popup, height=5, corner_radius=0,
                     fg_color=_RED).pack(fill="x")
        ctk.CTkLabel(popup, text="❌",
                     font=("Segoe UI Emoji", 40)).pack(pady=(18, 0))
        ctk.CTkLabel(popup, text="Build Failed",
                     font=("Segoe UI", 17, "bold"),
                     text_color=_RED).pack(pady=(4, 0))
        ctk.CTkLabel(popup, text="Check the Build Log for details.",
                     font=("Segoe UI", 12),
                     text_color=_TEXT_MID).pack(pady=(4, 0))
        ctk.CTkButton(popup, text="OK",
                      width=110, height=34,
                      font=("Segoe UI", 12, "bold"),
                      fg_color=_RED,
                      hover_color=("#b91c1c", "#dc2626"),
                      command=popup.destroy).pack(pady=16)

    # ── Backend helpers (same logic as original CLI version) ──────────────────
    def _patch_source(self, expiry: datetime) -> str:
        with open(SOURCE_FILE, "r", encoding="utf-8") as f:
            original = f.read()
        if not TRIAL_PATTERN.search(original):
            raise RuntimeError(
                "Could not find 'TRIAL_EXPIRY = None' in GST_Suite.py.\n"
                "Make sure the source file has not been modified manually."
            )
        replacement = (
            f"TRIAL_EXPIRY = datetime("
            f"{expiry.year}, {expiry.month}, {expiry.day}, "
            f"{expiry.hour}, {expiry.minute})"
        )
        patched = TRIAL_PATTERN.sub(replacement, original)
        with open(SOURCE_FILE, "w", encoding="utf-8") as f:
            f.write(patched)
        return original

    def _restore_source(self, original: str):
        with open(SOURCE_FILE, "w", encoding="utf-8") as f:
            f.write(original)

    def _rename_exe(self, expiry: datetime) -> str:
        if not os.path.exists(ORIGINAL_EXE):
            raise FileNotFoundError(f"Output EXE not found: {ORIGINAL_EXE}")
        tag        = expiry.strftime("%d%b%Y")
        trial_name = f"GST_Suite_Trial_{tag}.exe"
        trial_path = os.path.join(DIST_DIR, trial_name)
        if os.path.exists(trial_path):
            os.remove(trial_path)
        shutil.move(ORIGINAL_EXE, trial_path)
        return trial_path


# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = BuildTrialApp()
    app.mainloop()
