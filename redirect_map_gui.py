#!/usr/bin/env python3
"""
redirect_map_gui.py

A simple point-and-click front end for sitemap_to_redirect_map.py and
match_new_site.py. Instead of typing commands in a terminal, you type the
old-site and new-site URLs into two boxes and click buttons.

SETUP
-----
Put this file in the SAME folder as sitemap_to_redirect_map.py and
match_new_site.py (the "301-automate" folder from the repo). It doesn't
replace those two scripts or the one-time setup in the README (Python,
pip install requests beautifulsoup4 lxml openpyxl) -- it just runs them
for you and shows the output in a window.

RUNNING IT
----------
Double-click this file, OR open a terminal in that folder and run:

    python redirect_map_gui.py

TIP: To skip opening a terminal at all, right-click this file -> "Create
shortcut", then right-click the shortcut -> Properties -> change "Target"
so it starts with:  pythonw.exe  (instead of python.exe)
That launches the window with no black console box behind it. Then pin
that shortcut to your taskbar/desktop for one-click access.

OPTIONAL EXTRA POLISH
----------------------
This looks good out of the box with no extra installs. If you want it to
match Windows 11's native look even more closely, you can optionally run:

    pip install sv-ttk

and this file will automatically pick it up next time it starts -- nothing
else to configure. Not required.
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
STAGE1_SCRIPT = os.path.join(SCRIPT_DIR, "sitemap_to_redirect_map.py")
STAGE2_SCRIPT = os.path.join(SCRIPT_DIR, "match_new_site.py")

# ---------------------------------------------------------------------- #
# Color palette / fonts -- change these to retheme the whole app.
# ---------------------------------------------------------------------- #
BG = "#F4F5F8"            # window background
CARD_BG = "#FFFFFF"       # section ("card") background
BORDER = "#E3E5EA"
TEXT_PRIMARY = "#1F2430"
TEXT_SECONDARY = "#6B7280"
ACCENT = "#4F46E5"        # indigo
ACCENT_HOVER = "#4338CA"
ACCENT_DISABLED = "#C7C9F4"
ACCENT_SOFT = "#EEF0FF"
SUCCESS = "#16A34A"
ERROR = "#DC2626"
LOG_BG = "#111827"
LOG_FG = "#E5E7EB"
LOG_MUTED = "#9CA3AF"

FONT_FAMILY = "Segoe UI" if sys.platform.startswith("win") else "Helvetica"
MONO_FAMILY = "Consolas" if sys.platform.startswith("win") else "Courier New"

FONT_HEADER = (FONT_FAMILY, 18, "bold")
FONT_SUB = (FONT_FAMILY, 10)
FONT_SECTION = (FONT_FAMILY, 11, "bold")
FONT_BODY = (FONT_FAMILY, 10)
FONT_MUTED = (FONT_FAMILY, 9)
FONT_BUTTON = (FONT_FAMILY, 10, "bold")
FONT_MONO = (MONO_FAMILY, 10)


def domain_to_filename(url: str) -> str:
    """Turn https://www.example.com/ into Example301RW.xlsx as a friendly default."""
    if not url:
        return "redirect_map.xlsx"
    cleaned = url.strip()
    for prefix in ("https://", "http://"):
        if cleaned.startswith(prefix):
            cleaned = cleaned[len(prefix):]
            break
    cleaned = cleaned.split("/")[0]
    if cleaned.startswith("www."):
        cleaned = cleaned[4:]
    core = cleaned.split(".")[0]
    core = "".join(ch for ch in core if ch.isalnum())
    if not core:
        return "redirect_map.xlsx"
    return f"{core[0].upper()}{core[1:]}301RW.xlsx"


class RedirectMapApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Redirect Map Builder")
        self.minsize(700, 600)
        self.configure(bg=BG)

        self._log_queue: "queue.Queue[str]" = queue.Queue()
        self._running = False

        self._setup_style()
        self._check_scripts_present()
        self._build_ui()
        self._size_to_content()
        self._center_window()
        self.after(100, self._drain_log_queue)

    # ------------------------------------------------------------------ #
    # Look & feel
    # ------------------------------------------------------------------ #
    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=FONT_BODY)

        # Plain (window-background) surfaces
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT_PRIMARY, font=FONT_BODY)
        style.configure("Header.TLabel", background=BG, foreground=TEXT_PRIMARY, font=FONT_HEADER)
        style.configure("Sub.TLabel", background=BG, foreground=TEXT_SECONDARY, font=FONT_SUB)

        # Card (white section) surfaces
        style.configure("Card.TFrame", background=CARD_BG)
        style.configure("Card.TLabel", background=CARD_BG, foreground=TEXT_PRIMARY, font=FONT_BODY)
        style.configure("Muted.TLabel", background=CARD_BG, foreground=TEXT_SECONDARY, font=FONT_MUTED)
        style.configure("Badge.TLabel", background=CARD_BG, foreground=TEXT_PRIMARY, font=FONT_SECTION)

        style.configure(
            "TLabelframe", background=CARD_BG, bordercolor=BORDER,
            borderwidth=1, relief="solid",
        )
        style.configure("TLabelframe.Label", background=CARD_BG)

        style.configure(
            "TEntry", fieldbackground="#FFFFFF", foreground=TEXT_PRIMARY,
            bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
            borderwidth=1, relief="solid", padding=6,
        )
        style.map("TEntry", bordercolor=[("focus", ACCENT)])

        style.configure("TCheckbutton", background=CARD_BG, foreground=TEXT_PRIMARY, font=FONT_MUTED)
        style.map("TCheckbutton", background=[("active", CARD_BG)])

        # Primary action button
        style.configure(
            "Accent.TButton", font=FONT_BUTTON, foreground="#FFFFFF",
            background=ACCENT, borderwidth=0, focusthickness=0, padding=(14, 11),
        )
        style.map(
            "Accent.TButton",
            background=[("disabled", ACCENT_DISABLED), ("active", ACCENT_HOVER)],
            foreground=[("disabled", "#FFFFFF")],
        )

        # Secondary / utility button
        style.configure(
            "Secondary.TButton", font=FONT_MUTED, foreground=TEXT_PRIMARY,
            background="#EEF0F4", borderwidth=0, focusthickness=0, padding=(10, 7),
        )
        style.map("Secondary.TButton", background=[("active", "#E2E5EC")])

        style.configure(
            "Horizontal.TProgressbar", troughcolor="#E5E7EB",
            background=ACCENT, bordercolor="#E5E7EB", lightcolor=ACCENT, darkcolor=ACCENT,
        )

        # Optional extra polish if the user has installed it -- safe no-op otherwise.
        try:
            import sv_ttk
            sv_ttk.set_theme("light")
        except ImportError:
            pass

    def _size_to_content(self):
        """Open at whatever size actually fits everything (log panel included),
        capped so it never opens taller than the screen."""
        self.update_idletasks()
        req_w = max(self.winfo_reqwidth(), 820)
        req_h = max(self.winfo_reqheight(), 600)
        max_h = self.winfo_screenheight() - 100
        self.geometry(f"{req_w}x{min(req_h, max_h)}")

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _step_badge(self, parent, num, text):
        """Small numbered circle + bold title, used as a LabelFrame's labelwidget."""
        frame = ttk.Frame(parent, style="Card.TFrame")
        canvas = tk.Canvas(frame, width=24, height=24, bg=CARD_BG, highlightthickness=0)
        canvas.create_oval(2, 2, 22, 22, fill=ACCENT, outline="")
        canvas.create_text(12, 12, text=str(num), fill="#FFFFFF", font=(FONT_FAMILY, 10, "bold"))
        canvas.pack(side="left", padx=(2, 8), pady=4)
        ttk.Label(frame, text=text, style="Badge.TLabel").pack(side="left")
        return frame

    def _make_collapsible(self, parent, title):
        """
        A click-to-expand "Advanced options ▸" disclosure. Returns the body
        frame -- pack/grid your widgets into it. Collapsed by default so the
        two steps stay compact; nothing inside is lost, just tucked away.
        """
        container = ttk.Frame(parent, style="Card.TFrame")
        toggle_state = {"open": False}

        header = ttk.Frame(container, style="Card.TFrame")
        header.pack(fill="x")
        arrow_lbl = ttk.Label(header, text="▸ " + title, style="Muted.TLabel", cursor="hand2")
        arrow_lbl.pack(side="left")

        body = ttk.Frame(container, style="Card.TFrame")

        def toggle(_event=None):
            if toggle_state["open"]:
                body.pack_forget()
                arrow_lbl.configure(text="▸ " + title)
            else:
                body.pack(fill="x", pady=(8, 0))
                arrow_lbl.configure(text="▾ " + title)
            toggle_state["open"] = not toggle_state["open"]

        arrow_lbl.bind("<Button-1>", toggle)
        return container, body

    # ------------------------------------------------------------------ #
    # Setup checks
    # ------------------------------------------------------------------ #
    def _check_scripts_present(self):
        missing = [
            name for name, path in (
                ("sitemap_to_redirect_map.py", STAGE1_SCRIPT),
                ("match_new_site.py", STAGE2_SCRIPT),
            ) if not os.path.isfile(path)
        ]
        if missing:
            messagebox.showwarning(
                "Scripts not found",
                "This GUI needs to sit in the same folder as:\n\n"
                + "\n".join(missing)
                + "\n\nCurrent folder:\n" + SCRIPT_DIR
                + "\n\nMove this file next to them, or move them next to this file.",
            )

    # ------------------------------------------------------------------ #
    # UI construction
    # ------------------------------------------------------------------ #
    def _build_ui(self):
        outer = ttk.Frame(self, style="TFrame")
        outer.pack(fill="both", expand=True, padx=18, pady=16)

        # ---------------- Header ----------------
        header = ttk.Frame(outer, style="TFrame")
        header.pack(fill="x", pady=(0, 14))
        ttk.Label(header, text="Redirect Map Builder", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Build a 301 redirect map from an old site, then match it against the new one.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        # Two-column layout: Step 1 and Step 2 side by side, so the log
        # panel below still has real room to breathe on laptop-sized screens.
        content = ttk.Frame(outer, style="TFrame")
        content.pack(fill="x")
        content.columnconfigure(0, weight=1, uniform="steps")
        content.columnconfigure(1, weight=1, uniform="steps")

        # ---------------- Step 1: Old site ----------------
        step1 = ttk.LabelFrame(content, padding=(4, 8, 4, 12))
        step1.configure(labelwidget=self._step_badge(step1, 1, "Build the map from the OLD site"))
        step1.grid(row=0, column=0, sticky="nsew", padx=(0, 7))
        step1.columnconfigure(1, weight=1)

        ttk.Label(step1, text="Old site URL", style="Card.TLabel").grid(
            row=0, column=0, sticky="w", padx=10, pady=(4, 6)
        )
        self.old_url_var = tk.StringVar()
        old_entry = ttk.Entry(step1, textvariable=self.old_url_var)
        old_entry.grid(row=1, column=0, columnspan=3, sticky="we", padx=10, pady=(0, 10))
        old_entry.bind("<KeyRelease>", self._maybe_autofill_output)

        ttk.Label(step1, text="Output file", style="Card.TLabel").grid(
            row=2, column=0, sticky="w", padx=10, pady=(0, 6)
        )
        self.output_var = tk.StringVar(value="redirect_map.xlsx")
        ttk.Entry(step1, textvariable=self.output_var).grid(
            row=3, column=0, columnspan=2, sticky="we", padx=(10, 6), pady=(0, 10)
        )
        ttk.Button(step1, text="Browse…", style="Secondary.TButton", command=self._browse_output).grid(
            row=3, column=2, sticky="e", padx=(0, 10), pady=(0, 10)
        )
        self._output_autofilled = True  # tracks whether user has hand-edited it

        adv_container1, adv1 = self._make_collapsible(step1, "Advanced options")
        adv_container1.grid(row=4, column=0, columnspan=3, sticky="we", padx=10, pady=(2, 10))

        row1 = ttk.Frame(adv1, style="Card.TFrame")
        row1.pack(fill="x", pady=(0, 6))
        ttk.Label(row1, text="Max pages to crawl", style="Muted.TLabel").pack(side="left")
        self.max_pages_var = tk.StringVar(value="300")
        ttk.Entry(row1, textvariable=self.max_pages_var, width=6).pack(side="left", padx=(6, 0))

        self.no_crawl_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(adv1, text="Don't fall back to crawling", variable=self.no_crawl_var).pack(
            fill="x", pady=(0, 2)
        )
        self.no_rest_api_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(adv1, text="Skip WordPress REST API check", variable=self.no_rest_api_var).pack(
            fill="x", pady=(0, 6)
        )

        ttk.Label(
            adv1, text="Extra URLs (optional, one per line — orphaned pages)", style="Muted.TLabel"
        ).pack(fill="x", pady=(0, 4))
        self.extra_urls_text = tk.Text(
            adv1, height=3, wrap="word", font=FONT_BODY,
            bg="#FFFFFF", fg=TEXT_PRIMARY, insertbackground=TEXT_PRIMARY,
            relief="flat", borderwidth=0, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT, padx=6, pady=4,
        )
        self.extra_urls_text.pack(fill="x")

        self.run_stage1_btn = ttk.Button(
            step1, text="Build Redirect Map from Old Site", style="Accent.TButton", command=self._run_stage1
        )
        self.run_stage1_btn.grid(row=5, column=0, columnspan=3, sticky="we", padx=10, pady=(0, 6))

        # ---------------- Step 2: New site ----------------
        step2 = ttk.LabelFrame(content, padding=(4, 8, 4, 12))
        step2.configure(labelwidget=self._step_badge(step2, 2, "Match in the NEW site"))
        step2.grid(row=0, column=1, sticky="nsew", padx=(7, 0))
        step2.columnconfigure(1, weight=1)

        ttk.Label(step2, text="New site URL", style="Card.TLabel").grid(
            row=0, column=0, sticky="w", padx=10, pady=(4, 6)
        )
        self.new_url_var = tk.StringVar()
        ttk.Entry(step2, textvariable=self.new_url_var).grid(
            row=1, column=0, columnspan=3, sticky="we", padx=10, pady=(0, 10)
        )

        ttk.Label(step2, text="Workbook to update", style="Card.TLabel").grid(
            row=2, column=0, sticky="w", padx=10, pady=(0, 6)
        )
        ttk.Entry(step2, textvariable=self.output_var).grid(
            row=3, column=0, columnspan=2, sticky="we", padx=(10, 6), pady=(0, 4)
        )
        ttk.Button(
            step2, text="Browse…", style="Secondary.TButton", command=self._browse_workbook_for_stage2
        ).grid(row=3, column=2, sticky="e", padx=(0, 10), pady=(0, 4))
        ttk.Label(
            step2, text="Same file as above — this reuses whatever's in the Output file box",
            style="Muted.TLabel",
        ).grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 8))

        adv_container2, adv2 = self._make_collapsible(step2, "Advanced options")
        adv_container2.grid(row=5, column=0, columnspan=3, sticky="we", padx=10, pady=(2, 10))

        row2 = ttk.Frame(adv2, style="Card.TFrame")
        row2.pack(fill="x")
        ttk.Label(row2, text="Max pages to crawl", style="Muted.TLabel").pack(side="left")
        self.max_pages_var2 = tk.StringVar(value="300")
        ttk.Entry(row2, textvariable=self.max_pages_var2, width=6).pack(side="left", padx=(6, 18))
        self.no_crawl_var2 = tk.BooleanVar(value=False)
        ttk.Checkbutton(row2, text="Don't fall back to crawling", variable=self.no_crawl_var2).pack(side="left")

        self.run_stage2_btn = ttk.Button(
            step2, text="Match New Site into Workbook", style="Accent.TButton", command=self._run_stage2
        )
        self.run_stage2_btn.grid(row=6, column=0, columnspan=3, sticky="we", padx=10, pady=(0, 6))

        # ---------------- Status row ----------------
        status_row = ttk.Frame(outer, style="TFrame")
        status_row.pack(fill="x", pady=(14, 6))

        self.status_dot = tk.Canvas(status_row, width=12, height=12, bg=BG, highlightthickness=0)
        self._status_dot_id = self.status_dot.create_oval(1, 1, 11, 11, fill=TEXT_SECONDARY, outline="")
        self.status_dot.pack(side="left", padx=(0, 8))

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(status_row, textvariable=self.status_var, style="TLabel").pack(side="left")

        ttk.Button(
            status_row, text="Open Output Folder", style="Secondary.TButton", command=self._open_output_folder
        ).pack(side="right")

        self.progress = ttk.Progressbar(status_row, mode="indeterminate", length=140)
        self.progress.pack(side="right", padx=(0, 14))

        # ---------------- Log panel ----------------
        log_frame = ttk.LabelFrame(outer, padding=(4, 6, 4, 8))
        log_frame.configure(labelwidget=self._step_badge(log_frame, "≡", "Progress log"))
        log_frame.pack(fill="both", expand=True, pady=(4, 0))

        log_toolbar = ttk.Frame(log_frame, style="Card.TFrame")
        log_toolbar.pack(fill="x", padx=8, pady=(0, 4))
        ttk.Button(
            log_toolbar, text="Clear", style="Secondary.TButton", command=self._clear_log
        ).pack(side="right")

        self.log_widget = scrolledtext.ScrolledText(
            log_frame, state="disabled", wrap="word", height=9,
            bg=LOG_BG, fg=LOG_FG, insertbackground=LOG_FG, font=FONT_MONO,
            relief="flat", borderwidth=0, padx=10, pady=8,
        )
        self.log_widget.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.log_widget.tag_configure("success", foreground=SUCCESS, font=(MONO_FAMILY, 10, "bold"))
        self.log_widget.tag_configure("error", foreground=ERROR, font=(MONO_FAMILY, 10, "bold"))
        self.log_widget.tag_configure("banner", foreground="#93C5FD", font=(MONO_FAMILY, 10, "bold"))
        self.log_widget.tag_configure("muted", foreground=LOG_MUTED)

    # ------------------------------------------------------------------ #
    # Small helpers
    # ------------------------------------------------------------------ #
    def _maybe_autofill_output(self, _event=None):
        if self._output_autofilled:
            self.output_var.set(domain_to_filename(self.old_url_var.get()))

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=self.output_var.get() or "redirect_map.xlsx",
        )
        if path:
            self._output_autofilled = False
            self.output_var.set(path)

    def _browse_workbook_for_stage2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self._output_autofilled = False
            self.output_var.set(path)

    def _open_output_folder(self):
        out = self.output_var.get().strip() or "redirect_map.xlsx"
        folder = os.path.dirname(os.path.abspath(out)) or SCRIPT_DIR
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception as e:
            messagebox.showerror("Couldn't open folder", str(e))

    def _clear_log(self):
        self.log_widget.configure(state="normal")
        self.log_widget.delete("1.0", "end")
        self.log_widget.configure(state="disabled")

    def _log(self, text: str):
        self._log_queue.put(text)

    def _log_tag_for(self, line: str):
        if line.lstrip().startswith("✔"):
            return "success"
        if line.lstrip().startswith("✘"):
            return "error"
        if "Can't save" in line or "CLOSE the workbook" in line or "PermissionError" in line:
            return "error"
        if "====" in line:
            return "banner"
        if line.strip().startswith(("Locating sitemap", "Using sitemap", "Falling back", "Skipping", "Found ")):
            return "muted"
        return None

    def _drain_log_queue(self):
        try:
            while True:
                line = self._log_queue.get_nowait()
                self.log_widget.configure(state="normal")
                tag = self._log_tag_for(line)
                if tag:
                    self.log_widget.insert("end", line, tag)
                else:
                    self.log_widget.insert("end", line)
                self.log_widget.see("end")
                self.log_widget.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._drain_log_queue)

    def _set_status_dot(self, color: str):
        self.status_dot.itemconfig(self._status_dot_id, fill=color)

    def _set_running(self, running: bool, status: str = ""):
        self._running = running
        state = "disabled" if running else "normal"
        self.run_stage1_btn.configure(state=state)
        self.run_stage2_btn.configure(state=state)
        if running:
            self._set_status_dot(ACCENT)
            self.progress.start(12)
        else:
            self.progress.stop()
        if status:
            self.status_var.set(status)

    # ------------------------------------------------------------------ #
    # Validation
    # ------------------------------------------------------------------ #
    @staticmethod
    def _looks_like_url(url: str) -> bool:
        url = url.strip()
        return url.startswith("http://") or url.startswith("https://")

    # ------------------------------------------------------------------ #
    # Stage 1: build from old site
    # ------------------------------------------------------------------ #
    def _run_stage1(self):
        if self._running:
            return
        old_url = self.old_url_var.get().strip()
        output = self.output_var.get().strip() or "redirect_map.xlsx"

        if not self._looks_like_url(old_url):
            messagebox.showerror("Missing/invalid URL", "Enter the old site's URL, e.g. https://www.oldsite.com/")
            return
        if not os.path.isfile(STAGE1_SCRIPT):
            messagebox.showerror("Script missing", f"Can't find {STAGE1_SCRIPT}")
            return

        args = [sys.executable, STAGE1_SCRIPT, old_url, "-o", output]

        max_pages = self.max_pages_var.get().strip()
        if max_pages:
            args += ["--max-pages", max_pages]
        if self.no_crawl_var.get():
            args.append("--no-crawl")
        if self.no_rest_api_var.get():
            args.append("--no-rest-api")

        extra_urls = [
            line.strip() for line in self.extra_urls_text.get("1.0", "end").splitlines() if line.strip()
        ]
        if extra_urls:
            args += ["--extra-urls", *extra_urls]

        self._log(f"\n{'=' * 60}\nStep 1: building redirect map from {old_url}\n{'=' * 60}\n")
        self._run_in_background(args, done_status="Step 1 finished.")

    # ------------------------------------------------------------------ #
    # Stage 2: match new site
    # ------------------------------------------------------------------ #
    def _run_stage2(self):
        if self._running:
            return
        new_url = self.new_url_var.get().strip()
        workbook = self.output_var.get().strip()

        if not self._looks_like_url(new_url):
            messagebox.showerror("Missing/invalid URL", "Enter the new site's URL, e.g. https://www.newsite.com/")
            return
        if not workbook or not os.path.isfile(workbook):
            messagebox.showerror(
                "Workbook not found",
                f"Can't find the workbook:\n{workbook}\n\nRun Step 1 first, or Browse to an existing .xlsx.",
            )
            return
        if not os.path.isfile(STAGE2_SCRIPT):
            messagebox.showerror("Script missing", f"Can't find {STAGE2_SCRIPT}")
            return

        args = [sys.executable, STAGE2_SCRIPT, workbook, new_url]

        max_pages = self.max_pages_var2.get().strip()
        if max_pages:
            args += ["--max-pages", max_pages]
        if self.no_crawl_var2.get():
            args.append("--no-crawl")

        self._log(f"\n{'=' * 60}\nStep 2: matching {new_url} into {workbook}\n{'=' * 60}\n")
        self._run_in_background(args, done_status="Step 2 finished.")

    # ------------------------------------------------------------------ #
    # Subprocess runner
    # ------------------------------------------------------------------ #
    def _run_in_background(self, args, done_status: str):
        self._set_running(True, "Running…")

        def worker():
            try:
                process = subprocess.Popen(
                    args,
                    cwd=SCRIPT_DIR,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                )
                for line in process.stdout:
                    self._log(line)
                process.wait()
                if process.returncode == 0:
                    self._log("\n✔ Done.\n")
                    self.status_var.set(done_status)
                    self._set_status_dot(SUCCESS)
                else:
                    self._log(f"\n✘ Exited with an error (code {process.returncode}). Scroll up for details.\n")
                    self.status_var.set("Finished with errors — see log.")
                    self._set_status_dot(ERROR)
            except FileNotFoundError as e:
                self._log(f"\n✘ Couldn't start: {e}\n")
                self.status_var.set("Failed to start.")
                self._set_status_dot(ERROR)
            except Exception as e:  # noqa: BLE001
                self._log(f"\n✘ Unexpected error: {e}\n")
                self.status_var.set("Failed.")
                self._set_status_dot(ERROR)
            finally:
                self._set_running(False)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = RedirectMapApp()
    app.mainloop()