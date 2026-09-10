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
        self.geometry("780x680")
        self.minsize(680, 560)

        self._log_queue: "queue.Queue[str]" = queue.Queue()
        self._running = False

        self._check_scripts_present()
        self._build_ui()
        self.after(100, self._drain_log_queue)

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
        pad = {"padx": 10, "pady": 6}

        notebook_frame = ttk.Frame(self)
        notebook_frame.pack(fill="x", **pad)

        # ---------------- Step 1: Old site ----------------
        step1 = ttk.LabelFrame(notebook_frame, text="Step 1 — Build the map from the OLD site")
        step1.pack(fill="x", pady=(0, 10))

        ttk.Label(step1, text="Old site URL:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.old_url_var = tk.StringVar()
        old_entry = ttk.Entry(step1, textvariable=self.old_url_var, width=55)
        old_entry.grid(row=0, column=1, columnspan=2, sticky="we", padx=8, pady=6)
        old_entry.bind("<KeyRelease>", self._maybe_autofill_output)

        ttk.Label(step1, text="Output file:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        self.output_var = tk.StringVar(value="redirect_map.xlsx")
        ttk.Entry(step1, textvariable=self.output_var, width=40).grid(
            row=1, column=1, sticky="we", padx=8, pady=6
        )
        ttk.Button(step1, text="Browse…", command=self._browse_output).grid(
            row=1, column=2, sticky="w", padx=8, pady=6
        )
        self._output_autofilled = True  # tracks whether user has hand-edited it

        adv1 = ttk.Frame(step1)
        adv1.grid(row=2, column=0, columnspan=3, sticky="we", padx=8, pady=(0, 4))
        ttk.Label(adv1, text="Max pages to crawl:").pack(side="left")
        self.max_pages_var = tk.StringVar(value="300")
        ttk.Entry(adv1, textvariable=self.max_pages_var, width=6).pack(side="left", padx=(4, 16))
        self.no_crawl_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(adv1, text="Don't fall back to crawling", variable=self.no_crawl_var).pack(side="left")
        self.no_rest_api_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(adv1, text="Skip WordPress REST API check", variable=self.no_rest_api_var).pack(
            side="left", padx=(16, 0)
        )

        ttk.Label(step1, text="Extra URLs (optional, one per line — orphaned pages):").grid(
            row=3, column=0, columnspan=3, sticky="w", padx=8, pady=(4, 0)
        )
        self.extra_urls_text = tk.Text(step1, height=3, width=60)
        self.extra_urls_text.grid(row=4, column=0, columnspan=3, sticky="we", padx=8, pady=(0, 8))

        self.run_stage1_btn = ttk.Button(
            step1, text="Build Redirect Map from Old Site", command=self._run_stage1
        )
        self.run_stage1_btn.grid(row=5, column=0, columnspan=3, sticky="we", padx=8, pady=(0, 10))

        step1.columnconfigure(1, weight=1)

        # ---------------- Step 2: New site ----------------
        step2 = ttk.LabelFrame(notebook_frame, text="Step 2 — Match in the NEW site")
        step2.pack(fill="x")

        ttk.Label(step2, text="New site URL:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.new_url_var = tk.StringVar()
        ttk.Entry(step2, textvariable=self.new_url_var, width=55).grid(
            row=0, column=1, columnspan=2, sticky="we", padx=8, pady=6
        )

        ttk.Label(step2, text="Workbook to update:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(step2, textvariable=self.output_var, width=40).grid(
            row=1, column=1, sticky="we", padx=8, pady=6
        )
        ttk.Button(step2, text="Browse…", command=self._browse_workbook_for_stage2).grid(
            row=1, column=2, sticky="w", padx=8, pady=6
        )
        ttk.Label(
            step2, text="(Same file as above — this reuses whatever's in the Output file box)",
            foreground="#666",
        ).grid(row=2, column=0, columnspan=3, sticky="w", padx=8)

        adv2 = ttk.Frame(step2)
        adv2.grid(row=3, column=0, columnspan=3, sticky="we", padx=8, pady=(8, 4))
        ttk.Label(adv2, text="Max pages to crawl:").pack(side="left")
        self.max_pages_var2 = tk.StringVar(value="300")
        ttk.Entry(adv2, textvariable=self.max_pages_var2, width=6).pack(side="left", padx=(4, 16))
        self.no_crawl_var2 = tk.BooleanVar(value=False)
        ttk.Checkbutton(adv2, text="Don't fall back to crawling", variable=self.no_crawl_var2).pack(side="left")

        self.run_stage2_btn = ttk.Button(
            step2, text="Match New Site into Workbook", command=self._run_stage2
        )
        self.run_stage2_btn.grid(row=4, column=0, columnspan=3, sticky="we", padx=8, pady=(4, 10))

        step2.columnconfigure(1, weight=1)

        # ---------------- Bottom bar ----------------
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=(0, 4))
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")
        ttk.Button(bottom, text="Open Output Folder", command=self._open_output_folder).pack(side="right")

        # ---------------- Log panel ----------------
        log_frame = ttk.LabelFrame(self, text="Progress log")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.log_widget = scrolledtext.ScrolledText(log_frame, state="disabled", wrap="word", height=14)
        self.log_widget.pack(fill="both", expand=True, padx=4, pady=4)

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

    def _log(self, text: str):
        self._log_queue.put(text)

    def _drain_log_queue(self):
        try:
            while True:
                line = self._log_queue.get_nowait()
                self.log_widget.configure(state="normal")
                self.log_widget.insert("end", line)
                self.log_widget.see("end")
                self.log_widget.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._drain_log_queue)

    def _set_running(self, running: bool, status: str = ""):
        self._running = running
        state = "disabled" if running else "normal"
        self.run_stage1_btn.configure(state=state)
        self.run_stage2_btn.configure(state=state)
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
                else:
                    self._log(f"\n✘ Exited with an error (code {process.returncode}). Scroll up for details.\n")
                    self.status_var.set("Finished with errors — see log.")
            except FileNotFoundError as e:
                self._log(f"\n✘ Couldn't start: {e}\n")
                self.status_var.set("Failed to start.")
            except Exception as e:  # noqa: BLE001
                self._log(f"\n✘ Unexpected error: {e}\n")
                self.status_var.set("Failed.")
            finally:
                self._set_running(False)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = RedirectMapApp()
    app.mainloop()
