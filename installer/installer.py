"""
Meeting Recorder — Bulletproof Windows Installer
Handles all edge cases for mixed GPU, corporate, and consumer setups.
"""

import os
import sys
import json
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser
from pathlib import Path
import ctypes
import winreg
import shutil

APP_FILES    = {}   # populated by bundle.py
APP_ICON_B64 = ""

APP_NAME    = "Meeting Recorder"
APP_VERSION = "1.0.0"
PY_URL      = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
MIN_DISK_GB = 5.0

BG          = "#1a1c1e"
PANEL       = "#2d2f31"
INPUT       = "#3b3d3f"
ACCENT      = "#4fc3f7"
ACCENT2     = "#0277bd"
ACCENT_BG   = "#003a57"
TEXT        = "#e6e1e5"
MUTED       = "#938f99"
HINT        = "#6b6870"
BORDER      = "#4a4d50"
WARN_BG     = "#3a2a00"
WARN_FG     = "#ffcc80"
OK_BG       = "#1a3a1a"
OK_FG       = "#a5d6a7"
SUCCESS_DIM = "#a5d6a7"

FONT_TITLE = ("Segoe UI", 19, "bold")
FONT_BODY  = ("Segoe UI", 11)
FONT_SMALL = ("Segoe UI", 10)
FONT_MONO  = ("Cascadia Code", 10)

REQUIREMENTS_BASE = [
    "python-dotenv",
    "sounddevice",
    "soundfile",
    "scipy",
    "faster-whisper",
    "matplotlib",
    "huggingface_hub==0.23.0",
    "pyannote.audio==3.3.2",
    "pywin32",
    "anthropic",
]


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def get_free_disk_gb(path):
    try:
        total, used, free = shutil.disk_usage(path)
        return free / (1024 ** 3)
    except Exception:
        return 999.0


def has_nvidia_gpu():
    try:
        r = subprocess.run(["nvidia-smi"], capture_output=True, timeout=5)
        return r.returncode == 0
    except Exception:
        return False


def find_python311():
    candidates = ["python3.11", "python", "python3", "py"]
    for cmd in candidates:
        try:
            r = subprocess.run([cmd, "--version"], capture_output=True, text=True, timeout=5)
            if r.returncode == 0:
                ver = r.stdout.strip().split()[1]
                major, minor = int(ver.split(".")[0]), int(ver.split(".")[1])
                if major == 3 and minor == 11:
                    return cmd, ver
        except Exception:
            continue
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Python\PythonCore\3.11\InstallPath") as k:
            path = winreg.QueryValue(k, None)
            exe = Path(path) / "python.exe"
            if exe.exists():
                return str(exe), "3.11"
    except Exception:
        pass
    return None, None


def test_anthropic_key(key):
    try:
        import urllib.request
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/models",
            headers={"x-api-key": key, "anthropic-version": "2023-06-01"})
        with urllib.request.urlopen(req, timeout=10):
            return True, "API key is valid"
    except Exception as e:
        err = str(e)
        if "401" in err:
            return False, "Invalid API key — check you copied it correctly"
        if "403" in err:
            return False, "API key has no credits — add billing at console.anthropic.com"
        if "URLError" in err or "timeout" in err.lower():
            return False, "Cannot reach Anthropic — check internet connection"
        return False, f"Validation failed: {err[:100]}"


def test_hf_token(token):
    try:
        import urllib.request
        req = urllib.request.Request(
            "https://huggingface.co/api/whoami",
            headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read())
            name = data.get("name", "unknown")
            return True, f"Token valid — logged in as {name}"
    except Exception as e:
        err = str(e)
        if "401" in err:
            return False, "Invalid token — check you copied it correctly"
        return False, f"Validation failed: {err[:100]}"


def check_pyannote_access(token):
    import urllib.request
    for model_id in ["pyannote/speaker-diarization-3.1", "pyannote/segmentation-3.0"]:
        try:
            req = urllib.request.Request(
                f"https://huggingface.co/api/models/{model_id}",
                headers={"Authorization": f"Bearer {token}"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                pass
        except Exception as e:
            if "403" in str(e):
                return False, f"Terms not accepted for {model_id}"
    return True, "Model access confirmed"


def create_vbs(install_dir):
    vbs = (f'Set WshShell = CreateObject("WScript.Shell")\n'
           f'WshShell.Run "cmd /c cd {install_dir} && '
           f'call .venv\\Scripts\\activate && python main.py", 0, False\n')
    path = Path(install_dir) / "launch.vbs"
    path.write_text(vbs)
    return str(path)


def create_shortcut(install_dir, vbs_path):
    desktop = Path(os.path.expanduser("~")) / "Desktop"
    lnk = desktop / f"{APP_NAME}.lnk"
    ico = Path(install_dir) / "meeting_recorder.ico"
    ps = (f'$ws = New-Object -ComObject WScript.Shell\n'
          f'$s = $ws.CreateShortcut("{lnk}")\n'
          f'$s.TargetPath = "{vbs_path}"\n'
          f'$s.WorkingDirectory = "{install_dir}"\n'
          f'$s.Description = "{APP_NAME}"\n'
          f'$s.WindowStyle = 1\n')
    if ico.exists():
        ps += f'$s.IconLocation = "{ico}"\n'
    ps += "$s.Save()"
    subprocess.run(["powershell", "-Command", ps], capture_output=True)


class Installer(tk.Tk):

    STEPS = ["Welcome", "System Check", "Location", "Python 3.11",
             "Anthropic Key", "HuggingFace", "Model Terms", "Installing", "Done"]

    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} Installer v{APP_VERSION}")
        self.geometry("740x580")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._install_dir   = tk.StringVar(value=str(Path.home() / "meeting_recorder"))
        self._anthropic_key = tk.StringVar()
        self._hf_token      = tk.StringVar()
        self._python_cmd    = None
        self._python_ver    = None
        self._has_nvidia    = False
        self._step          = 0

        self._build_chrome()
        self._show_step(0)

    def _center(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - 740) // 2
        y = (self.winfo_screenheight() - 580) // 2
        self.geometry(f"740x580+{x}+{y}")

    def _build_chrome(self):
        sidebar = tk.Frame(self, bg=PANEL, width=190)
        sidebar.pack(side=tk.LEFT, fill=tk.Y)
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text="🎙", bg=PANEL, fg=ACCENT,
                 font=("Segoe UI", 26)).pack(pady=(24, 4))
        tk.Label(sidebar, text=APP_NAME, bg=PANEL, fg=TEXT,
                 font=("Segoe UI", 10, "bold"), wraplength=170).pack()
        tk.Label(sidebar, text=f"v{APP_VERSION}", bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 9)).pack(pady=(2, 20))
        self._step_labels = []
        for name in self.STEPS:
            frm = tk.Frame(sidebar, bg=PANEL)
            frm.pack(fill=tk.X, padx=10, pady=2)
            dot = tk.Label(frm, text="●", bg=PANEL, fg=HINT, font=("Segoe UI", 8))
            dot.pack(side=tk.LEFT, padx=(0, 6))
            lbl = tk.Label(frm, text=name, bg=PANEL, fg=HINT,
                           font=("Segoe UI", 9), anchor="w")
            lbl.pack(side=tk.LEFT)
            self._step_labels.append((dot, lbl))

        self._content = tk.Frame(self, bg=BG)
        self._content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._bottom = tk.Frame(self._content, bg=PANEL, height=58)
        self._bottom.pack(side=tk.BOTTOM, fill=tk.X)
        self._bottom.pack_propagate(False)
        self._back_btn = tk.Button(
            self._bottom, text="← Back", bg=PANEL, fg=MUTED,
            font=FONT_BODY, relief=tk.FLAT, padx=14, pady=8,
            cursor="hand2", command=self._back)
        self._back_btn.pack(side=tk.LEFT, padx=14, pady=10)
        self._next_btn = tk.Button(
            self._bottom, text="Next →", bg=ACCENT2, fg=TEXT,
            font=FONT_BODY, relief=tk.FLAT, padx=18, pady=8,
            cursor="hand2", command=self._next)
        self._next_btn.pack(side=tk.RIGHT, padx=14, pady=10)
        self._page = tk.Frame(self._content, bg=BG)
        self._page.pack(fill=tk.BOTH, expand=True, padx=24, pady=18)

    def _update_sidebar(self, idx):
        for i, (dot, lbl) in enumerate(self._step_labels):
            if i < idx:
                dot.config(fg=ACCENT2); lbl.config(fg=MUTED)
            elif i == idx:
                dot.config(fg=ACCENT); lbl.config(fg=TEXT)
            else:
                dot.config(fg=HINT); lbl.config(fg=HINT)

    def _clear_page(self):
        for w in self._page.winfo_children():
            w.destroy()

    def _show_step(self, step):
        self._step = step
        self._update_sidebar(step)
        self._clear_page()
        [self._pg_welcome, self._pg_syscheck, self._pg_location,
         self._pg_python, self._pg_anthropic, self._pg_huggingface,
         self._pg_pyannote, self._pg_install, self._pg_done][step]()

    def _next(self):
        s = self._step
        if s == 0:
            self._show_step(1)
        elif s == 1:
            self._show_step(2)
        elif s == 2:
            path = self._install_dir.get().strip()
            if not path:
                messagebox.showwarning("Required", "Please choose a location.")
                return
            free = get_free_disk_gb(Path(path).anchor or "C:\\")
            if free < MIN_DISK_GB:
                messagebox.showerror("Not Enough Space",
                    f"Need {MIN_DISK_GB:.0f} GB free, only {free:.1f} GB available.")
                return
            self._show_step(3)
        elif s == 3:
            cmd, ver = find_python311()
            if not cmd:
                messagebox.showwarning("Python 3.11 Required",
                    "Python 3.11 not found.\nInstall it using the link on this page, then click Next.")
                self._show_step(3)
                return
            self._python_cmd = cmd
            self._python_ver = ver
            self._show_step(4)
        elif s == 4:
            key = self._anthropic_key.get().strip()
            if not key:
                messagebox.showwarning("Required", "Anthropic API key is required.")
                return
            self._next_btn.config(state=tk.DISABLED, text="Validating...")
            def _v():
                ok, msg = test_anthropic_key(key)
                self.after(0, lambda: self._next_btn.config(state=tk.NORMAL, text="Validate & Continue →"))
                if ok:
                    self.after(0, lambda: self._show_step(5))
                else:
                    self.after(0, lambda: messagebox.showerror("API Key Error", msg))
            threading.Thread(target=_v, daemon=True).start()
        elif s == 5:
            token = self._hf_token.get().strip()
            if not token:
                messagebox.showwarning("Required", "HuggingFace token is required.")
                return
            self._next_btn.config(state=tk.DISABLED, text="Validating...")
            def _v():
                ok, msg = test_hf_token(token)
                self.after(0, lambda: self._next_btn.config(state=tk.NORMAL, text="Validate & Continue →"))
                if ok:
                    self.after(0, lambda: self._show_step(6))
                else:
                    self.after(0, lambda: messagebox.showerror("Token Error", msg))
            threading.Thread(target=_v, daemon=True).start()
        elif s == 6:
            token = self._hf_token.get().strip()
            self._next_btn.config(state=tk.DISABLED, text="Checking model access...")
            def _v():
                ok, msg = check_pyannote_access(token)
                self.after(0, lambda: self._next_btn.config(
                    state=tk.NORMAL, text="I've Accepted Both →"))
                if ok:
                    self.after(0, lambda: self._start_install())
                else:
                    self.after(0, lambda: messagebox.showerror("Model Terms Not Accepted",
                        f"{msg}\n\nPlease accept the terms for both models before continuing."))
            threading.Thread(target=_v, daemon=True).start()
        elif s == 8:
            self.destroy()

    def _start_install(self):
        self._show_step(7)
        threading.Thread(target=self._run_install, daemon=True).start()

    def _back(self):
        if 0 < self._step < 7:
            self._show_step(self._step - 1)

    # ── Page helpers ──────────────────────────────────────────────────

    def _heading(self, text):
        tk.Label(self._page, text=text, bg=BG, fg=TEXT,
                 font=FONT_TITLE, anchor="w").pack(anchor="w", pady=(0, 6))

    def _sub(self, text):
        tk.Label(self._page, text=text, bg=BG, fg=MUTED, font=FONT_BODY,
                 anchor="w", justify="left", wraplength=500
                 ).pack(anchor="w", pady=(0, 10))

    def _card(self, bg=PANEL, pady=10, padx=12):
        f = tk.Frame(self._page, bg=bg, pady=pady, padx=padx)
        f.pack(fill=tk.X, pady=(0, 8))
        return f

    def _guide_card(self, title, steps):
        c = self._card(bg=ACCENT_BG)
        tk.Label(c, text=title, bg=ACCENT_BG, fg=ACCENT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        for s in steps:
            tk.Label(c, text=s, bg=ACCENT_BG, fg=TEXT,
                     font=FONT_SMALL, anchor="w", justify="left"
                     ).pack(anchor="w", pady=1)

    def _link_btn(self, text, url, pady=(0, 10)):
        tk.Button(self._page, text=text, bg=ACCENT2, fg=TEXT,
                  font=FONT_BODY, relief=tk.FLAT, padx=14, pady=6,
                  cursor="hand2", command=lambda: webbrowser.open(url)
                  ).pack(anchor="w", pady=pady)

    def _key_field(self, var):
        row = tk.Frame(self._page, bg=BG)
        row.pack(fill=tk.X, pady=(0, 6))
        entry = tk.Entry(row, textvariable=var, bg=INPUT, fg=TEXT,
                         font=FONT_MONO, insertbackground=TEXT,
                         relief=tk.FLAT, highlightbackground=BORDER,
                         highlightthickness=1, show="•")
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(0, 8))
        tk.Button(row, text="Show/Hide", bg=PANEL, fg=MUTED,
                  font=FONT_SMALL, relief=tk.FLAT, cursor="hand2",
                  command=lambda e=entry: e.config(
                      show="" if e.cget("show") == "•" else "•")
                  ).pack(side=tk.LEFT)

    # ── Pages ─────────────────────────────────────────────────────────

    def _pg_welcome(self):
        self._back_btn.config(state=tk.DISABLED)
        self._next_btn.config(text="Get Started →", state=tk.NORMAL)
        self._heading(f"Welcome to {APP_NAME}")
        self._sub("This wizard sets up everything you need, step by step.")
        for icon, title, desc in [
            ("🔍", "System Check",   "Verifies your PC meets requirements"),
            ("🐍", "Python 3.11",    "Installs the right Python version"),
            ("📦", "AI Libraries",   "PyTorch, Whisper, pyannote, Anthropic SDK"),
            ("🔑", "API Keys",       "Guided setup with live key validation"),
            ("✅", "Model Terms",    "Step-by-step HuggingFace model acceptance"),
            ("🖥",  "Desktop Icon",  "Creates a shortcut to launch the app"),
        ]:
            row = tk.Frame(self._page, bg=PANEL, pady=7, padx=12)
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=icon, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 14)).pack(side=tk.LEFT, padx=(0, 10))
            col = tk.Frame(row, bg=PANEL)
            col.pack(side=tk.LEFT)
            tk.Label(col, text=title, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 10, "bold"), anchor="w").pack(anchor="w")
            tk.Label(col, text=desc, bg=PANEL, fg=MUTED,
                     font=("Segoe UI", 9), anchor="w").pack(anchor="w")
        tk.Label(self._page,
                 text="Total time: 15–25 minutes depending on internet speed.",
                 bg=BG, fg=HINT, font=FONT_SMALL).pack(anchor="w", pady=(10, 0))

    def _pg_syscheck(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="Next →", state=tk.NORMAL)
        self._heading("System Check")
        self._sub("Checking your PC before we begin...")

        self._has_nvidia = has_nvidia_gpu()
        free_gb = get_free_disk_gb(Path(self._install_dir.get()).anchor or "C:\\")
        disk_ok = free_gb >= MIN_DISK_GB
        admin   = is_admin()

        try:
            import urllib.request
            urllib.request.urlopen("https://pypi.org", timeout=5)
            net_ok = True
        except Exception:
            net_ok = False

        checks = [
            ("Administrator rights", admin,
             "Running as administrator ✓" if admin
             else "Not admin — if install fails, right-click and 'Run as administrator'"),
            ("Disk space", disk_ok,
             f"{free_gb:.1f} GB free ✓" if disk_ok
             else f"Only {free_gb:.1f} GB free — need {MIN_DISK_GB:.0f} GB. Free up space first."),
            ("GPU", True,
             "Nvidia GPU detected — fast GPU processing will be used ✓"
             if self._has_nvidia
             else "No Nvidia GPU — CPU processing will be used (slower but works fine)"),
            ("Internet", net_ok,
             "Internet connection available ✓" if net_ok
             else "Cannot reach pypi.org — check internet or contact IT about firewall"),
        ]

        for name, ok, msg in checks:
            bg = OK_BG if ok else WARN_BG
            fg = OK_FG if ok else WARN_FG
            row = tk.Frame(self._page, bg=bg, pady=8, padx=12)
            row.pack(fill=tk.X, pady=3)
            tk.Label(row, text="✓" if ok else "⚠", bg=bg, fg=fg,
                     font=("Segoe UI", 12, "bold"), width=2).pack(side=tk.LEFT)
            col = tk.Frame(row, bg=bg)
            col.pack(side=tk.LEFT, padx=(8, 0))
            tk.Label(col, text=name, bg=bg, fg=fg,
                     font=("Segoe UI", 10, "bold"), anchor="w").pack(anchor="w")
            tk.Label(col, text=msg, bg=bg, fg=fg,
                     font=("Segoe UI", 9), anchor="w",
                     wraplength=460, justify="left").pack(anchor="w")

        if not net_ok:
            self._next_btn.config(state=tk.DISABLED)
            tk.Label(self._page,
                     text="⚠ Internet required. Contact IT if on a corporate network.",
                     bg=BG, fg=WARN_FG, font=FONT_SMALL).pack(anchor="w", pady=(6, 0))

    def _pg_location(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="Next →", state=tk.NORMAL)
        self._heading("Install Location")
        self._sub("Choose where Meeting Recorder will be installed.")
        row = tk.Frame(self._page, bg=BG)
        row.pack(fill=tk.X, pady=(0, 10))
        entry = tk.Entry(row, textvariable=self._install_dir, bg=INPUT, fg=TEXT,
                         font=FONT_BODY, insertbackground=TEXT, relief=tk.FLAT,
                         highlightbackground=BORDER, highlightthickness=1)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(0, 8))
        tk.Button(row, text="Browse", bg=PANEL, fg=TEXT, font=FONT_BODY,
                  relief=tk.FLAT, padx=12, pady=6, cursor="hand2",
                  command=lambda: self._install_dir.set(
                      filedialog.askdirectory(title="Choose Install Folder").replace("/", "\\"))
                  ).pack(side=tk.LEFT)
        c = self._card(bg=ACCENT_BG)
        tk.Label(c, text="What gets installed here:", bg=ACCENT_BG, fg=ACCENT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        for item in ["• Python virtual environment", "• AI libraries (~3–4 GB)",
                     "• Application files", "• API key config (stays on your machine)"]:
            tk.Label(c, text=item, bg=ACCENT_BG, fg=TEXT,
                     font=FONT_SMALL).pack(anchor="w")
        tk.Label(self._page, text=f"Required: {MIN_DISK_GB:.0f} GB free",
                 bg=BG, fg=HINT, font=FONT_SMALL).pack(anchor="w", pady=(4, 0))

    def _pg_python(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="Check Again →")
        self._heading("Python 3.11")
        cmd, ver = find_python311()
        if cmd:
            self._python_cmd = cmd
            self._python_ver = ver
            self._next_btn.config(text="Next →")
            c = tk.Frame(self._page, bg=OK_BG, pady=12, padx=14)
            c.pack(fill=tk.X, pady=(0, 10))
            tk.Label(c, text=f"✓  Python {ver} found — ready!",
                     bg=OK_BG, fg=OK_FG, font=("Segoe UI", 11, "bold")).pack(anchor="w")
            tk.Label(c, text="Click Next to continue.",
                     bg=OK_BG, fg=OK_FG, font=FONT_SMALL).pack(anchor="w")
        else:
            c = tk.Frame(self._page, bg=WARN_BG, pady=10, padx=14)
            c.pack(fill=tk.X, pady=(0, 10))
            tk.Label(c, text="⚠  Python 3.11 not found",
                     bg=WARN_BG, fg=WARN_FG,
                     font=("Segoe UI", 11, "bold")).pack(anchor="w")
            tk.Label(c, text="Meeting Recorder requires Python 3.11 specifically.",
                     bg=WARN_BG, fg=WARN_FG, font=FONT_SMALL).pack(anchor="w")
            self._guide_card("Install Python 3.11 (follow these steps exactly):", [
                "1. Click 'Download Python 3.11.9' below",
                "2. Run the downloaded file",
                "3. ⚠ IMPORTANT: Check 'Add Python to PATH' at the bottom of the installer!",
                "4. Click 'Install Now' and wait",
                "5. Click 'Close' when done",
                "6. Come back here and click 'Check Again'",
            ])
            tk.Button(self._page, text="⬇  Download Python 3.11.9",
                      bg=ACCENT2, fg=TEXT, font=FONT_BODY, relief=tk.FLAT,
                      padx=14, pady=8, cursor="hand2",
                      command=lambda: webbrowser.open(PY_URL)
                      ).pack(anchor="w")

    def _pg_anthropic(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="Validate & Continue →", state=tk.NORMAL)
        self._heading("Anthropic API Key")
        self._sub("Powers AI meeting summarization.\nSaved only on your computer — never shared.")
        self._guide_card("Get your key in 5 minutes:", [
            "1. Click 'Open Anthropic Console' below",
            "2. Sign up for a free account (or log in)",
            "3. Click 'API Keys' in the left sidebar",
            "4. Click '+ Create Key', name it 'Meeting Recorder'",
            "5. Copy the key shown (starts with sk-ant-...)",
            "6. Click 'Billing' → 'Add Credits' — add at least $5",
            "7. Paste your key in the box below",
        ])
        self._link_btn("🌐  Open Anthropic Console",
                       "https://console.anthropic.com/settings/keys")
        tk.Label(self._page, text="Paste your Anthropic API key:",
                 bg=BG, fg=TEXT, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        self._key_field(self._anthropic_key)
        tk.Label(self._page,
                 text="~$0.04 per summary  |  $5 ≈ 125 summaries  |  Key will be validated before install",
                 bg=BG, fg=HINT, font=FONT_SMALL).pack(anchor="w")

    def _pg_huggingface(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="Validate & Continue →", state=tk.NORMAL)
        self._heading("HuggingFace Token")
        self._sub("Required for speaker detection (who said what).\n100% free — no payment needed.")
        self._guide_card("Get your free token in 3 minutes:", [
            "1. Click 'Open HuggingFace' below",
            "2. Sign up for a free account (or log in)",
            "3. Click your profile picture → Settings",
            "4. Click 'Access Tokens' in the left sidebar",
            "5. Click 'New token'",
            "6. Name it 'Meeting Recorder', Type = 'Read'",
            "7. Click 'Generate a token'",
            "8. Copy the token (starts with hf_...)",
            "9. Paste it in the box below",
        ])
        self._link_btn("🌐  Open HuggingFace", "https://huggingface.co/settings/tokens")
        tk.Label(self._page, text="Paste your HuggingFace token:",
                 bg=BG, fg=TEXT, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        self._key_field(self._hf_token)
        tk.Label(self._page, text="Token will be validated before install",
                 bg=BG, fg=HINT, font=FONT_SMALL).pack(anchor="w")

    def _pg_pyannote(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="I've Accepted Both →", state=tk.NORMAL)
        self._heading("Accept AI Model Terms")
        self._sub("Two AI models require you to accept their terms on HuggingFace.\nThis is a one-time step.")

        warn = tk.Frame(self._page, bg=WARN_BG, pady=8, padx=12)
        warn.pack(fill=tk.X, pady=(0, 8))
        tk.Label(warn, text="⚠  Make sure you are logged in to HuggingFace first!",
                 bg=WARN_BG, fg=WARN_FG,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")

        for num, title, url, steps in [
            ("1", "Speaker Diarization Model",
             "https://huggingface.co/pyannote/speaker-diarization-3.1",
             ["1. Click 'Open Model 1' below",
              "2. Scroll down to the agreement box",
              "3. Fill in your info and click 'Agree and access repository'",
              "4. You'll see model files — this means it worked ✓"]),
            ("2", "Segmentation Model",
             "https://huggingface.co/pyannote/segmentation-3.0",
             ["1. Click 'Open Model 2' below",
              "2. Same process — agree and access repository",
              "3. You'll see model files — this means it worked ✓"]),
        ]:
            c = self._card(bg=PANEL)
            tk.Label(c, text=f"Step {num} of 2 — {title}",
                     bg=PANEL, fg=ACCENT,
                     font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
            for s in steps:
                tk.Label(c, text=s, bg=PANEL, fg=TEXT,
                         font=FONT_SMALL, anchor="w").pack(anchor="w", pady=1)
            tk.Button(c, text=f"🌐  Open Model {num}", bg=ACCENT2, fg=TEXT,
                      font=FONT_SMALL, relief=tk.FLAT, padx=12, pady=4,
                      cursor="hand2", command=lambda u=url: webbrowser.open(u)
                      ).pack(anchor="w", pady=(6, 0))

        tk.Label(self._page,
                 text="After accepting both, click 'I've Accepted Both' — the installer will verify automatically.",
                 bg=BG, fg=HINT, font=FONT_SMALL, wraplength=500).pack(anchor="w", pady=(6, 0))

    def _pg_install(self):
        self._back_btn.config(state=tk.DISABLED)
        self._next_btn.config(state=tk.DISABLED)
        self._heading("Installing...")
        self._prog_label = tk.Label(self._page, text="Preparing...",
                                     bg=BG, fg=MUTED, font=FONT_BODY, anchor="w")
        self._prog_label.pack(anchor="w", pady=(0, 8))
        style = ttk.Style()
        style.theme_use("default")
        style.configure("M.Horizontal.TProgressbar",
                         troughcolor=PANEL, background=ACCENT,
                         borderwidth=0, thickness=8)
        self._progress = ttk.Progressbar(
            self._page, style="M.Horizontal.TProgressbar",
            length=660, mode="determinate")
        self._progress.pack(anchor="w", pady=(0, 12))
        self._log = tk.Text(self._page, bg=PANEL, fg=MUTED,
                            font=("Segoe UI", 9), relief=tk.FLAT,
                            state=tk.DISABLED, height=16, wrap=tk.WORD,
                            highlightthickness=0)
        self._log.pack(fill=tk.BOTH, expand=True)

    def _pg_done(self):
        self._back_btn.config(state=tk.DISABLED)
        self._next_btn.config(text="Finish", state=tk.NORMAL)
        self._heading("All Done! 🎉")
        self._sub("Meeting Recorder is installed and ready to use.")
        for icon, label in [
            ("✓", "Virtual environment created"),
            ("✓", f"PyTorch installed ({'GPU accelerated' if self._has_nvidia else 'CPU mode'})"),
            ("✓", "All AI libraries installed"),
            ("✓", "API keys saved"),
            ("✓", "HuggingFace login complete"),
            ("✓", "Desktop shortcut created"),
        ]:
            row = tk.Frame(self._page, bg=BG)
            row.pack(anchor="w", pady=2)
            tk.Label(row, text=icon, bg=BG, fg=ACCENT,
                     font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=(0, 10))
            tk.Label(row, text=label, bg=BG, fg=TEXT, font=FONT_BODY).pack(side=tk.LEFT)
        note = tk.Frame(self._page, bg=ACCENT_BG, pady=10, padx=14)
        note.pack(fill=tk.X, pady=(14, 0))
        for line in [
            "ℹ  First launch downloads AI models (~700MB) — takes 2-3 min, one time only.",
            "ℹ  If Windows shows a security warning, click 'More info' then 'Run anyway'.",
            "ℹ  A README.txt with quick start instructions is in your install folder.",
        ]:
            tk.Label(note, text=line, bg=ACCENT_BG, fg=ACCENT,
                     font=("Segoe UI", 9), anchor="w", justify="left"
                     ).pack(anchor="w", pady=2)
        tk.Button(self._page, text="🚀  Launch Meeting Recorder Now",
                  bg=ACCENT2, fg=TEXT, font=("Segoe UI", 12, "bold"),
                  relief=tk.FLAT, padx=20, pady=10, cursor="hand2",
                  command=lambda: [
                      os.startfile(str(Path(self._install_dir.get()) / "launch.vbs"))
                      if (Path(self._install_dir.get()) / "launch.vbs").exists() else None,
                      self.destroy()
                  ]).pack(anchor="w", pady=(16, 0))

    # ── Install logic ──────────────────────────────────────────────────

    def _log_write(self, msg):
        def _do():
            self._log.config(state=tk.NORMAL)
            self._log.insert(tk.END, msg + "\n")
            self._log.see(tk.END)
            self._log.config(state=tk.DISABLED)
        self.after(0, _do)

    def _set_progress(self, value, label=""):
        def _do():
            self._progress["value"] = value
            if label:
                self._prog_label.config(text=label)
        self.after(0, _do)

    def _run_install(self):
        try:
            install_dir = Path(self._install_dir.get())
            install_dir.mkdir(parents=True, exist_ok=True)
            python = self._python_cmd or "python"

            self._set_progress(3, "Creating virtual environment...")
            self._log_write("→ Creating virtual environment...")
            r = subprocess.run([python, "-m", "venv", str(install_dir / ".venv")],
                               capture_output=True, text=True)
            if r.returncode != 0:
                raise RuntimeError(
                    "Failed to create virtual environment.\n"
                    "Try running the installer as Administrator.\n"
                    f"Error: {r.stderr[:200]}")
            self._log_write("  ✓ Virtual environment created")

            pip = str(install_dir / ".venv" / "Scripts" / "pip.exe")
            py  = str(install_dir / ".venv" / "Scripts" / "python.exe")

            self._set_progress(6, "Upgrading pip...")
            subprocess.run([py, "-m", "pip", "install", "--upgrade", "pip"],
                           capture_output=True)
            self._log_write("  ✓ pip upgraded")

            # PyTorch — GPU or CPU
            if self._has_nvidia:
                self._set_progress(8, "Installing PyTorch with GPU support (large download, please wait)...")
                self._log_write("→ Installing PyTorch CUDA (GPU)...")
                r = subprocess.run(
                    [pip, "install", "torch==2.6.0", "torchaudio==2.6.0",
                     "--index-url", "https://download.pytorch.org/whl/cu124"],
                    capture_output=True, text=True)
                if r.returncode != 0:
                    self._log_write("  ⚠ CUDA failed, falling back to CPU...")
                    r = subprocess.run(
                        [pip, "install", "torch==2.6.0", "torchaudio==2.6.0",
                         "--index-url", "https://download.pytorch.org/whl/cpu"],
                        capture_output=True, text=True)
                    if r.returncode != 0:
                        raise RuntimeError(
                            "PyTorch install failed.\nCheck internet connection.\n" + r.stderr[:200])
                    self._log_write("  ✓ PyTorch CPU installed")
                else:
                    self._log_write("  ✓ PyTorch CUDA installed")
            else:
                self._set_progress(8, "Installing PyTorch CPU (large download, please wait)...")
                self._log_write("→ Installing PyTorch CPU...")
                r = subprocess.run(
                    [pip, "install", "torch==2.6.0", "torchaudio==2.6.0",
                     "--index-url", "https://download.pytorch.org/whl/cpu"],
                    capture_output=True, text=True)
                if r.returncode != 0:
                    raise RuntimeError(
                        "PyTorch install failed.\nCheck internet connection.\n" + r.stderr[:200])
                self._log_write("  ✓ PyTorch CPU installed")

            self._set_progress(32)

            # PyAudio with fallback
            self._set_progress(33, "Installing PyAudio...")
            self._log_write("→ Installing PyAudio...")
            r = subprocess.run([pip, "install", "pyaudio"], capture_output=True, text=True)
            if r.returncode != 0:
                self._log_write("  ⚠ PyAudio failed — trying pipwin fallback...")
                subprocess.run([pip, "install", "pipwin"], capture_output=True)
                r2 = subprocess.run([py, "-m", "pipwin", "install", "pyaudio"],
                                    capture_output=True, text=True)
                if r2.returncode != 0:
                    self._log_write("  ⚠ PyAudio could not be installed automatically.")
                    self._log_write("    Audio recording may not work.")
                    self._log_write("    Fix: install Microsoft C++ Build Tools from microsoft.com")
                else:
                    self._log_write("  ✓ PyAudio installed via pipwin")
            else:
                self._log_write("  ✓ PyAudio installed")

            # Core requirements
            total = len(REQUIREMENTS_BASE)
            for i, pkg in enumerate(REQUIREMENTS_BASE):
                pct  = 36 + int((i / total) * 38)
                name = pkg.split("==")[0]
                self._set_progress(pct, f"Installing {name}...")
                self._log_write(f"→ Installing {name}...")
                r = subprocess.run([pip, "install", pkg], capture_output=True, text=True)
                if r.returncode != 0:
                    stderr_low = r.stderr.lower()
                    if any(w in stderr_low for w in ["firewall", "proxy", "ssl", "certificate"]):
                        raise RuntimeError(
                            f"Network blocked installing {name}.\n"
                            "Your corporate firewall may be blocking pip.\n"
                            "Ask IT to whitelist pypi.org and files.pythonhosted.org")
                    self._log_write(f"  ⚠ {name} failed: {r.stderr[:80]}")
                else:
                    self._log_write(f"  ✓ {name}")

            # Pin numpy
            self._set_progress(76, "Pinning numpy version...")
            self._log_write("→ Pinning numpy...")
            subprocess.run([pip, "install", "numpy==2.1.3"], capture_output=True)
            self._log_write("  ✓ numpy==2.1.3")

            # Write app files
            self._set_progress(80, "Writing application files...")
            self._log_write("→ Writing application files...")
            self._write_app_files(install_dir)
            self._log_write("  ✓ Application files written")

            # Write .env
            self._set_progress(85, "Saving configuration...")
            self._log_write("→ Saving configuration...")
            (install_dir / ".env").write_text(
                f"ANTHROPIC_API_KEY={self._anthropic_key.get().strip()}\n"
                f"HF_TOKEN={self._hf_token.get().strip()}\n"
                f"WHISPER_MODEL=base\n"
                f"MAX_SPEAKERS=8\n"
                f"RECORDINGS_DIR=recordings\n",
                encoding="utf-8")
            self._log_write("  ✓ Configuration saved")

            # HuggingFace login
            self._set_progress(88, "Logging in to HuggingFace...")
            self._log_write("→ HuggingFace login...")
            subprocess.run(
                [py, "-c",
                 f'from huggingface_hub import login; login(token="{self._hf_token.get().strip()}")'],
                capture_output=True)
            self._log_write("  ✓ HuggingFace login complete")

            # Patch main.py
            self._set_progress(91, "Applying compatibility patches...")
            self._log_write("→ Applying patches...")
            self._patch_main(install_dir)
            self._log_write("  ✓ Patches applied")

            # Shortcut
            self._set_progress(95, "Creating desktop shortcut...")
            self._log_write("→ Creating desktop shortcut...")
            vbs = create_vbs(str(install_dir))
            create_shortcut(str(install_dir), vbs)
            self._log_write("  ✓ Desktop shortcut created")

            # README
            self._set_progress(98, "Writing quick start guide...")
            self._write_readme(install_dir)
            self._log_write("  ✓ README.txt written")

            self._set_progress(100, "Done!")
            self._log_write("\n" + "=" * 48)
            self._log_write("  INSTALLATION COMPLETE!")
            self._log_write("=" * 48)
            self._log_write("Double-click the desktop shortcut to launch.")
            self.after(1500, lambda: self._show_step(8))

        except Exception as e:
            self._log_write(f"\n{'='*48}\n  INSTALLATION FAILED\n{'='*48}")
            self._log_write(f"Error: {e}")
            self._log_write("\nCommon fixes:")
            self._log_write("  • Right-click installer → Run as Administrator")
            self._log_write("  • Check internet connection")
            self._log_write("  • Ask IT to whitelist pypi.org if on corporate network")
            self._set_progress(self._progress["value"], "Failed — see log above")
            self.after(0, lambda: messagebox.showerror("Failed",
                f"{str(e)[:300]}\n\nSee the log for details."))
            self.after(0, lambda: self._next_btn.config(
                state=tk.NORMAL, text="Retry",
                command=lambda: threading.Thread(
                    target=self._run_install, daemon=True).start()))

    def _write_app_files(self, install_dir: Path):
        for rel_path, content in APP_FILES.items():
            dest = install_dir / rel_path
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_text(content, encoding="utf-8")
        if APP_ICON_B64:
            import base64
            (install_dir / "meeting_recorder.ico").write_bytes(
                base64.b64decode(APP_ICON_B64))

    def _patch_main(self, install_dir: Path):
        main_path = install_dir / "main.py"
        if not main_path.exists():
            return
        content = main_path.read_text(encoding="utf-8")
        if "np.NaN = np.nan" in content:
            return
        patch = '''import torch
import numpy as np
from torch.torch_version import TorchVersion

if not hasattr(np, "NaN"):
    np.NaN = np.nan
if not hasattr(np, "NAN"):
    np.NAN = np.nan
torch.serialization.add_safe_globals([TorchVersion])
_orig_load = torch.load
def _patched_load(f, *args, **kwargs):
    kwargs["weights_only"] = False
    return _orig_load(f, *args, **kwargs)
torch.load = _patched_load

'''
        main_path.write_text(patch + content, encoding="utf-8")

    def _write_readme(self, install_dir: Path):
        (install_dir / "README.txt").write_text(
            f"""{APP_NAME} — Quick Start Guide
{'=' * 44}

LAUNCHING
  Double-click the desktop shortcut.
  If Windows warns "unknown publisher": click More info → Run anyway.

FIRST LAUNCH
  Downloads AI models (~700MB). Takes 2-5 min. Happens once only.

HOW TO USE
  1. Select your microphone
  2. Click Start Recording before your meeting
  3. Click Stop Recording when done
  4. Click Process & Transcribe — wait for completion
  5. Rename speakers if needed
  6. Click Summarize for an AI summary
  7. Click Email Summary to send to your Outlook inbox

RECORDINGS
  Saved in: {install_dir}\\recordings\\
  Click Open Recordings in the app to browse.

TROUBLESHOOTING
  App won't start   → Run as Administrator
  No audio          → Check microphone in Windows Settings → Privacy → Microphone
  Processing fails  → Verify HuggingFace model terms were accepted:
                      huggingface.co/pyannote/speaker-diarization-3.1
                      huggingface.co/pyannote/segmentation-3.0
  Summary fails     → Check Anthropic billing at console.anthropic.com
  Email fails       → Make sure Outlook desktop is open

Install location: {install_dir}
""", encoding="utf-8")

    def _on_close(self):
        if self._step == 7:
            if not messagebox.askyesno("Cancel?", "Installation in progress. Cancel?"):
                return
        self.destroy()


if __name__ == "__main__":
    app = Installer()
    app.mainloop()
