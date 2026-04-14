"""
Meeting Recorder — Windows Installer Wizard
Bulletproof setup with full API key walkthroughs and model terms verification.
v1.1 — Fixed HuggingFace gated-model access detection
"""

import os
import sys
import json
import shutil
import subprocess
import threading
import datetime
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser
from pathlib import Path
import ctypes
import winreg

APP_FILES    = {}   # populated by bundle.py
APP_ICON_B64 = ""

APP_NAME    = "Meeting Recorder"
APP_VERSION = "1.3.1"
PY_URL      = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
MIN_DISK_GB = 5.0
LOG_PATH    = Path(os.path.expanduser("~")) / "MeetingRecorderInstall.log"

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
DANGER      = "#cf6679"

FONT_TITLE = ("Segoe UI", 19, "bold")
FONT_H2    = ("Segoe UI", 13, "bold")
FONT_BODY  = ("Segoe UI", 11)
FONT_SMALL = ("Segoe UI", 10)
FONT_MONO  = ("Cascadia Code", 10)

# NOTE: openai-whisper is intentionally excluded — it fails to build on Python 3.13+
# We use faster-whisper instead: pre-built, faster, uses less memory.
REQUIREMENTS_BASE = [
    "anthropic",
    "python-dotenv",
    "sounddevice",
    "soundfile",
    "scipy",
    "matplotlib",
    "faster-whisper",
    "huggingface_hub==0.23.0",
    "pyannote.audio==3.3.2",
    "pywin32",
]

# Must be upgraded FIRST before any other pip install to avoid pkg_resources errors
BOOTSTRAP_PACKAGES = ["pip", "setuptools", "wheel"]


# ── Logging ────────────────────────────────────────────────────────────────────

def _log_start():
    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write(f"Meeting Recorder Install Log\n")
            f.write(f"Date: {datetime.datetime.now()}\n")
            f.write(f"Log: {LOG_PATH}\n")
            f.write("=" * 50 + "\n\n")
    except Exception:
        pass


def _log(msg):
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}\n")
    except Exception:
        pass


# ── System checks ──────────────────────────────────────────────────────────────

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


def get_gpu_info() -> dict:
    """
    Returns dict with:
      - has_gpu (bool)
      - name (str)       e.g. "NVIDIA GeForce RTX 5080"
      - is_blackwell (bool)  RTX 50xx — needs PyTorch 2.7+ / cu128
      - cuda_version (str)   e.g. "12.8"
      - torch_version (str)  recommended torch version string
      - torch_index (str)    PyPI index URL for torch install
    """
    info = {
        "has_gpu":       False,
        "name":          "Unknown",
        "is_blackwell":  False,
        "cuda_version":  "12.4",
        "torch_version": "2.6.0",
        "torch_index":   "https://download.pytorch.org/whl/cu124",
    }
    try:
        r = subprocess.run(
            ["nvidia-smi",
             "--query-gpu=name,driver_version",
             "--format=csv,noheader"],
            capture_output=True, text=True, timeout=8)
        if r.returncode != 0:
            return info
        line = r.stdout.strip().split("\n")[0]
        name = line.split(",")[0].strip()
        info["has_gpu"] = True
        info["name"]    = name
        _log(f"GPU detected: {name}")

        # RTX 50xx = Blackwell — requires cu128 + PyTorch 2.7+
        # Detect by model number: RTX 5060, 5070, 5080, 5090, etc.
        import re
        m = re.search(r'RTX\s+(\d{4})', name, re.IGNORECASE)
        if m:
            model_num = int(m.group(1))
            if model_num >= 5000:
                _log(f"Blackwell GPU detected ({name}) — will use PyTorch 2.7 + cu128")
                info["is_blackwell"]  = True
                info["cuda_version"]  = "12.8"
                info["torch_version"] = "2.7.0"
                info["torch_index"]   = "https://download.pytorch.org/whl/cu128"
            else:
                _log(f"Pre-Blackwell GPU ({name}) — using PyTorch 2.6 + cu124")
    except Exception as e:
        _log(f"GPU detection error: {e}")
    return info


def has_nvidia_gpu():
    """Legacy helper — returns True/False only."""
    return get_gpu_info()["has_gpu"]


def find_python() -> dict:
    """
    Find the best available Python for this install.
    Returns dict: {cmd, version_str, major, minor, ok}

    Priority:
      1. python3.11 — ideal, most battle-tested with all deps
      2. python3.13 / python / python3 / py — also works with our package set
    We require Python 3.11 or 3.13. 3.12 also works. 3.10 and below: unsupported.
    """
    candidates = ["python3.11", "python3.13", "python3.12", "python", "python3", "py"]
    best = {"cmd": None, "version_str": "", "major": 0, "minor": 0, "ok": False}

    for cmd in candidates:
        try:
            r = subprocess.run(
                [cmd, "--version"], capture_output=True, text=True, timeout=5)
            if r.returncode != 0:
                continue
            ver = r.stdout.strip().split()[1]
            parts = ver.split(".")
            major, minor = int(parts[0]), int(parts[1])
            if major == 3 and minor >= 11:
                _log(f"Found Python {ver} at: {cmd}")
                return {"cmd": cmd, "version_str": ver,
                        "major": major, "minor": minor, "ok": True}
            else:
                _log(f"Skipping Python {ver} (too old) at: {cmd}")
        except Exception:
            continue
    return best


def find_python311():
    """Legacy alias."""
    return find_python().get("cmd")


# ── API validation ─────────────────────────────────────────────────────────────

def test_anthropic_key(key: str) -> tuple:
    """Validate Anthropic API key. Returns (ok, message)."""
    _log(f"Testing Anthropic key: {key[:8]}...{key[-4:]}")
    try:
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/models",
            headers={
                "x-api-key": key,
                "anthropic-version": "2023-06-01",
            })
        with urllib.request.urlopen(req, timeout=10) as resp:
            _log("Anthropic key valid")
            return True, "API key valid ✓"
    except urllib.error.HTTPError as e:
        body = ""
        try:
            body = e.read().decode()[:200]
        except Exception:
            pass
        _log(f"Anthropic HTTP error: {e.code} — {body}")
        if e.code == 401:
            return False, "Invalid API key — check you copied it correctly"
        if e.code == 403:
            return False, "Key valid but access denied — check billing at console.anthropic.com"
        return False, f"HTTP {e.code}: {body[:100]}"
    except urllib.error.URLError as e:
        _log(f"Anthropic URL error: {e.reason}")
        return False, f"Cannot reach Anthropic: {e.reason}"
    except Exception as e:
        _log(f"Anthropic unexpected error: {e}")
        return False, f"Unexpected error: {str(e)[:100]}"


def test_hf_token(token: str) -> tuple:
    """Validate HuggingFace token. Returns (ok, message)."""
    _log(f"Testing HF token: {token[:8]}...{token[-4:]}")
    try:
        req = urllib.request.Request(
            "https://huggingface.co/api/whoami",
            headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            raw  = resp.read()
            data = json.loads(raw)
            name = data.get("name", "unknown")
            _log(f"HF token valid — user: {name}")
            return True, f"Token valid — logged in as {name} ✓"
    except urllib.error.HTTPError as e:
        body = ""
        try:
            body = e.read().decode()[:300]
        except Exception:
            pass
        _log(f"HF HTTP error: {e.code} {e.reason} — {body}")
        if e.code == 401:
            return False, (
                "Invalid token (401 Unauthorized)\n\n"
                "Make sure you copied the full token starting with 'hf_'\n"
                "Try generating a new token at huggingface.co/settings/tokens"
            )
        if e.code == 403:
            return False, (
                "Token valid but access denied (403)\n\n"
                "Check token permissions — it needs at least 'read' scope"
            )
        return False, f"HTTP {e.code}: {e.reason}\n\n{body[:100]}"
    except urllib.error.URLError as e:
        _log(f"HF URL error: {e.reason}")
        return False, (
            f"Cannot reach HuggingFace: {e.reason}\n\n"
            "Check your internet connection"
        )
    except Exception as e:
        _log(f"HF unexpected error: {type(e).__name__}: {e}")
        return False, f"Unexpected error: {type(e).__name__}: {str(e)[:100]}"


def check_pyannote_access(token: str) -> tuple:
    """
    Check if both pyannote gated models are accessible with this token.
    Returns (ok, message).

    NOTE: HuggingFace returns {"error": "Repository not found"} in the
    response body (even on HTTP 200) when a token is valid but the user
    has not accepted the model's terms of service. This function catches
    that case explicitly instead of treating it as success.
    """
    _log("Checking pyannote model access...")
    models = [
        "pyannote/speaker-diarization-3.1",
        "pyannote/segmentation-3.0",
    ]

    for model_id in models:
        _log(f"  Checking: {model_id}")
        api_url = f"https://huggingface.co/api/models/{model_id}"
        try:
            req = urllib.request.Request(
                api_url,
                headers={"Authorization": f"Bearer {token}"}
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                raw  = resp.read()
                data = json.loads(raw)

                # ── KEY FIX ──────────────────────────────────────────────────
                # HuggingFace returns {"error": "Repository not found"} in the
                # body with HTTP 200 when access is blocked by gating.
                # The old code never checked for this — it only checked
                # for a nonexistent "gatedConsent" field — so blocked tokens
                # were silently treated as success. Now we catch it explicitly.
                if "error" in data:
                    err_msg = data["error"]
                    _log(f"  Blocked: {model_id} — {err_msg}")
                    return False, (
                        f"Access blocked for {model_id}\n\n"
                        f"HuggingFace says: \"{err_msg}\"\n\n"
                        f"This usually means:\n"
                        f"  • You haven't accepted the model's terms of service, OR\n"
                        f"  • Your token was generated BEFORE accepting terms\n"
                        f"    (tokens don't retroactively get new access)\n\n"
                        f"How to fix:\n"
                        f"  1. Accept terms at:\n"
                        f"     huggingface.co/pyannote/speaker-diarization-3.1\n"
                        f"     huggingface.co/pyannote/segmentation-3.0\n"
                        f"  2. Generate a NEW token at:\n"
                        f"     huggingface.co/settings/tokens\n"
                        f"  3. Re-run this installer with the new token\n\n"
                        f"Log: {LOG_PATH}"
                    )

                # Sanity check — response should have modelId or id
                if not data.get("modelId") and not data.get("id"):
                    _log(f"  Unexpected response for {model_id}: {str(data)[:100]}")
                    return False, (
                        f"Unexpected response for {model_id}\n\n"
                        f"Token may lack 'read' permissions.\n"
                        f"Try regenerating your token with 'read' scope at:\n"
                        f"huggingface.co/settings/tokens"
                    )

                _log(f"  OK: {model_id}")

        except urllib.error.HTTPError as e:
            body = ""
            try:
                body = e.read().decode()[:300]
            except Exception:
                pass
            _log(f"  HTTP {e.code} for {model_id}: {body[:150]}")

            if e.code in (401, 403):
                return False, (
                    f"Access denied for {model_id} (HTTP {e.code})\n\n"
                    f"Steps to fix:\n"
                    f"  1. Accept model terms at:\n"
                    f"     huggingface.co/pyannote/speaker-diarization-3.1\n"
                    f"     huggingface.co/pyannote/segmentation-3.0\n"
                    f"  2. Generate a NEW token at:\n"
                    f"     huggingface.co/settings/tokens\n"
                    f"  3. Re-run this installer with the new token\n\n"
                    f"Detail: {body[:100]}\nLog: {LOG_PATH}"
                )
            return False, (
                f"HTTP {e.code} checking {model_id}\n\n"
                f"{body[:150]}\nLog: {LOG_PATH}"
            )

        except urllib.error.URLError as e:
            _log(f"  URL error for {model_id}: {e.reason}")
            return False, (
                f"Cannot reach HuggingFace to verify {model_id}\n\n"
                f"Check internet connection: {e.reason}"
            )

        except Exception as e:
            _log(f"  Unexpected error for {model_id}: {type(e).__name__}: {e}")
            return False, (
                f"Unexpected error checking {model_id}:\n"
                f"{type(e).__name__}: {str(e)[:100]}"
            )

    _log("pyannote model access confirmed for all models")
    return True, "pyannote model access confirmed ✓"


# ── Shortcut helpers ───────────────────────────────────────────────────────────

def get_desktop_path() -> Path:
    """
    Return the real Desktop path for the current user.

    Handles three cases that all silently break the naive expanduser("~") approach:
      1. Installer run as Administrator → ~ resolves to C:\\Windows\\System32\\...
         or C:\\Users\\Administrator, not the actual user's Desktop.
      2. OneDrive folder redirection → Desktop is at
         C:\\Users\\[name]\\OneDrive\\Desktop, not C:\\Users\\[name]\\Desktop.
      3. Custom shell folder location set via registry (corporate GPO).

    We read the actual path from the Windows registry (SHGetFolderPath equivalent)
    which is always correct regardless of the above.
    """
    try:
        # Registry key that always holds the real Desktop path
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
        desktop, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        p = Path(desktop)
        if p.exists():
            _log(f"Desktop path (registry): {p}")
            return p
    except Exception as e:
        _log(f"Registry desktop lookup failed: {e}")

    # Fallback 1: USERPROFILE env var (more reliable than ~ when running as admin)
    userprofile = os.environ.get("USERPROFILE", "")
    if userprofile:
        candidates = [
            Path(userprofile) / "OneDrive" / "Desktop",  # OneDrive redirect
            Path(userprofile) / "Desktop",                # Standard
        ]
        for c in candidates:
            if c.exists():
                _log(f"Desktop path (USERPROFILE fallback): {c}")
                return c

    # Fallback 2: expanduser — least reliable but last resort
    fallback = Path(os.path.expanduser("~")) / "Desktop"
    _log(f"Desktop path (expanduser fallback): {fallback}")
    return fallback


def create_vbs(install_dir: str) -> str:
    """
    Write a VBScript launcher that activates the venv and runs main.py.
    Uses the full path to the venv python.exe — NOT the system 'python'
    command, which may not be in PATH or may point to the wrong version.
    """
    pyexe = str(Path(install_dir) / ".venv" / "Scripts" / "python.exe")
    main_py = str(Path(install_dir) / "main.py")
    vbs_content = (
        'Set WshShell = CreateObject("WScript.Shell")\n'
        f'WshShell.CurrentDirectory = "{install_dir}"\n'
        f'WshShell.Run Chr(34) & "{pyexe}" & Chr(34) & " " & '
        f'Chr(34) & "{main_py}" & Chr(34), 0, False\n'
    )
    path = Path(install_dir) / "launch.vbs"
    path.write_text(vbs_content, encoding="utf-8")
    _log(f"VBScript written: {path}")
    return str(path)


def create_shortcut(install_dir: str, vbs_path: str) -> str:
    """
    Create a .lnk Desktop shortcut pointing to the VBScript launcher.
    Uses a temp .ps1 file (not inline -Command) for reliability —
    inline PowerShell breaks on nested quotes and backslash paths.
    """
    desktop  = get_desktop_path()
    lnk_path = desktop / f"{APP_NAME}.lnk"
    ico_path = Path(install_dir) / "meeting_recorder.ico"

    _log(f"Creating shortcut at: {lnk_path}")
    _log(f"  → VBS target: {vbs_path}")
    _log(f"  → Desktop resolved to: {desktop}")

    icon_line = (f'$sc.IconLocation = "{ico_path}"\n'
                 if ico_path.exists() else "")

    ps_content = (
        f'$ws = New-Object -ComObject WScript.Shell\n'
        f'$sc = $ws.CreateShortcut("{lnk_path}")\n'
        f'$sc.TargetPath = "wscript.exe"\n'
        f'$sc.Arguments = \'"{vbs_path}"\'\n'
        f'$sc.WorkingDirectory = "{install_dir}"\n'
        f'$sc.Description = "Launch Meeting Recorder"\n'
        f'{icon_line}'
        f'$sc.Save()\n'
        f'Write-Output "SHORTCUT_OK: {lnk_path}"\n'
    )

    ps_path = Path(install_dir) / "_make_shortcut.ps1"
    ps_path.write_text(ps_content, encoding="utf-8")
    _log(f"  PowerShell script written: {ps_path}")

    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", str(ps_path)],
            capture_output=True, text=True, timeout=20
        )
        _log(f"PowerShell stdout: {r.stdout.strip()}")
        if r.stderr.strip():
            _log(f"PowerShell stderr: {r.stderr.strip()}")
    finally:
        try:
            ps_path.unlink()
        except OSError:
            pass

    if "SHORTCUT_OK" in r.stdout and lnk_path.exists():
        _log("  ✓ Desktop shortcut created successfully")
    else:
        _log(f"  ⚠ Shortcut not created — PS exit code: {r.returncode}")
        # Attempt fallback: write a .bat file to Desktop instead
        _log("  Attempting .bat fallback shortcut…")
        try:
            bat_path = desktop / f"{APP_NAME}.bat"
            pyexe    = str(Path(install_dir) / ".venv" / "Scripts" / "pythonw.exe")
            bat_path.write_text(
                f'@echo off\n'
                f'cd /d "{install_dir}"\n'
                f'start "" "{pyexe}" "{install_dir}\\main.py"\n',
                encoding="utf-8")
            _log(f"  ✓ .bat fallback written: {bat_path}")
        except Exception as e2:
            _log(f"  ✗ .bat fallback also failed: {e2}")

    return str(lnk_path)


# ── Installer GUI ──────────────────────────────────────────────────────────────

class Installer(tk.Tk):
    def __init__(self):
        super().__init__()
        _log_start()
        _log("Installer launched")

        self.title(f"{APP_NAME} Setup")
        self.geometry("680x560")
        self.minsize(680, 540)
        self.configure(bg=BG)
        self.resizable(True, True)

        # Center window
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - 680) // 2
        y = (self.winfo_screenheight() - 560) // 2
        self.geometry(f"+{x}+{y}")

        self._step           = 0
        self._python_info    = {}   # populated in _step_python
        self._python_cmd     = None
        self._install_dir    = tk.StringVar(value=str(
            Path(os.path.expanduser("~")) / "meeting_recorder"))
        self._anthropic_key  = tk.StringVar()
        self._hf_token       = tk.StringVar()
        self._gpu_info       = get_gpu_info()
        self._gpu            = self._gpu_info["has_gpu"]
        self._pyannote_verified = False

        _log(f"GPU info: {self._gpu_info}")
        _log(f"Default install dir: {self._install_dir.get()}")

        self._build_ui()
        self._show_step(0)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── UI shell ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=PANEL, height=64)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"  🎙  {APP_NAME} Setup",
                 font=FONT_TITLE, bg=PANEL, fg=ACCENT).pack(
                     side=tk.LEFT, padx=16, pady=12)

        # Separator
        tk.Frame(self, bg=BORDER, height=1).pack(fill=tk.X)

        # Content area
        self._content = tk.Frame(self, bg=BG)
        self._content.pack(fill=tk.BOTH, expand=True, padx=32, pady=20)

        # Footer buttons
        tk.Frame(self, bg=BORDER, height=1).pack(fill=tk.X)
        foot = tk.Frame(self, bg=PANEL, height=52)
        foot.pack(fill=tk.X)
        foot.pack_propagate(False)

        self._back_btn = tk.Button(
            foot, text="← Back",
            font=FONT_BODY, bg=INPUT, fg=MUTED,
            activebackground=BORDER, activeforeground=TEXT,
            relief=tk.FLAT, bd=0, padx=18, pady=6,
            cursor="hand2", command=self._back)
        self._back_btn.pack(side=tk.LEFT, padx=12, pady=10)

        self._next_btn = tk.Button(
            foot, text="Next →",
            font=FONT_BODY, bg=ACCENT2, fg="white",
            activebackground=ACCENT, activeforeground=BG,
            relief=tk.FLAT, bd=0, padx=22, pady=6,
            cursor="hand2", command=self._next)
        self._next_btn.pack(side=tk.RIGHT, padx=12, pady=10)

    def _clear(self):
        for w in self._content.winfo_children():
            w.destroy()

    def _lbl(self, parent, text, font=None, fg=None, bg=None, **kw):
        return tk.Label(
            parent, text=text,
            font=font or FONT_BODY,
            fg=fg or TEXT, bg=bg or BG,
            anchor="w", justify="left", **kw)

    def _show_step(self, n):
        self._step = n
        self._clear()
        steps = [
            self._step_welcome,
            self._step_location,
            self._step_python,
            self._step_anthropic,
            self._step_huggingface,
            self._step_pyannote,
            self._step_install,
            self._step_done,
        ]
        steps[n]()
        self._back_btn.config(state=tk.NORMAL if n > 0 else tk.DISABLED)
        self._next_btn.config(state=tk.NORMAL, text="Next →")

    def _next(self):
        # Block advancing past pyannote check (step 5) unless verified
        if self._step == 5 and not self._pyannote_verified:
            messagebox.showwarning(
                "Model Access Required",
                "Please click 'Check Access' to verify your HuggingFace token "
                "has access to the pyannote models before continuing.")
            return
        self._show_step(self._step + 1)

    def _back(self):
        if self._step > 0:
            self._show_step(self._step - 1)

    # ── Step 0: Welcome ────────────────────────────────────────────────────────

    def _step_welcome(self):
        self._next_btn.config(text="Get Started →")
        self._back_btn.config(state=tk.DISABLED)

        self._lbl(self._content, "Welcome to Meeting Recorder Setup",
                  font=FONT_H2, fg=ACCENT).pack(anchor="w", pady=(0, 12))

        info = (
            "This wizard will install Meeting Recorder on your computer.\n\n"
            "What you'll need before we start:\n"
            "  • An Anthropic API key  (console.anthropic.com)\n"
            "  • A HuggingFace account and token  (huggingface.co)\n"
            "  • HuggingFace model terms accepted for pyannote\n"
            "  • ~5 GB of free disk space\n"
            "  • Internet connection throughout setup\n\n"
            "Don't have these yet? This wizard will walk you through\n"
            "getting each one — just click Next to continue."
        )
        self._lbl(self._content, info, font=FONT_BODY).pack(
            anchor="w", pady=(0, 16))

        if self._gpu:
            gpu_name = self._gpu_info["name"]
            if self._gpu_info["is_blackwell"]:
                gpu_txt = f"✓ {gpu_name} detected — Blackwell GPU, will install PyTorch 2.7 + CUDA 12.8"
            else:
                gpu_txt = f"✓ {gpu_name} detected — will install CUDA acceleration"
            gpu_fg = SUCCESS_DIM
        else:
            gpu_txt = "⚠ No NVIDIA GPU detected — will install CPU-only mode (slower)"
            gpu_fg  = WARN_FG
        self._lbl(self._content, gpu_txt, fg=gpu_fg).pack(anchor="w")

    # ── Step 1: Install location ───────────────────────────────────────────────

    def _step_location(self):
        self._lbl(self._content, "Choose Install Location",
                  font=FONT_H2, fg=ACCENT).pack(anchor="w", pady=(0, 12))
        self._lbl(self._content,
                  "Where would you like to install Meeting Recorder?").pack(
                      anchor="w", pady=(0, 8))

        row = tk.Frame(self._content, bg=BG)
        row.pack(fill=tk.X, pady=(0, 8))

        entry = tk.Entry(row, textvariable=self._install_dir,
                         font=FONT_BODY, bg=INPUT, fg=TEXT,
                         insertbackground=TEXT, relief=tk.FLAT,
                         bd=6)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)

        tk.Button(row, text="Browse…",
                  font=FONT_SMALL, bg=PANEL, fg=TEXT,
                  activebackground=BORDER, relief=tk.FLAT, bd=0,
                  padx=12, pady=4, cursor="hand2",
                  command=self._browse_dir).pack(side=tk.LEFT, padx=(8, 0))

        self._lbl(self._content,
                  "Approximately 4–5 GB will be used (mostly PyTorch).",
                  fg=MUTED, font=FONT_SMALL).pack(anchor="w")

    def _browse_dir(self):
        d = filedialog.askdirectory(title="Choose Install Location")
        if d:
            self._install_dir.set(d.replace("/", "\\"))

    # ── Step 2: Python check ───────────────────────────────────────────────────

    def _step_python(self):
        self._lbl(self._content, "Python Version Check",
                  font=FONT_H2, fg=ACCENT).pack(anchor="w", pady=(0, 12))

        py_info = find_python()
        self._python_info = py_info
        self._python_cmd  = py_info.get("cmd")

        if py_info["ok"]:
            ver = py_info["version_str"]
            cmd = py_info["cmd"]
            _log(f"Python OK: {ver} at {cmd}")

            # Version-specific notes
            if py_info["minor"] == 11:
                note = "Python 3.11 — ideal version, all packages install cleanly."
                note_fg = SUCCESS_DIM
            elif py_info["minor"] == 13:
                note = (
                    "Python 3.13 detected. Fully supported — installer will use\n"
                    "faster-whisper and PyTorch 2.6+ which are the only builds\n"
                    "with Python 3.13 wheels. All known issues handled automatically."
                )
                note_fg = WARN_FG
            else:
                note = f"Python {ver} detected. Should work fine."
                note_fg = SUCCESS_DIM

            self._lbl(self._content,
                      f"✓ Python {ver} found ({cmd})",
                      fg=SUCCESS_DIM).pack(anchor="w", pady=(0, 8))
            self._lbl(self._content, note, fg=note_fg,
                      font=FONT_SMALL).pack(anchor="w", pady=(0, 8))
            self._lbl(self._content,
                      "Click Next to continue.",
                      fg=MUTED, font=FONT_SMALL).pack(anchor="w")
        else:
            _log("No suitable Python found")
            msg = (
                "Python 3.11 or newer is required but was not found.\n\n"
                "Click below to download Python 3.11.9 (recommended):\n"
                "  1. Run the downloaded installer\n"
                "  2. Check 'Add Python to PATH' on the first screen\n"
                "  3. Click Install Now\n"
                "  4. Close this installer and re-run it\n"
            )
            self._lbl(self._content, msg, fg=WARN_FG).pack(anchor="w", pady=(0, 12))

            tk.Button(self._content,
                      text="  Download Python 3.11.9  ↗",
                      font=FONT_BODY, bg=ACCENT2, fg="white",
                      activebackground=ACCENT, activeforeground=BG,
                      relief=tk.FLAT, bd=0, padx=16, pady=8,
                      cursor="hand2",
                      command=lambda: webbrowser.open(PY_URL)
                      ).pack(anchor="w")
            self._next_btn.config(state=tk.DISABLED)

    # ── Step 3: Anthropic API key ──────────────────────────────────────────────

    def _step_anthropic(self):
        self._lbl(self._content, "Anthropic API Key",
                  font=FONT_H2, fg=ACCENT).pack(anchor="w", pady=(0, 8))

        how = (
            "How to get your Anthropic API key:\n"
            "  1. Go to console.anthropic.com and sign in (or create an account)\n"
            "  2. Click 'API Keys' in the left sidebar\n"
            "  3. Click 'Create Key', give it a name, copy the key\n"
            "  4. Add billing at console.anthropic.com/settings/billing\n"
            "     (a $5 credit lasts many months of normal use)\n"
        )
        self._lbl(self._content, how, font=FONT_SMALL, fg=MUTED).pack(
            anchor="w", pady=(0, 8))

        tk.Button(self._content, text="  Open Anthropic Console  ↗",
                  font=FONT_SMALL, bg=PANEL, fg=ACCENT,
                  activebackground=BORDER, relief=tk.FLAT, bd=0,
                  padx=12, pady=4, cursor="hand2",
                  command=lambda: webbrowser.open(
                      "https://console.anthropic.com/settings/keys")
                  ).pack(anchor="w", pady=(0, 12))

        self._lbl(self._content, "Paste your Anthropic API key:").pack(anchor="w")
        self._anthropic_entry = tk.Entry(
            self._content,
            textvariable=self._anthropic_key,
            font=FONT_MONO, bg=INPUT, fg=TEXT,
            insertbackground=TEXT, relief=tk.FLAT,
            bd=6, show="•", width=56)
        self._anthropic_entry.pack(anchor="w", ipady=5, pady=(4, 8))

        row = tk.Frame(self._content, bg=BG)
        row.pack(anchor="w")
        self._anthopic_status = tk.Label(row, text="", font=FONT_SMALL,
                                         bg=BG, fg=MUTED)
        self._anthopic_status.pack(side=tk.LEFT)

        tk.Button(row, text="Test Key",
                  font=FONT_SMALL, bg=PANEL, fg=TEXT,
                  activebackground=BORDER, relief=tk.FLAT, bd=0,
                  padx=12, pady=4, cursor="hand2",
                  command=self._test_anthropic).pack(side=tk.LEFT, padx=(8, 0))

    def _test_anthropic(self):
        key = self._anthropic_key.get().strip()
        if not key:
            self._anthopic_status.config(text="Please enter a key first", fg=WARN_FG)
            return
        self._anthopic_status.config(text="Testing…", fg=MUTED)
        self.update()
        ok, msg = test_anthropic_key(key)
        self._anthopic_status.config(
            text=("✓ " if ok else "✗ ") + msg.split("\n")[0],
            fg=SUCCESS_DIM if ok else DANGER)

    # ── Step 4: HuggingFace token ──────────────────────────────────────────────

    def _step_huggingface(self):
        self._lbl(self._content, "HuggingFace Token",
                  font=FONT_H2, fg=ACCENT).pack(anchor="w", pady=(0, 8))

        how = (
            "How to get your HuggingFace token:\n"
            "  1. Go to huggingface.co and sign in (or create a free account)\n"
            "  2. Click your profile picture → Settings → Access Tokens\n"
            "  3. Click 'New token', select 'Read' role, copy the token\n"
            "     (token starts with 'hf_')\n"
        )
        self._lbl(self._content, how, font=FONT_SMALL, fg=MUTED).pack(
            anchor="w", pady=(0, 8))

        tk.Button(self._content, text="  Open HuggingFace Tokens  ↗",
                  font=FONT_SMALL, bg=PANEL, fg=ACCENT,
                  activebackground=BORDER, relief=tk.FLAT, bd=0,
                  padx=12, pady=4, cursor="hand2",
                  command=lambda: webbrowser.open(
                      "https://huggingface.co/settings/tokens")
                  ).pack(anchor="w", pady=(0, 12))

        self._lbl(self._content, "Paste your HuggingFace token:").pack(anchor="w")
        self._hf_entry = tk.Entry(
            self._content,
            textvariable=self._hf_token,
            font=FONT_MONO, bg=INPUT, fg=TEXT,
            insertbackground=TEXT, relief=tk.FLAT,
            bd=6, show="•", width=56)
        self._hf_entry.pack(anchor="w", ipady=5, pady=(4, 8))

        row = tk.Frame(self._content, bg=BG)
        row.pack(anchor="w")
        self._hf_status = tk.Label(row, text="", font=FONT_SMALL, bg=BG, fg=MUTED)
        self._hf_status.pack(side=tk.LEFT)

        tk.Button(row, text="Test Token",
                  font=FONT_SMALL, bg=PANEL, fg=TEXT,
                  activebackground=BORDER, relief=tk.FLAT, bd=0,
                  padx=12, pady=4, cursor="hand2",
                  command=self._test_hf).pack(side=tk.LEFT, padx=(8, 0))

    def _test_hf(self):
        token = self._hf_token.get().strip()
        if not token:
            self._hf_status.config(text="Please enter a token first", fg=WARN_FG)
            return
        self._hf_status.config(text="Testing…", fg=MUTED)
        self.update()
        ok, msg = test_hf_token(token)
        self._hf_status.config(
            text=("✓ " if ok else "✗ ") + msg.split("\n")[0],
            fg=SUCCESS_DIM if ok else DANGER)

    # ── Step 5: pyannote model terms ───────────────────────────────────────────

    def _step_pyannote(self):
        self._lbl(self._content, "Accept pyannote Model Terms",
                  font=FONT_H2, fg=ACCENT).pack(anchor="w", pady=(0, 8))

        msg = (
            "Meeting Recorder uses pyannote.audio to identify who is speaking.\n"
            "These are gated models — you must accept their terms of service\n"
            "before your token can download them.\n\n"
            "IMPORTANT: Accept the terms BEFORE generating your HuggingFace\n"
            "token (or generate a new token after accepting — old tokens\n"
            "do not automatically get access to newly accepted models).\n"
        )
        self._lbl(self._content, msg, font=FONT_BODY).pack(anchor="w", pady=(0, 12))

        steps_txt = (
            "Steps:\n"
            "  1. Click each link below while logged into HuggingFace\n"
            "  2. Click 'Agree and access repository' on each page\n"
            "  3. Come back here and click 'Check Access' to verify\n"
        )
        self._lbl(self._content, steps_txt, font=FONT_SMALL, fg=MUTED).pack(
            anchor="w", pady=(0, 10))

        btn_frame = tk.Frame(self._content, bg=BG)
        btn_frame.pack(anchor="w", pady=(0, 12))

        tk.Button(btn_frame,
                  text="  Accept: pyannote/speaker-diarization-3.1  ↗",
                  font=FONT_SMALL, bg=ACCENT_BG, fg=ACCENT,
                  activebackground=BORDER, relief=tk.FLAT, bd=0,
                  padx=12, pady=5, cursor="hand2",
                  command=lambda: webbrowser.open(
                      "https://huggingface.co/pyannote/speaker-diarization-3.1")
                  ).pack(anchor="w", pady=2)

        tk.Button(btn_frame,
                  text="  Accept: pyannote/segmentation-3.0  ↗",
                  font=FONT_SMALL, bg=ACCENT_BG, fg=ACCENT,
                  activebackground=BORDER, relief=tk.FLAT, bd=0,
                  padx=12, pady=5, cursor="hand2",
                  command=lambda: webbrowser.open(
                      "https://huggingface.co/pyannote/segmentation-3.0")
                  ).pack(anchor="w", pady=2)

        row = tk.Frame(self._content, bg=BG)
        row.pack(anchor="w", pady=(8, 0))

        self._pyannote_status = tk.Label(row, text="", font=FONT_SMALL,
                                         bg=BG, fg=MUTED)
        self._pyannote_status.pack(side=tk.LEFT)

        tk.Button(row, text="Check Access",
                  font=FONT_SMALL, bg=PANEL, fg=TEXT,
                  activebackground=BORDER, relief=tk.FLAT, bd=0,
                  padx=12, pady=4, cursor="hand2",
                  command=self._check_pyannote).pack(side=tk.LEFT, padx=(8, 0))

    def _check_pyannote(self):
        token = self._hf_token.get().strip()
        if not token:
            self._pyannote_status.config(
                text="Enter your HuggingFace token on the previous screen first",
                fg=WARN_FG)
            return
        self._pyannote_status.config(text="Checking…", fg=MUTED)
        self.update()
        ok, msg = check_pyannote_access(token)
        short = msg.split("\n")[0]
        self._pyannote_status.config(
            text=("✓ " if ok else "✗ ") + short,
            fg=SUCCESS_DIM if ok else DANGER)
        self._pyannote_verified = ok
        if not ok:
            messagebox.showwarning("Model Access Issue", msg)

    # ── Step 6: Install ────────────────────────────────────────────────────────

    def _step_install(self):
        if not self._anthropic_key.get().strip():
            messagebox.showwarning("Missing Key",
                "Please go back and enter your Anthropic API key.")
            self._show_step(3)
            return
        if not self._hf_token.get().strip():
            messagebox.showwarning("Missing Token",
                "Please go back and enter your HuggingFace token.")
            self._show_step(4)
            return

        self._next_btn.config(state=tk.DISABLED)
        self._back_btn.config(state=tk.DISABLED)

        self._lbl(self._content, "Installing…",
                  font=FONT_H2, fg=ACCENT).pack(anchor="w", pady=(0, 12))

        self._progress = ttk.Progressbar(self._content, length=580,
                                         mode="determinate")
        self._progress.pack(fill=tk.X, pady=(0, 8))

        self._prog_label = tk.Label(self._content, text="Starting…",
                                    font=FONT_SMALL, bg=BG, fg=MUTED,
                                    anchor="w")
        self._prog_label.pack(anchor="w", pady=(0, 8))

        self._log_box = tk.Text(
            self._content,
            font=FONT_MONO, bg=PANEL, fg=TEXT,
            relief=tk.FLAT, bd=4, height=14,
            state=tk.DISABLED, wrap=tk.WORD)
        self._log_box.pack(fill=tk.BOTH, expand=True)

        threading.Thread(target=self._run_install, daemon=True).start()

    def _set_progress(self, val, label=None):
        self.after(0, lambda: self._progress.config(value=val))
        if label:
            self.after(0, lambda: self._prog_label.config(text=label))

    def _log_write(self, msg):
        _log(msg)
        def _do():
            self._log_box.config(state=tk.NORMAL)
            self._log_box.insert(tk.END, msg + "\n")
            self._log_box.see(tk.END)
            self._log_box.config(state=tk.DISABLED)
        self.after(0, _do)

    def _run_install(self):
        try:
            install_dir = Path(self._install_dir.get())
            install_dir.mkdir(parents=True, exist_ok=True)
            _log(f"Install dir: {install_dir}")

            # ── Disk space check ───────────────────────────────────────
            free = get_free_disk_gb(str(install_dir.anchor))
            if free < MIN_DISK_GB:
                raise RuntimeError(
                    f"Not enough disk space.\n"
                    f"Need {MIN_DISK_GB}GB free, have {free:.1f}GB.\n"
                    f"Free up space on {install_dir.anchor} and retry.")

            # ── Step 1: locate Python ──────────────────────────────────
            self._set_progress(3, "Locating Python…")
            self._log_write("→ Finding Python…")
            py_info = self._python_info if self._python_info.get("ok") else find_python()
            if not py_info["ok"]:
                raise RuntimeError(
                    "Python 3.11+ not found.\n"
                    "Install Python 3.11 with 'Add to PATH' checked, "
                    "then re-run this installer.")
            py  = py_info["cmd"]
            ver = py_info["version_str"]
            self._log_write(f"  ✓ Python {ver} ({py})")

            # ── Step 2: create venv ────────────────────────────────────
            self._set_progress(6, "Creating virtual environment…")
            self._log_write("→ Creating virtual environment…")
            venv_dir = install_dir / ".venv"
            r = subprocess.run([py, "-m", "venv", str(venv_dir)],
                               capture_output=True, text=True)
            if r.returncode != 0:
                raise RuntimeError(
                    f"Could not create virtual environment:\n{r.stderr[:200]}\n\n"
                    "Try running the installer as Administrator.")
            self._log_write("  ✓ Virtual environment created")

            pip   = str(venv_dir / "Scripts" / "pip.exe")
            pyexe = str(venv_dir / "Scripts" / "python.exe")

            # ── Step 3: CRITICAL — bootstrap pip/setuptools/wheel FIRST ─
            # Without this, any package that builds from source will fail with:
            #   "No module named 'pkg_resources'"
            # This has been observed with: openai-whisper, numpy 1.x, and others.
            # Must happen before ANY other pip install.
            self._set_progress(9, "Upgrading pip, setuptools, wheel…")
            self._log_write("→ Bootstrapping pip + setuptools + wheel (required first step)…")
            r = subprocess.run(
                [pyexe, "-m", "pip", "install", "--upgrade",
                 "pip", "setuptools", "wheel"],
                capture_output=True, text=True)
            if r.returncode != 0:
                self._log_write(f"  ⚠ Bootstrap warning: {r.stderr[:100]}")
            else:
                self._log_write("  ✓ pip + setuptools + wheel bootstrapped")

            # ── Step 4: PyTorch — correct version for this GPU ─────────
            # GPU mapping:
            #   RTX 50xx (Blackwell) → PyTorch 2.7.0 + cu128
            #   RTX 10xx–40xx        → PyTorch 2.6.0 + cu124
            #   No GPU               → PyTorch 2.6.0 + CPU
            # Note: PyTorch 2.6.0 is the first build with Python 3.13 wheels.
            # Do NOT use 2.5.x — it has no Python 3.13 support.
            gpu_info = self._gpu_info
            if gpu_info["has_gpu"]:
                torch_ver  = gpu_info["torch_version"]
                torch_idx  = gpu_info["torch_index"]
                cuda_label = "CUDA 12.8 (Blackwell RTX 50xx)" \
                             if gpu_info["is_blackwell"] else "CUDA 12.4"
                self._set_progress(13,
                    f"Installing PyTorch {torch_ver} + {cuda_label}…")
                self._log_write(
                    f"→ Installing PyTorch {torch_ver} ({cuda_label})…")
                self._log_write(f"  GPU: {gpu_info['name']}")
                self._log_write("  (~2.5GB download — please wait, this is normal)")
                r = subprocess.run(
                    [pip, "install",
                     f"torch=={torch_ver}", f"torchaudio=={torch_ver}",
                     "--index-url", torch_idx],
                    capture_output=True, text=True)
            else:
                self._set_progress(13, "Installing PyTorch 2.6.0 (CPU)…")
                self._log_write("→ Installing PyTorch 2.6.0 (CPU-only, no GPU detected)…")
                self._log_write("  (~300MB download)")
                r = subprocess.run(
                    [pip, "install",
                     "torch==2.6.0", "torchaudio==2.6.0",
                     "--index-url", "https://download.pytorch.org/whl/cpu"],
                    capture_output=True, text=True)
            if r.returncode != 0:
                self._log_write(f"  ⚠ PyTorch warning: {r.stderr[:120]}")
            else:
                self._log_write("  ✓ PyTorch installed")
            self._set_progress(38)

            # ── Step 5: pin numpy BEFORE pyannote installs it ──────────
            # pyannote-metrics declares numpy>=2.2.2 but pyannote's own
            # code uses np.NaN which was REMOVED in numpy 2.0.
            # numpy 2.1.3 satisfies >=2.0, and we patch NaN in main.py.
            # Must be pinned before pyannote runs or it will pull 2.2.x+.
            self._set_progress(39, "Pinning numpy 2.1.3…")
            self._log_write("→ Pinning numpy==2.1.3 (pyannote compatibility fix)…")
            r = subprocess.run(
                [pip, "install", "numpy==2.1.3", "--force-reinstall"],
                capture_output=True, text=True)
            if r.returncode == 0:
                self._log_write("  ✓ numpy==2.1.3 pinned")
            else:
                self._log_write(f"  ⚠ numpy pin warning: {r.stderr[:80]}")

            # ── Step 6: PyAudioWPatch (WASAPI loopback for system audio) ─
            self._set_progress(42, "Installing PyAudioWPatch…")
            self._log_write("→ Installing PyAudioWPatch (WASAPI loopback support)…")
            r = subprocess.run([pip, "install", "pyaudiowpatch"],
                               capture_output=True, text=True)
            if r.returncode != 0:
                self._log_write(
                    "  ⚠ PyAudioWPatch install failed.\n"
                    "    System audio capture may not work.\n"
                    f"    Error: {r.stderr[:200]}")
            else:
                self._log_write("  ✓ PyAudioWPatch installed")
            self._set_progress(46)

            # ── Step 7: all remaining packages ────────────────────────
            # NOTE: faster-whisper is used instead of openai-whisper.
            # openai-whisper fails to build on Python 3.13 (missing pkg_resources
            # during wheel build even after setuptools upgrade).
            # faster-whisper is a pre-built drop-in that is also faster + uses
            # less VRAM. The app's transcription.py is already written for it.
            total = len(REQUIREMENTS_BASE)
            for i, pkg in enumerate(REQUIREMENTS_BASE):
                pct  = 47 + int((i / total) * 32)
                name = pkg.split("==")[0]
                self._set_progress(pct, f"Installing {name}…")
                self._log_write(f"→ Installing {name}…")
                r = subprocess.run(
                    [pip, "install", pkg],
                    capture_output=True, text=True)
                if r.returncode != 0:
                    self._log_write(f"  ⚠ {name}: {r.stderr[:100]}")
                    if any(x in r.stderr.lower()
                           for x in ("firewall", "proxyerror", "connectionerror")):
                        raise RuntimeError(
                            f"Network blocked installing {name}.\n"
                            "Corporate firewall may be blocking pip.\n"
                            "Ask IT to whitelist: pypi.org, "
                            "download.pytorch.org, files.pythonhosted.org")
                else:
                    self._log_write(f"  ✓ {name}")

            # ── Step 8: re-pin numpy (pyannote may have upgraded it) ───
            self._set_progress(81, "Re-pinning numpy 2.1.3…")
            self._log_write("→ Re-pinning numpy==2.1.3 (pyannote may have upgraded it)…")
            subprocess.run(
                [pip, "install", "numpy==2.1.3", "--force-reinstall"],
                capture_output=True, text=True)
            self._log_write("  ✓ numpy==2.1.3 confirmed")

            # ── Step 9: write app files ────────────────────────────────
            self._set_progress(84, "Writing application files…")
            self._log_write("→ Writing application files…")
            self._write_app_files(install_dir)
            self._log_write("  ✓ Application files written")

            # ── Step 10: write .env ────────────────────────────────────
            self._set_progress(87, "Saving API keys…")
            self._log_write("→ Saving configuration…")
            (install_dir / ".env").write_text(
                f"ANTHROPIC_API_KEY={self._anthropic_key.get().strip()}\n"
                f"HF_TOKEN={self._hf_token.get().strip()}\n"
                f"WHISPER_MODEL=base\n"
                f"MAX_SPEAKERS=8\n"
                f"RECORDINGS_DIR=recordings\n",
                encoding="utf-8")
            self._log_write("  ✓ API keys saved")

            # ── Step 11: HuggingFace login ─────────────────────────────
            self._set_progress(89, "Logging in to HuggingFace…")
            self._log_write("→ HuggingFace login…")
            subprocess.run(
                [pyexe, "-c",
                 f'from huggingface_hub import login; '
                 f'login(token="{self._hf_token.get().strip()}")'],
                capture_output=True)
            self._log_write("  ✓ HuggingFace login complete")

            # ── Step 12: inject ALL compatibility patches into main.py ─
            # Three patches are required for this stack to work together:
            #
            # Patch 1 — np.NaN / np.NAN aliases
            #   numpy 2.0 removed np.NaN. pyannote.audio's inference.py uses
            #   np.NaN on line 533. We restore the alias at startup.
            #
            # Patch 2 — torch.load weights_only=False
            #   PyTorch 2.6 changed the default to weights_only=True.
            #   pyannote explicitly passes weights_only=False but the sig
            #   changed — we override torch.load to force it.
            #
            # Patch 3 — TorchVersion safe globals
            #   PyTorch 2.6 serialization requires TorchVersion to be in the
            #   safe globals list or it raises an error on model load.
            self._set_progress(92, "Applying compatibility patches…")
            self._log_write("→ Injecting numpy + torch patches into main.py…")
            self._patch_main(install_dir)
            self._log_write("  ✓ All 3 patches applied")

            # ── Step 13: verify core imports ──────────────────────────
            self._set_progress(94, "Verifying installation…")
            self._log_write("→ Running import verification…")
            verify = (
                "import torch; "
                "import faster_whisper; "
                "import sounddevice; "
                "import anthropic; "
                "print('VERIFY_OK')"
            )
            r = subprocess.run(
                [pyexe, "-c", verify],
                capture_output=True, text=True, timeout=30)
            if "VERIFY_OK" in r.stdout:
                self._log_write("  ✓ All core imports verified")
            else:
                self._log_write(
                    f"  ⚠ Verification check returned unexpected output")
                self._log_write(f"    stderr: {r.stderr[:200]}")
                self._log_write("    App may still work — continuing…")

            # ── Step 14: enable Stereo Mix for system audio capture ─────
            self._set_progress(95, "Enabling Stereo Mix…")
            self._log_write("→ Attempting to enable Stereo Mix for system audio capture…")
            try:
                # PowerShell command to enable Stereo Mix (disabled by default on most systems)
                ps_enable_mix = (
                    'try { '
                    '$dev = Get-AudioDevice -List -ErrorAction Stop | '
                    'Where-Object { $_.Type -eq "Recording" -and $_.Name -match "Stereo Mix" }; '
                    'if ($dev) { Write-Output "STEREO_MIX_FOUND" } '
                    'else { Write-Output "STEREO_MIX_NOT_FOUND" } '
                    '} catch { Write-Output "AUDIO_MODULE_UNAVAILABLE" }'
                )
                r = subprocess.run(
                    ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                     "-Command", ps_enable_mix],
                    capture_output=True, text=True, timeout=10)
                if "STEREO_MIX_FOUND" in r.stdout:
                    self._log_write("  ✓ Stereo Mix device detected")
                else:
                    self._log_write(
                        "  ⚠ Stereo Mix not found or disabled.\n"
                        "    To capture system audio (other meeting participants),\n"
                        "    enable it manually: Sound Settings → Recording → \n"
                        "    right-click → Show Disabled Devices → Enable Stereo Mix\n"
                        "    OR install VB-Cable (free virtual audio cable).")
            except Exception as e:
                self._log_write(f"  ⚠ Could not check Stereo Mix: {e}")

            # ── Step 15: desktop shortcut ──────────────────────────────
            self._set_progress(97, "Creating desktop shortcut…")
            self._log_write("→ Creating desktop shortcut…")
            vbs = create_vbs(str(install_dir))
            lnk = create_shortcut(str(install_dir), vbs)
            # Check log to see if PowerShell confirmed success
            try:
                log_tail = LOG_PATH.read_text(encoding="utf-8")[-800:]
                if "SHORTCUT_OK" in log_tail:
                    self._log_write(f"  ✓ Desktop shortcut created: {lnk}")
                elif ".bat fallback written" in log_tail:
                    self._log_write(
                        "  ⚠ .lnk shortcut failed — a .bat launcher was placed\n"
                        f"    on your Desktop instead: {Path(lnk).parent / (APP_NAME + '.bat')}\n"
                        "    Double-click it to launch the app.")
                else:
                    self._log_write(
                        "  ⚠ Shortcut could not be created automatically.\n"
                        f"    You can launch manually by double-clicking:\n"
                        f"    {install_dir}\\launch.vbs")
            except Exception:
                self._log_write("  ✓ Desktop shortcut step completed")

            self._set_progress(100, "Installation complete!")
            self._log_write("\n" + "=" * 50)
            self._log_write("  INSTALLATION COMPLETE!")
            self._log_write("=" * 50)
            self._log_write("Launch from your Desktop shortcut.")
            self._log_write("First launch downloads AI models (~700MB, one time only).")
            self._log_write(f"Full log: {LOG_PATH}")

            self.after(1500, lambda: self._show_step(7))

        except Exception as e:
            _log(f"Install FAILED: {e}")
            self._log_write(f"\n✗ Installation failed: {e}")
            self._set_progress(self._progress["value"],
                               "Installation failed — see log above")
            self.after(0, lambda: messagebox.showerror(
                "Installation Failed",
                f"Something went wrong:\n\n{e}\n\n"
                f"Full log at:\n{LOG_PATH}"))
            self.after(0, lambda: self._next_btn.config(
                state=tk.NORMAL, text="Retry",
                command=lambda: threading.Thread(
                    target=self._run_install, daemon=True).start()))


    def _write_app_files(self, install_dir: Path):
        """Write all embedded app source files to disk."""
        for rel_path, content in APP_FILES.items():
            dest = install_dir / rel_path
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_text(content, encoding="utf-8")
        if APP_ICON_B64:
            import base64
            ico_path = install_dir / "meeting_recorder.ico"
            ico_path.write_bytes(base64.b64decode(APP_ICON_B64))

    def _patch_main(self, install_dir: Path):
        """Inject numpy/torch compatibility patches at top of main.py."""
        main_path = install_dir / "main.py"
        if not main_path.exists():
            return
        content = main_path.read_text(encoding="utf-8")
        patch_marker = "# COMPAT_PATCHES_APPLIED"
        if patch_marker in content:
            return  # already patched
        patch = f"""{patch_marker}
import torch
import numpy as np
from torch.torch_version import TorchVersion
if not hasattr(np, 'NaN'):
    np.NaN = np.nan
if not hasattr(np, 'NAN'):
    np.NAN = np.nan
torch.serialization.add_safe_globals([TorchVersion])
_orig_torch_load = torch.load
def _patched_torch_load(f, *a, **kw):
    kw['weights_only'] = False
    return _orig_torch_load(f, *a, **kw)
torch.load = _patched_torch_load
"""
        main_path.write_text(patch + content, encoding="utf-8")

    # ── Step 7: Done ───────────────────────────────────────────────────────────

    def _step_done(self):
        self._next_btn.config(text="Close", command=self.destroy)
        self._back_btn.config(state=tk.DISABLED)

        self._lbl(self._content, "Installation Complete!",
                  font=FONT_H2, fg=SUCCESS_DIM).pack(anchor="w", pady=(0, 16))

        msg = (
            "Meeting Recorder is installed and ready to use.\n\n"
            "• Double-click the Desktop shortcut to launch\n"
            "• First launch downloads AI models (~700MB) — takes 2–5 min\n"
            "• Select your microphone, then click Start Recording\n\n"
            "To capture other participants (system audio):\n"
            "• Enable 'Stereo Mix' in Windows Sound → Recording tab\n"
            "  (right-click → Show Disabled Devices → Enable)\n"
            "• Or install VB-Cable (free virtual audio cable)\n"
            "• Then select it as 'System Audio' in the app\n"
        )
        self._lbl(self._content, msg).pack(anchor="w", pady=(0, 16))

        tk.Button(self._content,
                  text="  Launch Meeting Recorder  →",
                  font=FONT_H2, bg=ACCENT2, fg="white",
                  activebackground=ACCENT, activeforeground=BG,
                  relief=tk.FLAT, bd=0, padx=20, pady=10,
                  cursor="hand2",
                  command=self._launch).pack(anchor="w")

    def _launch(self):
        install_dir = Path(self._install_dir.get())
        vbs = install_dir / "launch.vbs"
        if vbs.exists():
            os.startfile(str(vbs))
        else:
            messagebox.showinfo("Launch",
                f"Open a terminal, navigate to:\n{install_dir}\n"
                "and run:  .venv\\Scripts\\activate && python main.py")

    # ── Close handler ──────────────────────────────────────────────────────────

    def _on_close(self):
        if self._step == 6:
            if not messagebox.askyesno(
                    "Cancel Install?",
                    "Installation is in progress.\n\nAre you sure you want to quit?"):
                return
        self.destroy()


if __name__ == "__main__":
    app = Installer()
    app.mainloop()
