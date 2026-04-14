"""
Meeting Recorder — one-command CLI setup.

Usage:
    git clone https://github.com/joshuarodriguez82/meeting-recorder
    cd meeting-recorder
    python setup.py

Creates a venv, installs all dependencies, writes an empty .env,
and prints the command to launch the app.

Skips the GUI installer — use this when .exe is blocked by security
policies or when you prefer a CLI-only workflow.
"""

import os
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VENV = ROOT / ".venv"
PYEXE = VENV / "Scripts" / "python.exe"
PIP = VENV / "Scripts" / "pip.exe"


def run(cmd, check=True):
    print(f"  $ {' '.join(str(c) for c in cmd)}")
    result = subprocess.run(cmd, check=check)
    return result.returncode


def detect_gpu():
    """Return (torch_version, index_url) based on GPU."""
    try:
        out = subprocess.run(
            ["nvidia-smi", "--query-gpu=name", "--format=csv,noheader"],
            capture_output=True, text=True, timeout=5,
        )
        if out.returncode == 0 and out.stdout.strip():
            name = out.stdout.strip().split("\n")[0]
            print(f"  GPU detected: {name}")
            # RTX 50xx = Blackwell = CUDA 12.8 / torch 2.7
            m = re.search(r"RTX\s+(\d{4})", name, re.IGNORECASE)
            if m and int(m.group(1)) >= 5000:
                print("  Blackwell GPU — using torch 2.7 + cu128")
                return ("2.7.0", "https://download.pytorch.org/whl/cu128")
            print("  NVIDIA GPU — using torch 2.6 + cu124")
            return ("2.6.0", "https://download.pytorch.org/whl/cu124")
    except Exception:
        pass
    print("  No GPU detected — using CPU-only torch")
    return ("2.6.0", "https://download.pytorch.org/whl/cpu")


def main():
    print("=" * 60)
    print("  Meeting Recorder — CLI Setup")
    print("=" * 60)

    # 1. Create venv
    if not VENV.exists():
        print("\n[1/5] Creating virtual environment...")
        run([sys.executable, "-m", "venv", str(VENV)])
    else:
        print("\n[1/5] Virtual environment already exists — skipping.")

    # 2. Upgrade pip
    print("\n[2/5] Upgrading pip, setuptools, wheel...")
    run([str(PYEXE), "-m", "pip", "install", "--upgrade",
         "pip", "setuptools", "wheel"])

    # 3. Install PyTorch (GPU-specific)
    print("\n[3/5] Installing PyTorch...")
    torch_ver, index_url = detect_gpu()
    run([str(PIP), "install",
         f"torch=={torch_ver}", f"torchaudio=={torch_ver}",
         "--index-url", index_url])

    # 4. Install all other packages
    print("\n[4/5] Installing app dependencies...")
    packages = [
        "numpy==2.1.3",
        "pyaudiowpatch",
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
    run([str(PIP), "install", *packages])

    # 5. Create .env if missing
    env_path = ROOT / ".env"
    if not env_path.exists():
        print("\n[5/5] Creating empty .env file...")
        env_path.write_text(
            "ANTHROPIC_API_KEY=\n"
            "HF_TOKEN=\n"
            "WHISPER_MODEL=base\n"
            "MAX_SPEAKERS=10\n"
            "RECORDINGS_DIR=recordings\n"
            "EMAIL_TO=\n",
            encoding="utf-8",
        )
    else:
        print("\n[5/5] .env already exists — keeping it.")

    # Done
    print("\n" + "=" * 60)
    print("  SETUP COMPLETE")
    print("=" * 60)
    print("\nTo launch the app:")
    print(f'  {PYEXE} {ROOT / "main.py"}')
    print("\nOr activate the venv first:")
    print(f"  {VENV / 'Scripts' / 'activate.bat'}")
    print("  python main.py")
    print("\nThen in the app: File > Settings to add your API keys.")
    print("\nGet keys from:")
    print("  https://console.anthropic.com")
    print("  https://huggingface.co/settings/tokens")
    print("\nAccept model terms at:")
    print("  https://huggingface.co/pyannote/speaker-diarization-3.1")
    print("  https://huggingface.co/pyannote/segmentation-3.0")
    print()


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as e:
        print(f"\n[ERROR] Command failed: {e}")
        sys.exit(1)
