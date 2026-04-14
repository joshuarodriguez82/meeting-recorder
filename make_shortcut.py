"""
Creates a Meeting Recorder shortcut on the Desktop.

Run this from the meeting-recorder folder after setup.py:
    python make_shortcut.py

Works even when the installer's shortcut creation fails.
"""

import os
import subprocess
import sys
from pathlib import Path


def get_desktop() -> Path:
    """Find the real Desktop path (handles OneDrive redirect)."""
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
        desktop, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        if Path(desktop).exists():
            return Path(desktop)
    except Exception:
        pass

    userprofile = os.environ.get("USERPROFILE", "")
    for candidate in (Path(userprofile) / "OneDrive" / "Desktop",
                      Path(userprofile) / "Desktop"):
        if candidate.exists():
            return candidate
    return Path(os.path.expanduser("~")) / "Desktop"


def main():
    install_dir = Path(__file__).resolve().parent
    pyexe = install_dir / ".venv" / "Scripts" / "pythonw.exe"
    main_py = install_dir / "main.py"
    icon = install_dir / "meeting_recorder.ico"

    if not pyexe.exists():
        print(f"ERROR: venv not found at {pyexe}")
        print("Run 'python setup.py' first to create the venv.")
        sys.exit(1)
    if not main_py.exists():
        print(f"ERROR: main.py not found at {main_py}")
        sys.exit(1)

    desktop = get_desktop()
    lnk_path = desktop / "Meeting Recorder.lnk"
    print(f"Creating shortcut: {lnk_path}")

    # Write a temp PowerShell script and run it
    ps_script = install_dir / "_make_shortcut_temp.ps1"
    icon_line = f'$sc.IconLocation = "{icon}"\n' if icon.exists() else ""
    ps_content = (
        f'$ws = New-Object -ComObject WScript.Shell\n'
        f'$sc = $ws.CreateShortcut("{lnk_path}")\n'
        f'$sc.TargetPath = "{pyexe}"\n'
        f'$sc.Arguments = \'"{main_py}"\'\n'
        f'$sc.WorkingDirectory = "{install_dir}"\n'
        f'$sc.Description = "Launch Meeting Recorder"\n'
        f'{icon_line}'
        f'$sc.Save()\n'
        f'Write-Output "OK"\n'
    )
    ps_script.write_text(ps_content, encoding="utf-8")

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", str(ps_script)],
            capture_output=True, text=True, timeout=20)
        if "OK" in result.stdout and lnk_path.exists():
            print(f"\n[SUCCESS] Shortcut created: {lnk_path}")
            return
        print(f"[WARN] PowerShell output: {result.stdout.strip()}")
        if result.stderr.strip():
            print(f"[WARN] PowerShell stderr: {result.stderr.strip()}")
    except Exception as e:
        print(f"[WARN] PowerShell method failed: {e}")
    finally:
        try:
            ps_script.unlink()
        except OSError:
            pass

    # Fallback: .bat file on desktop
    print("\nFalling back to .bat launcher...")
    bat_path = desktop / "Meeting Recorder.bat"
    bat_path.write_text(
        f'@echo off\n'
        f'cd /d "{install_dir}"\n'
        f'start "" "{pyexe}" "{main_py}"\n',
        encoding="utf-8")
    print(f"[SUCCESS] .bat launcher created: {bat_path}")
    print("Double-click it to launch Meeting Recorder.")


if __name__ == "__main__":
    main()
