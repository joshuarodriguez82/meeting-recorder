"""
Run from C:\meeting_recorder to bundle all app files into the installer.

Usage:
    cd C:\meeting_recorder
    python installer\bundle.py
"""

import json
import base64
from pathlib import Path

INCLUDE = [
    "main.py",
    "config/__init__.py",
    "config/settings.py",
    "core/__init__.py",
    "core/audio_capture.py",
    "core/diarization.py",
    "core/summarizer.py",
    "core/transcription.py",
    "models/__init__.py",
    "models/segment.py",
    "models/session.py",
    "models/speaker.py",
    "services/__init__.py",
    "services/calendar_service.py",
    "services/calendar_monitor.py",
    "services/export_service.py",
    "services/recording_service.py",
    "services/session_service.py",
    "ui/__init__.py",
    "ui/app_window.py",
    "ui/calendar_panel.py",
    "ui/device_panel.py",
    "ui/follow_up_tracker.py",
    "ui/prep_brief_dialog.py",
    "ui/session_browser.py",
    "ui/settings_dialog.py",
    "ui/speaker_panel.py",
    "ui/styles.py",
    "ui/transcript_panel.py",
    "utils/__init__.py",
    "utils/audio_utils.py",
    "utils/logger.py",
]


def main():
    root = Path(__file__).parent.parent
    print(f"Bundling from: {root}\n")

    app_files = {}
    missing   = []
    for rel in INCLUDE:
        p = root / rel
        if p.exists():
            app_files[rel.replace("\\", "/")] = p.read_text(encoding="utf-8")
            print(f"  ✓ {rel}")
        else:
            missing.append(rel)
            print(f"  ✗ MISSING: {rel}")

    if missing:
        print(f"\n⚠ {len(missing)} files missing — they will be skipped.")

    # Icon
    ico = root / "meeting_recorder.ico"
    if ico.exists():
        ico_b64 = base64.b64encode(ico.read_bytes()).decode()
        print(f"  ✓ Icon embedded ({len(ico_b64)} chars)")
    else:
        ico_b64 = ""
        print("  ✗ meeting_recorder.ico not found")

    # Read installer template
    template = (Path(__file__).parent / "installer.py").read_text(encoding="utf-8")

    # Inject
    files_code = "APP_FILES    = " + json.dumps(app_files, indent=2, ensure_ascii=False)
    icon_code  = f'APP_ICON_B64 = "{ico_b64}"'
    output = template.replace(
        'APP_FILES    = {}   # populated by bundle.py\nAPP_ICON_B64 = ""',
        files_code + "\n" + icon_code
    )

    out_path = Path(__file__).parent / "installer_bundled.py"
    out_path.write_text(output, encoding="utf-8")
    print(f"\n✓ Bundled installer: {out_path}")
    print(f"  {len(app_files)} files embedded")
    print("\nNext: run build.bat to compile the .exe")


if __name__ == "__main__":
    main()
