"""
Suppress transient console windows on Windows.

Under pythonw.exe the main app has no console, but any subprocess spawned
without CREATE_NO_WINDOW can still flash a cmd.exe window briefly — most
often when a third-party library (ffmpeg via torchaudio/pyannote, COM
helpers, etc.) runs a helper process.

`install()` monkey-patches subprocess.Popen to always add CREATE_NO_WINDOW
on Windows. Must be called before any module that may spawn subprocesses.
"""

import sys

CREATE_NO_WINDOW = 0x08000000  # win32 process creation flag


def install() -> None:
    if sys.platform != "win32":
        return
    import subprocess

    _orig_init = subprocess.Popen.__init__

    def _patched_init(self, *args, **kwargs):
        flags = kwargs.get("creationflags") or 0
        kwargs["creationflags"] = flags | CREATE_NO_WINDOW
        return _orig_init(self, *args, **kwargs)

    subprocess.Popen.__init__ = _patched_init  # type: ignore[method-assign]
