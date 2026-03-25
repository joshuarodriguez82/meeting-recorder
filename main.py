"""
Meeting Recorder — application entry point.
"""

import sys
import torch
import numpy as np
from torch.torch_version import TorchVersion

# Patch NumPy 2.0 compatibility with pyannote
if not hasattr(np, 'NaN'):
    np.NaN = np.nan
if not hasattr(np, 'NAN'):
    np.NAN = np.nan

# Patch PyTorch 2.6 compatibility with pyannote
torch.serialization.add_safe_globals([TorchVersion])
_original_torch_load = torch.load
def _patched_torch_load(f, *args, **kwargs):
    kwargs['weights_only'] = False
    return _original_torch_load(f, *args, **kwargs)
torch.load = _patched_torch_load

from config.settings import Settings
from ui.app_window import AppWindow
from utils.logger import get_logger

logger = get_logger(__name__)


def main() -> None:
    try:
        settings = Settings.from_env()
    except EnvironmentError as e:
        print(f"\n[CONFIG ERROR]\n{e}\n")
        sys.exit(1)
    try:
        app.iconbitmap("meeting_recorder.ico")
    except Exception:
        pass

    logger.info("Starting Meeting Recorder...")
    app = AppWindow(settings)
    app.mainloop()


if __name__ == "__main__":
    main()