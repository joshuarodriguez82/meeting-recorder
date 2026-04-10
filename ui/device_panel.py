import json
import os
import tkinter as tk
from tkinter import ttk
from typing import Optional
from core.audio_capture import list_input_devices, list_output_devices
from ui import styles
from utils.logger import get_logger

logger = get_logger(__name__)

PREFS_FILE = "device_prefs.json"


def _load_prefs() -> dict:
    try:
        if os.path.exists(PREFS_FILE):
            with open(PREFS_FILE, "r") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _save_prefs(mic_index, out_index) -> None:
    try:
        with open(PREFS_FILE, "w") as f:
            json.dump({"mic_index": mic_index, "out_index": out_index}, f)
    except Exception:
        pass


class DevicePanel(tk.Frame):

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=styles.BG_PANEL, **kwargs)
        self._input_devices = list_input_devices()
        self._output_devices = list_output_devices()
        self._prefs = _load_prefs()
        self._build()

    def _build(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "M.TCombobox",
            fieldbackground=styles.BG_INPUT,
            background=styles.BG_INPUT,
            foreground=styles.TEXT_PRIMARY,
            selectbackground=styles.ACCENT_BG,
            selectforeground=styles.ACCENT,
            borderwidth=0,
            relief="flat",
            padding=(10, 8),
        )
        style.map("M.TCombobox",
                  fieldbackground=[("readonly", styles.BG_INPUT)],
                  foreground=[("readonly", styles.TEXT_PRIMARY)])

        mic_row = tk.Frame(self, bg=styles.BG_PANEL)
        mic_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(mic_row, text="Microphone", bg=styles.BG_PANEL,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL, width=12,
                 anchor="w").pack(side=tk.LEFT, padx=(2, 8))

        self._mic_var = tk.StringVar()
        mic_names = [f"[{d['index']}] {d['name']}" for d in self._input_devices]
        self._mic_combo = ttk.Combobox(mic_row, textvariable=self._mic_var,
                                        values=mic_names, state="readonly",
                                        style="M.TCombobox")
        self._mic_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if mic_names:
            self._mic_combo.current(0)
            saved = self._prefs.get("mic_index")
            if saved is not None:
                for i, d in enumerate(self._input_devices):
                    if d["index"] == saved:
                        self._mic_combo.current(i)
                        break
        self._mic_combo.bind("<<ComboboxSelected>>", lambda e: self._on_change())

        out_row = tk.Frame(self, bg=styles.BG_PANEL)
        out_row.pack(fill=tk.X)
        tk.Label(out_row, text="System Audio", bg=styles.BG_PANEL,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL, width=12,
                 anchor="w").pack(side=tk.LEFT, padx=(2, 8))

        self._out_var = tk.StringVar()
        out_names = ["[None] — Skip"] + [
            f"[{d['index']}] {d['name']}" for d in self._output_devices]
        self._out_combo = ttk.Combobox(out_row, textvariable=self._out_var,
                                        values=out_names, state="readonly",
                                        style="M.TCombobox")
        self._out_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._out_combo.current(0)
        saved_out = self._prefs.get("out_index")
        if saved_out is not None:
            for i, d in enumerate(self._output_devices):
                if d["index"] == saved_out:
                    self._out_combo.current(i + 1)
                    break
        self._out_combo.bind("<<ComboboxSelected>>", lambda e: self._on_change())

    def _on_change(self):
        _save_prefs(self.get_mic_index(), self.get_output_index())

    def get_mic_index(self) -> Optional[int]:
        val = self._mic_var.get()
        if not val:
            return None
        return int(val.split("]")[0].replace("[", "").strip())

    def get_output_index(self) -> Optional[int]:
        val = self._out_var.get()
        if not val or val.startswith("[None]"):
            return None
        return int(val.split("]")[0].replace("[", "").strip())
