"""
Settings dialog — edit API keys, model config, audio devices, and paths.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from config.settings import Settings
from core.audio_capture import list_input_devices, list_output_devices
from ui import styles


class SettingsDialog(tk.Toplevel):

    def __init__(self, parent, settings: Settings, device_panel=None):
        super().__init__(parent)
        self._settings = settings
        self._device_panel = device_panel
        self._saved = False

        self.title("Settings")
        self.geometry("560x640")
        self.minsize(520, 480)
        self.resizable(True, True)
        self.configure(bg=styles.BG_DARK)
        self.transient(parent)
        self.grab_set()

        # Center on parent
        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - 560) // 2
        py = parent.winfo_y() + (parent.winfo_height() - 640) // 2
        self.geometry(f"+{px}+{py}")

        self._build()

    def _build(self):
        # Top-level container with pinned header + scrollable body + pinned footer
        container = tk.Frame(self, bg=styles.BG_DARK)
        container.pack(fill=tk.BOTH, expand=True)

        # Pinned header (packed first — top)
        header = tk.Frame(container, bg=styles.BG_DARK)
        header.pack(side=tk.TOP, fill=tk.X, padx=20, pady=(16, 8))
        tk.Label(header, text="Settings", bg=styles.BG_DARK,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_HEADER).pack(
                     anchor="w")

        # Pinned footer (packed second — bottom). Contents added later.
        self._footer = tk.Frame(container, bg=styles.BG_DARK)
        self._footer.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(8, 16))

        # Scrollable body (packed last — fills remaining space)
        body_wrap = tk.Frame(container, bg=styles.BG_DARK)
        body_wrap.pack(side=tk.TOP, fill=tk.BOTH, expand=True,
                        padx=16, pady=(0, 6))

        canvas = tk.Canvas(body_wrap, bg=styles.BG_DARK,
                            highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(body_wrap, orient="vertical",
                                    command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        outer = tk.Frame(canvas, bg=styles.BG_DARK)
        canvas_window = canvas.create_window((0, 0), window=outer, anchor="nw")

        def _on_frame_resize(e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_resize(e):
            canvas.itemconfig(canvas_window, width=e.width)

        outer.bind("<Configure>", _on_frame_resize)
        canvas.bind("<Configure>", _on_canvas_resize)

        # Enable mousewheel scrolling
        def _on_mousewheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Keep outer as the parent for all settings sections
        outer.configure(padx=4, pady=4)

        # API Keys section
        self._section(outer, "API Keys")

        self._anthropic_var = tk.StringVar(value=self._settings.anthropic_api_key)
        self._field(outer, "Anthropic API Key", self._anthropic_var, show="*")

        self._hf_var = tk.StringVar(value=self._settings.hf_token)
        self._field(outer, "HuggingFace Token", self._hf_var, show="*")

        # Audio Devices section
        self._section(outer, "Audio Devices")

        self._input_devices = list_input_devices()
        self._output_devices = list_output_devices()

        mic_row = tk.Frame(outer, bg=styles.BG_DARK)
        mic_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(mic_row, text="Microphone", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL,
                 width=16, anchor="w").pack(side=tk.LEFT)
        self._mic_var = tk.StringVar()
        mic_names = [f"[{d['index']}] {d['name']}" for d in self._input_devices]
        self._mic_combo = ttk.Combobox(mic_row, textvariable=self._mic_var,
                                        values=mic_names, state="readonly")
        self._mic_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        # Restore saved selection
        if self._device_panel and mic_names:
            saved_idx = self._device_panel.get_mic_index()
            for i, d in enumerate(self._input_devices):
                if d["index"] == saved_idx:
                    self._mic_combo.current(i)
                    break
            else:
                self._mic_combo.current(0)

        out_row = tk.Frame(outer, bg=styles.BG_DARK)
        out_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(out_row, text="System Audio", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL,
                 width=16, anchor="w").pack(side=tk.LEFT)
        self._out_var = tk.StringVar()
        out_names = ["[None] — Skip"] + [
            f"[{d['index']}] {d['name']}" for d in self._output_devices]
        self._out_combo = ttk.Combobox(out_row, textvariable=self._out_var,
                                        values=out_names, state="readonly")
        self._out_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._out_combo.current(0)
        if self._device_panel:
            saved_out = self._device_panel.get_output_index()
            if saved_out is not None:
                for i, d in enumerate(self._output_devices):
                    if d["index"] == saved_out:
                        self._out_combo.current(i + 1)
                        break

        # Model section
        self._section(outer, "Model Configuration")

        model_row = tk.Frame(outer, bg=styles.BG_DARK)
        model_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(model_row, text="Whisper Model", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL,
                 width=16, anchor="w").pack(side=tk.LEFT)
        self._model_var = tk.StringVar(value=self._settings.whisper_model)
        model_combo = ttk.Combobox(model_row, textvariable=self._model_var,
                                   values=["tiny", "base", "small", "medium", "large"],
                                   state="readonly", width=12)
        model_combo.pack(side=tk.LEFT)

        speakers_row = tk.Frame(outer, bg=styles.BG_DARK)
        speakers_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(speakers_row, text="Max Speakers", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL,
                 width=16, anchor="w").pack(side=tk.LEFT)
        self._speakers_var = tk.IntVar(value=self._settings.max_speakers)
        spinbox = tk.Spinbox(speakers_row, from_=2, to=20,
                             textvariable=self._speakers_var, width=5,
                             bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,
                             font=styles.FONT_BODY, relief=tk.FLAT,
                             highlightbackground=styles.BORDER, highlightthickness=1)
        spinbox.pack(side=tk.LEFT)

        claude_row = tk.Frame(outer, bg=styles.BG_DARK)
        claude_row.pack(fill=tk.X, pady=(0, 4))
        tk.Label(claude_row, text="Claude Model", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL,
                 width=16, anchor="w").pack(side=tk.LEFT)
        self._claude_var = tk.StringVar(value=self._settings.claude_model)
        claude_combo = ttk.Combobox(
            claude_row, textvariable=self._claude_var,
            values=["claude-haiku-4-5", "claude-sonnet-4-5",
                    "claude-3-5-haiku-latest"],
            state="readonly", width=28)
        claude_combo.pack(side=tk.LEFT)
        tk.Label(outer, text="Haiku 4.5 is ~4x cheaper than Sonnet. "
                 "Good for summaries.",
                 bg=styles.BG_DARK, fg=styles.TEXT_HINT,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=(120, 0))

        # Paths section
        self._section(outer, "Storage")

        dir_row = tk.Frame(outer, bg=styles.BG_DARK)
        dir_row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(dir_row, text="Recordings Folder", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL,
                 width=16, anchor="w").pack(side=tk.LEFT)
        self._dir_var = tk.StringVar(value=self._settings.recordings_dir)
        tk.Entry(dir_row, textvariable=self._dir_var,
                 bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,
                 font=styles.FONT_BODY, relief=tk.FLAT,
                 highlightbackground=styles.BORDER, highlightthickness=1
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        tk.Button(dir_row, text="Browse", bg=styles.BG_PANEL,
                  fg=styles.ACCENT, font=styles.FONT_SMALL,
                  relief=tk.FLAT, cursor="hand2",
                  command=self._browse_dir).pack(side=tk.LEFT, padx=(6, 0))

        # Email section
        self._section(outer, "Email (Outlook)")

        self._email_var = tk.StringVar(value=self._settings.email_to)
        self._field(outer, "Send To", self._email_var)

        tk.Label(outer, text="Leave blank to send to yourself. Requires Outlook.",
                 bg=styles.BG_DARK, fg=styles.TEXT_HINT,
                 font=("Segoe UI", 9)).pack(anchor="w")

        # Calendar section
        self._section(outer, "Calendar")

        notify_row = tk.Frame(outer, bg=styles.BG_DARK)
        notify_row.pack(fill=tk.X, pady=(0, 4))
        tk.Label(notify_row, text="Notify Before", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL,
                 width=16, anchor="w").pack(side=tk.LEFT)
        self._notify_var = tk.IntVar(value=self._settings.notify_minutes_before)
        notify_spin = tk.Spinbox(notify_row, from_=0, to=30,
                                  textvariable=self._notify_var, width=5,
                                  bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,
                                  font=styles.FONT_BODY, relief=tk.FLAT,
                                  highlightbackground=styles.BORDER,
                                  highlightthickness=1)
        notify_spin.pack(side=tk.LEFT)
        tk.Label(notify_row, text="minutes before meeting starts (0 = off)",
                 bg=styles.BG_DARK, fg=styles.TEXT_HINT,
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(8, 0))

        # Workflow section
        self._section(outer, "Workflow")

        auto_row = tk.Frame(outer, bg=styles.BG_DARK)
        auto_row.pack(fill=tk.X, pady=(0, 4))
        self._auto_process_var = tk.BooleanVar(
            value=self._settings.auto_process_after_stop)
        auto_check = tk.Checkbutton(
            auto_row, text="Auto-process after recording stops",
            variable=self._auto_process_var,
            bg=styles.BG_DARK, fg=styles.TEXT_PRIMARY,
            font=styles.FONT_BODY, activebackground=styles.BG_DARK,
            activeforeground=styles.TEXT_PRIMARY,
            selectcolor=styles.BG_INPUT, bd=0,
            highlightthickness=0)
        auto_check.pack(side=tk.LEFT, anchor="w")
        tk.Label(outer,
                 text="Runs transcribe → summarize → action items → "
                      "requirements automatically.",
                 bg=styles.BG_DARK, fg=styles.TEXT_HINT,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=(28, 0))

        # Populate the pinned footer with Save/Cancel (always visible)
        tk.Button(self._footer, text="Save", bg=styles.ACCENT, fg="#ffffff",
                  font=styles.FONT_BODY, relief=tk.FLAT, padx=20, pady=6,
                  cursor="hand2", command=self._save).pack(side=tk.RIGHT)
        tk.Button(self._footer, text="Cancel", bg=styles.BG_PANEL,
                  fg=styles.TEXT_MUTED, font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=20, pady=6, cursor="hand2",
                  command=self.destroy).pack(side=tk.RIGHT, padx=(0, 8))

    def _section(self, parent, title):
        tk.Label(parent, text=title, bg=styles.BG_DARK,
                 fg=styles.ACCENT, font=styles.FONT_TITLE).pack(
                     anchor="w", pady=(12, 4))

    def _field(self, parent, label, var, show=""):
        row = tk.Frame(parent, bg=styles.BG_DARK)
        row.pack(fill=tk.X, pady=(0, 8))
        tk.Label(row, text=label, bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL,
                 width=16, anchor="w").pack(side=tk.LEFT)
        tk.Entry(row, textvariable=var, show=show,
                 bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,
                 font=styles.FONT_BODY, relief=tk.FLAT,
                 highlightbackground=styles.BORDER, highlightthickness=1
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)

    def _browse_dir(self):
        path = filedialog.askdirectory(initialdir=self._dir_var.get())
        if path:
            self._dir_var.set(path)

    def _save(self):
        api_key = self._anthropic_var.get().strip()
        hf_token = self._hf_var.get().strip()
        if not api_key or not hf_token:
            messagebox.showwarning("Missing Keys",
                                   "API Key and HuggingFace Token are required.",
                                   parent=self)
            return
        try:
            Settings.save_to_env(
                anthropic_api_key=api_key,
                hf_token=hf_token,
                whisper_model=self._model_var.get(),
                max_speakers=self._speakers_var.get(),
                recordings_dir=self._dir_var.get().strip() or "recordings",
                email_to=self._email_var.get().strip(),
                claude_model=self._claude_var.get(),
                notify_minutes_before=self._notify_var.get(),
                auto_process_after_stop=self._auto_process_var.get(),
            )
            # Update device selections
            if self._device_panel:
                mic_val = self._mic_var.get()
                out_val = self._out_var.get()
                if mic_val:
                    self._device_panel._mic_var.set(mic_val)
                if out_val:
                    self._device_panel._out_var.set(out_val)
                self._device_panel._on_change()
            self._saved = True
            messagebox.showinfo("Settings Saved",
                                "Settings saved. Restart the app for "
                                "API/model changes to take effect.",
                                parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {e}", parent=self)

    @property
    def saved(self) -> bool:
        return self._saved
