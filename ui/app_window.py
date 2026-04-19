"""
Main application window — Material You dark theme.
"""

import asyncio
import datetime
import os
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from typing import Optional
import uuid

from config.settings import Settings
from core.diarization import DiarizationEngine
from core.summarizer import Summarizer, MEETING_TEMPLATES
from core.transcription import TranscriptionEngine
from models.session import Session
from services.calendar_monitor import CalendarMonitor
from services.calendar_service import get_todays_meetings, is_outlook_available
from services.export_service import ExportService
from services.recording_service import RecordingService
from services.session_service import SessionService
from ui import styles
from ui.calendar_panel import CalendarPanel
from ui.device_panel import DevicePanel
from ui.follow_up_tracker import FollowUpTracker
from ui.prep_brief_dialog import PrepBriefDialog
from ui.session_browser import SessionBrowser
from ui.settings_dialog import SettingsDialog
from ui.speaker_panel import SpeakerPanel
from ui.transcript_panel import TranscriptPanel
from utils.logger import get_logger

logger = get_logger(__name__)


def _run_async_in_thread(coro_factory, on_success, on_error):
    def _worker():
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            result = loop.run_until_complete(coro_factory())
            on_success(result)
        except Exception as e:
            logger.exception("Background async task failed")
            on_error(e)
        finally:
            loop.close()
    threading.Thread(target=_worker, daemon=True).start()


class AppWindow(tk.Tk):

    def __init__(self, settings: Settings):
        super().__init__()
        self._settings = settings
        self._session: Optional[Session] = None
        self._active_meeting: Optional[dict] = None
        self._progress_after = None
        self._models_ready = False

        self._transcription: Optional[TranscriptionEngine] = None
        self._diarization: Optional[DiarizationEngine] = None
        self._summarizer    = Summarizer(settings.anthropic_api_key, model=settings.claude_model) if settings.anthropic_api_key else None
        self._session_svc   = SessionService(settings.recordings_dir)
        self._export_svc    = ExportService(settings.recordings_dir)

        # Create recording service immediately — recording doesn't need AI models
        self._recording_svc = RecordingService(
            settings=settings,
            on_status=self._thread_safe_status,
        )

        self._build_window()
        self._build_layout()

        # Recording is available immediately
        self._set_status("Ready to record")
        # Load ML models in background (for processing step)
        if self._settings.is_configured:
            threading.Thread(target=self._load_models, daemon=True).start()
        else:
            self._on_not_configured()

        # Load calendar + start monitor (both fail gracefully if no Outlook)
        self._calendar_monitor: Optional[CalendarMonitor] = None
        threading.Thread(target=self._load_calendar, daemon=True).start()
        if settings.notify_minutes_before > 0:
            self._calendar_monitor = CalendarMonitor(
                on_upcoming=lambda m: self.after(0, self._on_meeting_upcoming, m),
                notify_minutes_before=settings.notify_minutes_before,
            )
            self._calendar_monitor.start()

    def _build_window(self) -> None:
        self.title("Meeting Recorder")
        self.geometry("920x820")
        self.minsize(800, 680)
        self.configure(bg=styles.BG_DARK)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        try:
            self.iconbitmap("meeting_recorder.ico")
        except Exception:
            pass
        self._build_menu()

    def _build_menu(self) -> None:
        menubar = tk.Menu(self, bg=styles.BG_PANEL, fg=styles.TEXT_PRIMARY,
                          activebackground=styles.ACCENT_BG,
                          activeforeground=styles.ACCENT, relief=tk.FLAT)

        file_menu = tk.Menu(menubar, tearoff=0, bg=styles.BG_PANEL,
                            fg=styles.TEXT_PRIMARY,
                            activebackground=styles.ACCENT_BG,
                            activeforeground=styles.ACCENT)
        file_menu.add_command(label="Settings...", command=self._open_settings)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        sessions_menu = tk.Menu(menubar, tearoff=0, bg=styles.BG_PANEL,
                                 fg=styles.TEXT_PRIMARY,
                                 activebackground=styles.ACCENT_BG,
                                 activeforeground=styles.ACCENT)
        sessions_menu.add_command(label="Session History...",
                                   command=self._open_session_history)
        sessions_menu.add_command(label="Follow-Up Tracker...",
                                   command=self._open_follow_up_tracker)
        sessions_menu.add_command(label="Meeting Prep Brief...",
                                   command=self._open_prep_brief)
        menubar.add_cascade(label="Sessions", menu=sessions_menu)

        help_menu = tk.Menu(menubar, tearoff=0, bg=styles.BG_PANEL,
                            fg=styles.TEXT_PRIMARY,
                            activebackground=styles.ACCENT_BG,
                            activeforeground=styles.ACCENT)
        help_menu.add_command(label="Open Logs Folder", command=self._open_logs_folder)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menubar)

    def _build_layout(self) -> None:
        outer = tk.Frame(self, bg=styles.BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=styles.PAD_LG, pady=styles.PAD_LG)

        # ── Top bar ──────────────────────────────────────────────────
        topbar = tk.Frame(outer, bg=styles.ACCENT)
        topbar.pack(fill=tk.X, pady=(0, styles.PAD), ipady=10)

        icon_frame = tk.Frame(topbar, bg=styles.ACCENT, width=42, height=42)
        icon_frame.pack(side=tk.LEFT, padx=(12, 0))
        icon_frame.pack_propagate(False)
        tk.Label(icon_frame, text="🎙", bg=styles.ACCENT,
                 font=(styles.FONT_FAMILY, 16)).place(relx=0.5, rely=0.5, anchor="center")

        title_block = tk.Frame(topbar, bg=styles.ACCENT)
        title_block.pack(side=tk.LEFT, padx=(10, 0))
        tk.Label(title_block, text="Meeting Recorder", bg=styles.ACCENT,
                 fg="#ffffff", font=styles.FONT_HEADER).pack(anchor="w")

        self._status_var = tk.StringVar(value="Ready")
        tk.Label(topbar, textvariable=self._status_var,
                 bg=styles.ACCENT, fg="#e3f2fd",
                 font=styles.FONT_SMALL).pack(side=tk.RIGHT, padx=12, anchor="s")

        # ── Meeting info card ─────────────────────────────────────────
        info_card = tk.Frame(outer, bg=styles.BG_PANEL,
                              highlightbackground=styles.BORDER,
                              highlightthickness=1)
        info_card.pack(fill=tk.X, pady=(0, styles.PAD))

        # Meeting name row
        name_row = tk.Frame(info_card, bg=styles.BG_PANEL)
        name_row.pack(fill=tk.X, padx=14, pady=(10, 6))
        tk.Label(name_row, text="Meeting", bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=styles.FONT_SMALL,
                 width=10, anchor="w").pack(side=tk.LEFT)
        self._meeting_name_var = tk.StringVar(
            value=datetime.datetime.now().strftime("%Y-%m-%d Meeting"))
        self._name_entry = tk.Entry(
            name_row,
            textvariable=self._meeting_name_var,
            bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,
            insertbackground=styles.TEXT_PRIMARY,
            font=styles.FONT_BODY, relief=tk.FLAT,
            highlightbackground=styles.BORDER, highlightthickness=1,
        )
        self._name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        self._meeting_name_var.trace_add("write", self._on_name_change)

        # Template + tags row
        template_row = tk.Frame(info_card, bg=styles.BG_PANEL)
        template_row.pack(fill=tk.X, padx=14, pady=(0, 6))
        tk.Label(template_row, text="Template", bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=styles.FONT_SMALL,
                 width=10, anchor="w").pack(side=tk.LEFT)
        self._template_var = tk.StringVar(value="General")
        template_combo = ttk.Combobox(
            template_row, textvariable=self._template_var,
            values=list(MEETING_TEMPLATES.keys()),
            state="readonly", width=24)
        template_combo.pack(side=tk.LEFT)

        # Client/Project tags row
        tag_row = tk.Frame(info_card, bg=styles.BG_PANEL)
        tag_row.pack(fill=tk.X, padx=14, pady=(0, 10))
        tk.Label(tag_row, text="Client", bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=styles.FONT_SMALL,
                 width=10, anchor="w").pack(side=tk.LEFT)
        self._client_var = tk.StringVar()
        self._client_combo = ttk.Combobox(
            tag_row, textvariable=self._client_var,
            values=self._gather_existing("client"), width=20)
        self._client_combo.pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(tag_row, text="Project", bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=styles.FONT_SMALL,
                 width=8, anchor="w").pack(side=tk.LEFT)
        self._project_var = tk.StringVar()
        self._project_combo = ttk.Combobox(
            tag_row, textvariable=self._project_var,
            values=self._gather_existing("project"), width=24)
        self._project_combo.pack(side=tk.LEFT)

        # Device panel (lives in settings dialog, created here for access)
        self._device_panel = DevicePanel(self)

        # ── Progress stages (compact) ─────────────────────────────────
        prog_row = tk.Frame(outer, bg=styles.BG_DARK)
        prog_row.pack(fill=tk.X, pady=(0, 6))

        self._stage_labels = {}
        stages = [
            ("transcribe", "Transcribe"),
            ("diarize",    "Diarize"),
            ("speakers",   "Speakers"),
            ("complete",   "Complete"),
        ]
        for key, label in stages:
            dot = tk.Label(prog_row, text="○", bg=styles.BG_DARK,
                           fg=styles.TEXT_HINT, font=(styles.FONT_FAMILY, 10))
            dot.pack(side=tk.LEFT)
            lbl = tk.Label(prog_row, text=label, bg=styles.BG_DARK,
                           fg=styles.TEXT_HINT, font=styles.FONT_SMALL)
            lbl.pack(side=tk.LEFT, padx=(0, 16))
            self._stage_labels[key] = (dot, lbl)

        # ── Action row 1 ──────────────────────────────────────────────
        row1 = tk.Frame(outer, bg=styles.BG_DARK)
        row1.pack(fill=tk.X, pady=(0, 8))

        self._rec_btn = self._pill_button(
            row1, "⏺  Start Recording", styles.DANGER, self._toggle_recording)
        self._rec_btn.pack(side=tk.LEFT, padx=(0, 8))

        self._load_btn = self._pill_button(
            row1, "📂  Load File", styles.ACCENT, self._load_audio_file)
        self._load_btn.pack(side=tk.LEFT, padx=(0, 8))

        self._process_btn = self._pill_button(
            row1, "⚙  Process", styles.ACCENT_DIM, self._process, outline=True)
        self._process_btn.pack(side=tk.LEFT)
        self._process_btn.config(state=tk.DISABLED)

        # ── AI actions row ─────────────────────────────────────────────
        row2 = tk.Frame(outer, bg=styles.BG_DARK)
        row2.pack(fill=tk.X, pady=(0, 8))

        self._summarize_btn = self._pill_button(
            row2, "✨  Summarize", styles.BG_INPUT, self._summarize, outline=True)
        self._summarize_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._summarize_btn.config(state=tk.DISABLED)

        self._action_items_btn = self._pill_button(
            row2, "📋  Action Items", styles.BG_INPUT, self._extract_action_items, outline=True)
        self._action_items_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._action_items_btn.config(state=tk.DISABLED)

        self._requirements_btn = self._pill_button(
            row2, "📝  Requirements", styles.BG_INPUT, self._extract_requirements, outline=True)
        self._requirements_btn.pack(side=tk.LEFT)
        self._requirements_btn.config(state=tk.DISABLED)

        # ── Export row ────────────────────────────────────────────────
        row3 = tk.Frame(outer, bg=styles.BG_DARK)
        row3.pack(fill=tk.X, pady=(0, styles.PAD))

        self._export_btn = self._pill_button(
            row3, "💾  Export", styles.BG_INPUT, self._export, outline=True)
        self._export_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._export_btn.config(state=tk.DISABLED)

        self._email_btn = self._pill_button(
            row3, "✉  Email", styles.BG_INPUT, self._email_summary, outline=True)
        self._email_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._email_btn.config(state=tk.DISABLED)

        self._folder_btn = self._pill_button(
            row3, "📁  Recordings", styles.BG_INPUT, self._open_recordings, outline=True)
        self._folder_btn.pack(side=tk.LEFT)

        # ── Calendar card (collapsible) ───────────────────────────────
        cal_card = tk.Frame(outer, bg=styles.BG_PANEL,
                             highlightbackground=styles.BORDER,
                             highlightthickness=1)
        cal_card.pack(fill=tk.X, pady=(0, styles.PAD))

        cal_header = tk.Frame(cal_card, bg=styles.BG_PANEL)
        cal_header.pack(fill=tk.X, padx=14, pady=(8, 0))
        tk.Label(cal_header, text="TODAY'S MEETINGS",
                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,
                 font=styles.FONT_SMALL).pack(side=tk.LEFT)
        tk.Button(cal_header, text="Refresh",
                  bg=styles.BG_PANEL, fg=styles.ACCENT,
                  font=styles.FONT_SMALL, relief=tk.FLAT,
                  cursor="hand2", bd=0,
                  command=self._refresh_calendar).pack(side=tk.RIGHT)
        self._cal_toggle_var = tk.StringVar(value="Hide")
        self._cal_toggle = tk.Button(
            cal_header, textvariable=self._cal_toggle_var,
            bg=styles.BG_PANEL, fg=styles.ACCENT, font=styles.FONT_SMALL,
            relief=tk.FLAT, cursor="hand2", bd=0,
            command=self._toggle_calendar)
        self._cal_toggle.pack(side=tk.RIGHT, padx=(0, 8))

        self._cal_body = tk.Frame(cal_card, bg=styles.BG_PANEL)
        self._cal_body.pack(fill=tk.X, padx=8, pady=(0, 8))
        self._cal_expanded = True
        self._calendar_panel = CalendarPanel(
            self._cal_body, on_record=self._start_from_meeting)
        self._calendar_panel.pack(fill=tk.X)

        # ── Speaker card (collapsible) ────────────────────────────────
        speaker_card = tk.Frame(outer, bg=styles.BG_PANEL,
                                 highlightbackground=styles.BORDER,
                                 highlightthickness=1)
        speaker_card.pack(fill=tk.X, pady=(0, styles.PAD))

        speaker_header = tk.Frame(speaker_card, bg=styles.BG_PANEL)
        speaker_header.pack(fill=tk.X, padx=14, pady=(8, 0))
        tk.Label(speaker_header, text="SPEAKERS",
                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,
                 font=styles.FONT_SMALL).pack(side=tk.LEFT)
        self._speaker_toggle_var = tk.StringVar(value="Show")
        self._speaker_toggle = tk.Button(
            speaker_header, textvariable=self._speaker_toggle_var,
            bg=styles.BG_PANEL, fg=styles.ACCENT, font=styles.FONT_SMALL,
            relief=tk.FLAT, cursor="hand2", bd=0,
            command=self._toggle_speakers)
        self._speaker_toggle.pack(side=tk.RIGHT)

        self._speaker_body = tk.Frame(speaker_card, bg=styles.BG_PANEL)
        # Start collapsed
        self._speaker_expanded = False
        self._speaker_panel = SpeakerPanel(self._speaker_body, on_rename=self._on_rename_speaker)
        self._speaker_panel.pack(fill=tk.X, padx=8, pady=(0, 8))

        # ── Transcript card ───────────────────────────────────────────
        transcript_card = tk.Frame(outer, bg=styles.BG_PANEL,
                                    highlightbackground=styles.BORDER,
                                    highlightthickness=1)
        transcript_card.pack(fill=tk.BOTH, expand=True)
        tk.Label(transcript_card, text="TRANSCRIPT",
                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,
                 font=styles.FONT_SMALL).pack(anchor="w", padx=14, pady=(10, 4))
        self._transcript_panel = TranscriptPanel(transcript_card)
        self._transcript_panel.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

    # ------------------------------------------------------------------ #
    # Model loading (background)
    # ------------------------------------------------------------------ #

    def _load_models(self) -> None:
        try:
            logger.info("Loading transcription engine...")
            self._transcription = TranscriptionEngine(self._settings.whisper_model)
            logger.info("Loading diarization engine...")
            self._diarization = DiarizationEngine(
                self._settings.hf_token, self._settings.max_speakers)
            # Attach engines to the existing recording service
            self._recording_svc.set_engines(
                transcription_engine=self._transcription,
                diarization_engine=self._diarization,
            )
            self.after(0, self._on_models_ready)
        except Exception as e:
            logger.error(f"Failed to load models: {e}")
            err_msg = str(e)
            self.after(0, lambda: self._on_model_load_failed(err_msg))

    def _on_model_load_failed(self, err_msg: str) -> None:
        self._set_status("Ready to record (processing unavailable)")
        hint = ""
        if "401" in err_msg or "403" in err_msg or "token" in err_msg.lower():
            hint = (
                "\n\nThis usually means:\n"
                "  • HuggingFace token is invalid or missing\n"
                "  • You haven't accepted the pyannote model terms:\n"
                "    - huggingface.co/pyannote/speaker-diarization-3.1\n"
                "    - huggingface.co/pyannote/segmentation-3.0\n\n"
                "Fix via File > Settings, then restart the app.")
        elif "cuda" in err_msg.lower() or "gpu" in err_msg.lower():
            hint = (
                "\n\nGPU/CUDA issue. The app needs an NVIDIA GPU with "
                "CUDA support.\nCheck that nvidia-smi works in a terminal.")
        messagebox.showerror(
            "Model Load Failed",
            f"Could not load AI models:\n\n{err_msg}{hint}")

    def _on_not_configured(self) -> None:
        self._set_status("Ready to record (API keys needed for processing)")
        logger.info("No API keys — recording enabled, processing disabled.")

    def _on_models_ready(self) -> None:
        self._models_ready = True
        self._rec_btn.config(state=tk.NORMAL)
        self._load_btn.config(state=tk.NORMAL)
        self._set_status("Ready")
        logger.info("All models loaded. Ready to record.")

    # ------------------------------------------------------------------ #
    # Meeting name
    # ------------------------------------------------------------------ #

    def _on_name_change(self, *args) -> None:
        if self._session:
            self._session.display_name = self._meeting_name_var.get().strip()

    def _get_meeting_name(self) -> str:
        name = self._meeting_name_var.get().strip()
        if not name:
            name = datetime.datetime.now().strftime("%Y-%m-%d Meeting")
        return name

    def _sync_tags_to_session(self) -> None:
        """Copy client/project values from UI entries into the current session."""
        if self._session:
            self._session.client = self._client_var.get().strip()
            self._session.project = self._project_var.get().strip()

    # ------------------------------------------------------------------ #
    # Progress stages
    # ------------------------------------------------------------------ #

    def _set_stage(self, stage: str, state: str) -> None:
        dot, lbl = self._stage_labels[stage]
        if state == "active":
            dot.config(text="●", fg=styles.ACCENT)
            lbl.config(fg=styles.ACCENT)
            self._animate_dot(stage)
        elif state == "done":
            dot.config(text="✓", fg=styles.SUCCESS_DIM)
            lbl.config(fg=styles.SUCCESS_DIM)
        else:
            dot.config(text="○", fg=styles.TEXT_HINT)
            lbl.config(fg=styles.TEXT_HINT)

    def _animate_dot(self, stage: str) -> None:
        frames = ["●", "○", "●", "○"]
        def _tick(i=0):
            if stage not in self._stage_labels:
                return
            dot, _ = self._stage_labels[stage]
            if dot.cget("text") == "✓":
                return
            if dot.cget("fg") not in (styles.ACCENT, styles.TEXT_HINT):
                return
            dot.config(text=frames[i % len(frames)])
            self._progress_after = self.after(500, lambda: _tick(i + 1))
        _tick()

    def _reset_stages(self) -> None:
        for key in self._stage_labels:
            self._set_stage(key, "pending")

    # ------------------------------------------------------------------ #
    # Recording
    # ------------------------------------------------------------------ #

    def _toggle_recording(self) -> None:
        if not self._recording_svc:
            return
        if not self._recording_svc.is_recording:
            self._active_meeting = None
            mic_idx = self._device_panel.get_mic_index()
            out_idx = self._device_panel.get_output_index()
            try:
                self._session = self._recording_svc.start_recording(mic_idx, out_idx)
                self._session.display_name = self._get_meeting_name()
            except Exception as e:
                messagebox.showerror("Recording Error", str(e))
                return
            self._reset_stages()
            self._transcript_panel.clear()
            self._rec_btn.config(text="⏹  Stop Recording",
                                  bg=styles.DANGER_DIM, fg="#ffffff")
            self._load_btn.config(state=tk.DISABLED)
            self._process_btn.config(state=tk.DISABLED)
            self._summarize_btn.config(state=tk.DISABLED)
            self._action_items_btn.config(state=tk.DISABLED)
            self._requirements_btn.config(state=tk.DISABLED)
            self._export_btn.config(state=tk.DISABLED)
            self._email_btn.config(state=tk.DISABLED)
            self._set_status("Recording...")
        else:
            self._session = self._recording_svc.stop_recording()
            self._rec_btn.config(text="⏺  Start Recording",
                                  bg=styles.DANGER, fg="#ffffff")
            self._load_btn.config(state=tk.NORMAL)
            if self._session:
                self._session.template = self._template_var.get()
                self._sync_tags_to_session()
            if self._session and self._session.audio_path:
                self._process_btn.config(state=tk.NORMAL,
                                          bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY)
                # Persist the session record immediately so client/project tags are saved
                # even if the user never processes it
                try:
                    self._session_svc.save(self._session)
                except Exception as e:
                    logger.warning(f"Could not save session metadata: {e}")
                # Kick off auto-process if enabled
                if self._settings.auto_process_after_stop and self._models_ready:
                    self._set_status("Auto-processing...")
                    self._process()
            self._set_status("Recording saved. Ready to process.")

    # ------------------------------------------------------------------ #
    # Load file
    # ------------------------------------------------------------------ #

    def _load_audio_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Audio File",
            filetypes=[
                ("Audio Files", "*.wav *.mp3 *.m4a *.flac *.ogg *.aac"),
                ("All files", "*.*"),
            ]
        )
        if not file_path:
            return
        session_id = uuid.uuid4().hex[:8].upper()
        self._session = Session(session_id=session_id)
        self._session.started_at = datetime.datetime.now()
        self._session.ended_at   = datetime.datetime.now()
        self._session.audio_path = file_path
        self._session.display_name = self._get_meeting_name()
        self._active_meeting = None
        self._recording_svc.set_session(self._session)
        self._reset_stages()
        self._transcript_panel.clear()
        self._summarize_btn.config(state=tk.DISABLED)
        self._export_btn.config(state=tk.DISABLED)
        self._email_btn.config(state=tk.DISABLED)
        self._process_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                  fg=styles.TEXT_PRIMARY)
        self._set_status(f"Loaded: {os.path.basename(file_path)}")

    # ------------------------------------------------------------------ #
    # Processing
    # ------------------------------------------------------------------ #

    def _process(self) -> None:
        if not self._session or not self._session.audio_path:
            messagebox.showwarning("No Audio", "Please record or load an audio file first.")
            return
        self._process_btn.config(state=tk.DISABLED, bg=styles.BG_INPUT,
                                  fg=styles.TEXT_MUTED)
        self._reset_stages()
        self._set_stage("transcribe", "active")
        self._set_status("Transcribing audio...")
        recording_svc = self._recording_svc

        def _coro():
            return recording_svc.process_session()

        def _on_success(result):
            self.after(0, lambda: self._on_process_complete(result))

        def _on_error(e):
            self.after(0, lambda: messagebox.showerror("Processing Error", str(e)))
            self.after(0, lambda: self._set_status("Processing failed."))
            self.after(0, lambda: self._reset_stages())
            self.after(0, lambda: self._process_btn.config(
                state=tk.NORMAL, bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY))

        _run_async_in_thread(_coro, _on_success, _on_error)

    def _on_process_complete(self, session: Session) -> None:
        self._session = session
        self._session.display_name = self._get_meeting_name()
        self._sync_tags_to_session()
        self._set_stage("transcribe", "done")
        self._set_stage("diarize", "done")
        self._set_stage("speakers", "active")
        self._set_status("Identifying speakers...")
        self._speaker_panel.populate(session)
        self._transcript_panel.set_text(session.full_transcript())
        self._summarize_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                    fg=styles.TEXT_PRIMARY)
        self._action_items_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                       fg=styles.TEXT_PRIMARY)
        self._requirements_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                       fg=styles.TEXT_PRIMARY)
        self._export_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                 fg=styles.TEXT_PRIMARY)
        try:
            self._session_svc.save(session)
            self._export_svc.export_transcript(session)
        except Exception as e:
            logger.error(f"Failed to save session: {e}")
        self._auto_identify_speakers(session)

        # Auto-chain: summarize → action items → requirements
        if self._settings.auto_process_after_stop and self._summarizer:
            self.after(500, self._auto_chain_summarize)

    def _auto_chain_summarize(self) -> None:
        """Step 1 of auto-chain: summarize, then trigger action items."""
        if not self._session or not self._session.segments:
            return
        self._set_status("Auto: generating summary...")
        transcript = self._session.full_transcript()
        session_ref = self._session
        template = self._template_var.get()
        session_ref.template = template

        def _coro():
            return self._summarizer.summarize(transcript, template=template)

        def _ok(result):
            def _apply():
                session_ref.summary = result
                self._update_transcript_display()
                try:
                    self._export_svc.export_summary(session_ref)
                    self._session_svc.save(session_ref)
                except Exception:
                    pass
                self._email_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                        fg=styles.TEXT_PRIMARY)
                self.after(200, self._auto_chain_action_items)
            self.after(0, _apply)

        def _err(e):
            self.after(0, lambda: self._set_status(f"Auto-summary failed: {e}"))

        _run_async_in_thread(_coro, _ok, _err)

    def _auto_chain_action_items(self) -> None:
        """Step 2: extract action items, then trigger requirements."""
        self._set_status("Auto: extracting action items...")
        transcript = self._session.full_transcript()
        session_ref = self._session

        def _coro():
            return self._summarizer.extract_action_items(transcript)

        def _ok(result):
            def _apply():
                session_ref.action_items = result
                self._update_transcript_display()
                try:
                    self._export_svc.export_action_items(session_ref)
                    self._session_svc.save(session_ref)
                except Exception:
                    pass
                self.after(200, self._auto_chain_requirements)
            self.after(0, _apply)

        def _err(e):
            self.after(0, lambda: self._set_status(f"Auto-action-items failed: {e}"))

        _run_async_in_thread(_coro, _ok, _err)

    def _auto_chain_requirements(self) -> None:
        """Step 3: extract requirements, then done."""
        self._set_status("Auto: extracting requirements...")
        transcript = self._session.full_transcript()
        session_ref = self._session

        def _coro():
            return self._summarizer.extract_requirements(transcript)

        def _ok(result):
            def _apply():
                session_ref.requirements = result
                self._update_transcript_display()
                try:
                    self._export_svc.export_requirements(session_ref)
                    self._session_svc.save(session_ref)
                except Exception:
                    pass
                self._set_status("Auto-process complete.")
            self.after(0, _apply)

        def _err(e):
            self.after(0, lambda: self._set_status(f"Auto-requirements failed: {e}"))

        _run_async_in_thread(_coro, _ok, _err)

    def _auto_identify_speakers(self, session: Session) -> None:
        transcript = session.full_transcript()
        summarizer = self._summarizer

        def _coro():
            return summarizer.identify_speakers(transcript)

        def _on_success(suggestions: dict):
            self.after(0, lambda: self._show_speaker_suggestions(suggestions, session))

        def _on_error(e):
            logger.warning(f"Speaker identification failed: {e}")
            self.after(0, lambda: self._set_stage("speakers", "done"))
            self.after(0, lambda: self._set_stage("complete", "done"))
            self.after(0, lambda: self._set_status("Processing complete."))

        _run_async_in_thread(_coro, _on_success, _on_error)

    def _show_speaker_suggestions(self, suggestions: dict, session: Session) -> None:
        self._set_stage("speakers", "done")
        self._set_stage("complete", "done")
        self._set_status("Processing complete.")

        if not suggestions:
            return

        overlay = tk.Toplevel(self)
        overlay.title("Speaker Names Detected")
        overlay.configure(bg=styles.BG_PANEL)
        overlay.resizable(False, False)
        overlay.transient(self)

        overlay.update_idletasks()
        x = self.winfo_x() + (self.winfo_width()  // 2) - 200
        y = self.winfo_y() + (self.winfo_height() // 2) - 150
        overlay.geometry(f"400x{80 + len(suggestions) * 36 + 80}+{x}+{y}")

        tk.Label(overlay, text="Names detected from introductions",
                 bg=styles.BG_PANEL, fg=styles.TEXT_PRIMARY,
                 font=("Segoe UI", 12, "bold")).pack(pady=(16, 4))
        tk.Label(overlay,
                 text="Auto-applying in 5 seconds — click Apply to confirm now",
                 bg=styles.BG_PANEL, fg=styles.TEXT_MUTED,
                 font=styles.FONT_SMALL).pack(pady=(0, 12))

        for speaker_id, name in suggestions.items():
            row = tk.Frame(overlay, bg=styles.BG_INPUT, pady=6, padx=12)
            row.pack(fill=tk.X, padx=16, pady=2)
            tk.Label(row, text=speaker_id, bg=styles.ACCENT_BG,
                     fg=styles.ACCENT, font=styles.FONT_SMALL,
                     padx=8, pady=2).pack(side=tk.LEFT)
            tk.Label(row, text="→", bg=styles.BG_INPUT,
                     fg=styles.TEXT_MUTED, font=styles.FONT_BODY).pack(side=tk.LEFT, padx=8)
            tk.Label(row, text=name, bg=styles.BG_INPUT,
                     fg=styles.TEXT_PRIMARY, font=styles.FONT_BODY).pack(side=tk.LEFT)

        countdown_var = tk.StringVar(value="Apply (5)")
        apply_btn = tk.Button(
            overlay, textvariable=countdown_var,
            bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY,
            font=styles.FONT_BODY, relief=tk.FLAT, padx=20, pady=8,
            cursor="hand2",
        )
        apply_btn.pack(pady=16)

        applied = [False]

        def _apply():
            if applied[0]:
                return
            applied[0] = True
            for speaker_id, name in suggestions.items():
                session.rename_speaker(speaker_id, name)
            self._speaker_panel.populate(session)
            self._transcript_panel.set_text(session.full_transcript())
            try:
                self._session_svc.save(session)
                self._export_svc.export_transcript(session)
            except Exception as e:
                logger.error(f"Failed to save after speaker rename: {e}")
            overlay.destroy()

        apply_btn.config(command=_apply)

        def _countdown(n=5):
            if applied[0]:
                return
            if n <= 0:
                _apply()
                return
            countdown_var.set(f"Apply ({n})")
            overlay.after(1000, lambda: _countdown(n - 1))

        _countdown()

        tk.Button(overlay, text="Skip", bg=styles.BG_PANEL,
                  fg=styles.TEXT_MUTED, font=styles.FONT_SMALL,
                  relief=tk.FLAT, cursor="hand2",
                  command=overlay.destroy).pack()

    # ------------------------------------------------------------------ #
    # Summarize
    # ------------------------------------------------------------------ #

    def _summarize(self) -> None:
        if not self._summarizer:
            messagebox.showwarning("API Key Required",
                "Anthropic API key required for summarization.\n"
                "Add it in File > Settings.")
            return
        if not self._session or not self._session.segments:
            messagebox.showwarning("No Transcript", "Please process a recording first.")
            return
        transcript_snapshot = self._session.full_transcript()
        session_ref = self._session
        template = self._template_var.get()
        session_ref.template = template
        self._summarize_btn.config(state=tk.DISABLED, bg=styles.BG_INPUT,
                                    fg=styles.TEXT_MUTED)
        self._set_status(f"Generating AI summary ({template})...")
        summarizer = self._summarizer

        def _coro():
            return summarizer.summarize(transcript_snapshot, template=template)

        def _on_success(summary):
            def _apply():
                session_ref.summary = summary
                self._update_transcript_display()
                self._set_status("Summary complete.")
                self._summarize_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                            fg=styles.TEXT_PRIMARY)
                self._email_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                        fg=styles.TEXT_PRIMARY)
                try:
                    self._export_svc.export_summary(session_ref)
                    self._session_svc.save(session_ref)
                except Exception as ex:
                    logger.error(f"Failed to auto-save summary: {ex}")
            self.after(0, _apply)

        def _on_error(e):
            self.after(0, lambda: messagebox.showerror("Summarization Error", str(e)))
            self.after(0, lambda: self._set_status("Summarization failed."))
            self.after(0, lambda: self._summarize_btn.config(
                state=tk.NORMAL, bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY))

        _run_async_in_thread(_coro, _on_success, _on_error)

    def _extract_action_items(self) -> None:
        if not self._summarizer:
            messagebox.showwarning("API Key Required",
                "Anthropic API key required. Add it in File > Settings.")
            return
        if not self._session or not self._session.segments:
            messagebox.showwarning("No Transcript", "Please process a recording first.")
            return
        transcript_snapshot = self._session.full_transcript()
        session_ref = self._session
        self._action_items_btn.config(state=tk.DISABLED, bg=styles.BG_INPUT,
                                       fg=styles.TEXT_MUTED)
        self._set_status("Extracting action items...")
        summarizer = self._summarizer

        def _coro():
            return summarizer.extract_action_items(transcript_snapshot)

        def _on_success(result):
            def _apply():
                session_ref.action_items = result
                self._update_transcript_display()
                self._set_status("Action items extracted.")
                self._action_items_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                               fg=styles.TEXT_PRIMARY)
                try:
                    self._export_svc.export_action_items(session_ref)
                    self._session_svc.save(session_ref)
                except Exception as ex:
                    logger.error(f"Failed to auto-save action items: {ex}")
            self.after(0, _apply)

        def _on_error(e):
            self.after(0, lambda: messagebox.showerror("Extraction Error", str(e)))
            self.after(0, lambda: self._set_status("Action items extraction failed."))
            self.after(0, lambda: self._action_items_btn.config(
                state=tk.NORMAL, bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY))

        _run_async_in_thread(_coro, _on_success, _on_error)

    def _extract_requirements(self) -> None:
        if not self._summarizer:
            messagebox.showwarning("API Key Required",
                "Anthropic API key required. Add it in File > Settings.")
            return
        if not self._session or not self._session.segments:
            messagebox.showwarning("No Transcript", "Please process a recording first.")
            return
        transcript_snapshot = self._session.full_transcript()
        session_ref = self._session
        self._requirements_btn.config(state=tk.DISABLED, bg=styles.BG_INPUT,
                                       fg=styles.TEXT_MUTED)
        self._set_status("Extracting requirements...")
        summarizer = self._summarizer

        def _coro():
            return summarizer.extract_requirements(transcript_snapshot)

        def _on_success(result):
            def _apply():
                session_ref.requirements = result
                self._update_transcript_display()
                self._set_status("Requirements extracted.")
                self._requirements_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                               fg=styles.TEXT_PRIMARY)
                try:
                    self._export_svc.export_requirements(session_ref)
                    self._session_svc.save(session_ref)
                except Exception as ex:
                    logger.error(f"Failed to auto-save requirements: {ex}")
            self.after(0, _apply)

        def _on_error(e):
            self.after(0, lambda: messagebox.showerror("Extraction Error", str(e)))
            self.after(0, lambda: self._set_status("Requirements extraction failed."))
            self.after(0, lambda: self._requirements_btn.config(
                state=tk.NORMAL, bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY))

        _run_async_in_thread(_coro, _on_success, _on_error)

    def _toggle_speakers(self) -> None:
        if self._speaker_expanded:
            self._speaker_body.pack_forget()
            self._speaker_toggle_var.set("Show")
            self._speaker_expanded = False
        else:
            self._speaker_body.pack(fill=tk.X)
            self._speaker_toggle_var.set("Hide")
            self._speaker_expanded = True

    def _update_transcript_display(self) -> None:
        """Rebuild the transcript panel with all available sections."""
        if not self._session:
            return
        parts = [self._session.full_transcript()]
        if self._session.summary:
            parts.append("── SUMMARY ──\n\n" + self._session.summary)
        if self._session.action_items:
            parts.append("── ACTION ITEMS ──\n\n" + self._session.action_items)
        if self._session.requirements:
            parts.append("── REQUIREMENTS ──\n\n" + self._session.requirements)
        self._transcript_panel.set_text("\n\n".join(parts))

    # ------------------------------------------------------------------ #
    # Email summary
    # ------------------------------------------------------------------ #

    def _email_summary(self) -> None:
        if not self._session or not self._session.summary:
            messagebox.showwarning("No Summary", "Please summarize the recording first.")
            return

        title    = self._get_meeting_name()
        date_str = datetime.datetime.now().strftime("%B %d, %Y").replace(" 0", " ")
        t_path   = None
        try:
            t_path = self._export_svc.export_transcript(self._session)
        except Exception:
            pass

        self._email_btn.config(state=tk.DISABLED, fg=styles.TEXT_MUTED)
        self._set_status("Sending email...")
        session_summary = self._session.summary
        session_ref = self._session
        email_to = self._settings.email_to

        def _send():
            try:
                import pythoncom
                import win32com.client
                import time

                pythoncom.CoInitialize()
                outlook = None
                for attempt in range(4):
                    try:
                        outlook = win32com.client.GetActiveObject("Outlook.Application")
                        break
                    except Exception:
                        try:
                            outlook = win32com.client.Dispatch("Outlook.Application")
                            break
                        except Exception:
                            time.sleep(2)

                if outlook is None:
                    raise RuntimeError(
                        "Could not connect to Outlook after 4 attempts.\n"
                        "Make sure Outlook is fully open (not just in the tray).")

                ns   = outlook.GetNamespace("MAPI")
                mail = outlook.CreateItem(0)

                transcript_line = ""
                if t_path:
                    transcript_line = (
                        f"<p style='font-size:12px;color:#888;margin-top:16px;'>"
                        f"Transcript saved to: {t_path}</p>"
                    )

                to_html = self._summarizer.summary_to_html

                sections_html = ""
                if session_summary:
                    sections_html += self._email_section("Summary", to_html(session_summary))
                if session_ref.action_items:
                    sections_html += self._email_section("Action Items", to_html(session_ref.action_items))
                if session_ref.requirements:
                    sections_html += self._email_section("Requirements", to_html(session_ref.requirements))

                html = f"""
<html><body style="font-family:Segoe UI,sans-serif;color:#1a1a1a;max-width:680px;">
<div style="background:#1565c0;padding:20px 24px;border-radius:8px;margin-bottom:20px;">
  <h2 style="color:#ffffff;margin:0;font-size:18px;">Meeting Notes</h2>
  <p style="color:#e3f2fd;margin:4px 0 0;font-size:13px;">{title} &mdash; {date_str}</p>
</div>
{sections_html}
{transcript_line}
<p style="font-size:11px;color:#aaa;margin-top:24px;
          border-top:1px solid #eee;padding-top:12px;">
  Sent automatically by Meeting Recorder
</p>
</body></html>
"""
                mail.Subject    = f"Meeting Notes: {title} ({date_str})"
                mail.BodyFormat = 2
                mail.HTMLBody   = html
                mail.To         = email_to if email_to else ns.CurrentUser.Address
                mail.Send()
                pythoncom.CoUninitialize()

                def _done():
                    self._set_status("Summary emailed to you.")
                    messagebox.showinfo(
                        "Email Sent",
                        "The meeting summary has been sent to your Outlook inbox.")
                    self._email_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                            fg=styles.TEXT_PRIMARY)
                self.after(0, _done)

            except Exception as e:
                def _fail(err=e):
                    self._set_status("Email failed.")
                    messagebox.showerror(
                        "Email Failed",
                        f"Could not send via Outlook:\n{err}\n\n"
                        "Make sure Outlook is fully open and try again.")
                    self._email_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                            fg=styles.TEXT_PRIMARY)
                self.after(0, _fail)

        threading.Thread(target=_send, daemon=True).start()

    # ------------------------------------------------------------------ #
    # Export & utilities
    # ------------------------------------------------------------------ #

    def _export(self) -> None:
        if not self._session:
            return
        try:
            self._session.display_name = self._get_meeting_name()
            paths = []
            paths.append(self._export_svc.export_transcript(self._session))
            if self._session.summary:
                paths.append(self._export_svc.export_summary(self._session))
            if self._session.action_items:
                paths.append(self._export_svc.export_action_items(self._session))
            if self._session.requirements:
                paths.append(self._export_svc.export_requirements(self._session))
            msg = "Exported files:\n" + "\n".join(paths)
            messagebox.showinfo("Export Complete", msg)
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

    def _open_recordings(self) -> None:
        folder = os.path.abspath(self._settings.recordings_dir)
        os.makedirs(folder, exist_ok=True)
        subprocess.Popen(f'explorer "{folder}"')

    def _on_rename_speaker(self, speaker_id: str, name: str) -> None:
        if self._session:
            self._session.rename_speaker(speaker_id, name)
            self._transcript_panel.set_text(self._session.full_transcript())

    def _set_status(self, message: str) -> None:
        self._status_var.set(message)

    def _thread_safe_status(self, message: str) -> None:
        if "__stage:" in message:
            parts = message.split("__stage:")
            for part in parts:
                if not part:
                    continue
                part = part.rstrip("_")
                bits = part.split(":")
                if len(bits) == 2:
                    stage, state = bits
                    self.after(0, lambda s=stage, st=state: self._set_stage(s, st))
        else:
            self.after(0, lambda: self._set_status(message))

    def _gather_existing(self, key: str) -> list:
        """Collect distinct values (client or project) from past sessions."""
        try:
            sessions = self._session_svc.list_sessions()
        except Exception:
            return []
        values = sorted({s.get(key, "") for s in sessions if s.get(key, "").strip()})
        return values

    def _pill_button(self, parent, text, color, command, outline=False) -> tk.Button:
        if outline:
            return tk.Button(
                parent, text=text, bg=styles.BG_PANEL, fg=styles.ACCENT,
                font=styles.FONT_BODY, relief=tk.FLAT, padx=16, pady=8,
                cursor="hand2", activebackground=styles.ACCENT_BG,
                activeforeground=styles.ACCENT, command=command,
                highlightbackground=styles.ACCENT_DIM, highlightthickness=1,
            )
        return tk.Button(
            parent, text=text, bg=color, fg="#ffffff",
            font=styles.FONT_BODY, relief=tk.FLAT, padx=16, pady=8,
            cursor="hand2", activebackground=color,
            activeforeground="#ffffff", command=command,
        )

    @staticmethod
    def _email_section(title: str, content_html: str) -> str:
        return (
            f'<div style="margin-bottom:16px;">'
            f'<h3 style="color:#1565c0;font-size:15px;margin:0 0 8px;'
            f'border-bottom:2px solid #bbdefb;padding-bottom:4px;">{title}</h3>'
            f'<div style="background:#f5f5f5;padding:16px 20px;border-radius:8px;'
            f'font-size:14px;line-height:1.7;color:#222;">'
            f'{content_html}</div></div>'
        )

    def _open_settings(self) -> None:
        SettingsDialog(self, self._settings, device_panel=self._device_panel)

    # ------------------------------------------------------------------ #
    # Calendar
    # ------------------------------------------------------------------ #

    def _load_calendar(self) -> None:
        """Background: fetch today's Outlook meetings."""
        try:
            meetings = get_todays_meetings()
        except Exception as e:
            logger.warning(f"Calendar load failed: {e}")
            meetings = []
        self.after(0, lambda: self._calendar_panel.load(meetings))

    def _refresh_calendar(self) -> None:
        self._calendar_panel.load([])  # show loading state
        tk.Label(self._calendar_panel, text="Refreshing...",
                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,
                 font=styles.FONT_SMALL).pack(pady=8)
        threading.Thread(target=self._load_calendar, daemon=True).start()

    def _toggle_calendar(self) -> None:
        if self._cal_expanded:
            self._cal_body.pack_forget()
            self._cal_toggle_var.set("Show")
            self._cal_expanded = False
        else:
            self._cal_body.pack(fill=tk.X, padx=8, pady=(0, 8))
            self._cal_toggle_var.set("Hide")
            self._cal_expanded = True

    def _start_from_meeting(self, meeting: dict) -> None:
        """Start recording with meeting name pre-filled from calendar."""
        if not self._recording_svc:
            return
        if self._recording_svc.is_recording:
            messagebox.showinfo("Already Recording",
                                "A recording is already in progress.")
            return
        # Format: "Meeting Subject - YYYY-MM-DD"
        date_str = meeting["start"].strftime("%Y-%m-%d")
        subject = meeting["subject"].strip() or "Meeting"
        self._meeting_name_var.set(f"{subject} - {date_str}")
        self._toggle_recording()

    def _on_meeting_upcoming(self, meeting: dict) -> None:
        """Popup notification shown when a meeting is about to start."""
        popup = tk.Toplevel(self)
        popup.title("Meeting Starting Soon")
        popup.configure(bg=styles.BG_PANEL)
        popup.transient(self)
        popup.resizable(False, False)
        # Position top-right of screen
        popup.update_idletasks()
        sw = self.winfo_screenwidth()
        popup.geometry(f"380x160+{sw - 400}+60")

        inner = tk.Frame(popup, bg=styles.BG_PANEL,
                         highlightbackground=styles.ACCENT,
                         highlightthickness=2)
        inner.pack(fill=tk.BOTH, expand=True)

        tk.Label(inner, text="📅  Meeting Starting Soon",
                 bg=styles.BG_PANEL, fg=styles.ACCENT,
                 font=styles.FONT_TITLE).pack(anchor="w", padx=14, pady=(10, 4))

        subject = meeting["subject"]
        if len(subject) > 42:
            subject = subject[:40] + "…"
        tk.Label(inner, text=subject, bg=styles.BG_PANEL,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_BODY,
                 anchor="w", justify="left", wraplength=340).pack(
                     anchor="w", padx=14)

        start_str = meeting["start"].strftime("%H:%M")
        tk.Label(inner, text=f"Starts at {start_str}",
                 bg=styles.BG_PANEL, fg=styles.TEXT_MUTED,
                 font=styles.FONT_SMALL).pack(anchor="w", padx=14, pady=(2, 10))

        btn_row = tk.Frame(inner, bg=styles.BG_PANEL)
        btn_row.pack(fill=tk.X, padx=14, pady=(0, 10))

        def on_start():
            popup.destroy()
            self._start_from_meeting(meeting)

        def on_dismiss():
            if self._calendar_monitor:
                self._calendar_monitor.dismiss(meeting)
            popup.destroy()

        tk.Button(btn_row, text="Start Recording", bg=styles.ACCENT,
                  fg="#ffffff", font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=16, pady=6, cursor="hand2",
                  command=on_start).pack(side=tk.RIGHT)
        tk.Button(btn_row, text="Dismiss", bg=styles.BG_INPUT,
                  fg=styles.TEXT_MUTED, font=styles.FONT_BODY,
                  relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
                  command=on_dismiss).pack(side=tk.RIGHT, padx=(0, 6))

        # Auto-dismiss after 60 seconds
        popup.after(60000, lambda: popup.winfo_exists() and popup.destroy())

    # ------------------------------------------------------------------ #
    # Session history
    # ------------------------------------------------------------------ #

    def _open_session_history(self) -> None:
        SessionBrowser(
            self, self._session_svc,
            on_open=self._load_session_by_id,
            on_bulk_process=self._bulk_process,
            recordings_dir=self._settings.recordings_dir,
        )

    def _open_follow_up_tracker(self) -> None:
        FollowUpTracker(
            self, self._session_svc,
            on_open_session=self._load_session_by_id,
        )

    def _open_prep_brief(self) -> None:
        """Generate a prep brief based on current meeting name + client/project."""
        if not self._summarizer:
            messagebox.showwarning("API Key Required",
                "Anthropic API key required for prep briefs. "
                "Add it in File > Settings.")
            return

        subject = self._meeting_name_var.get().strip() or "Upcoming Meeting"
        client = self._client_var.get().strip()
        project = self._project_var.get().strip()

        # Find related sessions (match by client or project, excluding current session)
        all_sessions = self._session_svc.list_sessions()
        related = []
        for s in all_sessions:
            if self._session and s.get("session_id") == self._session.session_id:
                continue
            match_client = bool(client and s.get("client") == client)
            match_project = bool(project and s.get("project") == project)
            if match_client or match_project:
                related.append(s)

        if not related and not (client or project):
            if not messagebox.askyesno(
                    "No Client/Project Set",
                    "No client or project tag set — brief will pull from ALL "
                    "past meetings, which may be noisy.\n\n"
                    "Continue anyway?"):
                return
            related = all_sessions[:10]  # newest 10 as fallback

        def generator(prior_notes, upcoming_subject, on_result, on_error):
            def _coro():
                return self._summarizer.meeting_prep_brief(
                    prior_notes, upcoming_subject)
            _run_async_in_thread(_coro, on_result, on_error)

        PrepBriefDialog(self, subject, related, generator)

    def _load_session_by_id(self, session_id: str) -> None:
        try:
            session = self._session_svc.load_full(session_id)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return
        if not session:
            messagebox.showwarning("Not Found",
                                   f"Session {session_id} not found.")
            return
        self._session = session
        if self._recording_svc:
            self._recording_svc.set_session(session)
        self._meeting_name_var.set(session.display_name or "")
        self._template_var.set(session.template or "General")
        self._client_var.set(session.client or "")
        self._project_var.set(session.project or "")
        self._reset_stages()
        self._update_transcript_display()
        if session.speakers:
            self._speaker_panel.populate(session)
        # Enable relevant buttons based on what the session has
        if session.audio_path:
            self._process_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                      fg=styles.TEXT_PRIMARY)
        if session.segments:
            self._summarize_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                        fg=styles.TEXT_PRIMARY)
            self._action_items_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                           fg=styles.TEXT_PRIMARY)
            self._requirements_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                           fg=styles.TEXT_PRIMARY)
            self._export_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                     fg=styles.TEXT_PRIMARY)
            # Mark stages done
            self._set_stage("transcribe", "done")
            self._set_stage("diarize", "done")
            self._set_stage("speakers", "done")
            self._set_stage("complete", "done")
        if session.summary:
            self._email_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                    fg=styles.TEXT_PRIMARY)
        self._set_status(f"Loaded: {session.display_name or session.session_id}")

    def _bulk_process(self, session_ids: list) -> None:
        """Sequentially process a list of unprocessed sessions."""
        if not self._models_ready or not self._recording_svc:
            messagebox.showwarning(
                "Models Not Ready",
                "Transcription/diarization models aren't loaded. "
                "Check File > Settings and restart.")
            return
        logger.info(f"Bulk processing {len(session_ids)} sessions...")

        def _worker():
            for i, sid in enumerate(session_ids, 1):
                try:
                    session = self._session_svc.load_full(sid)
                    if not session or not session.audio_path:
                        continue
                    self.after(0, lambda s=session, i=i, n=len(session_ids):
                               self._set_status(
                                   f"Processing {i}/{n}: {s.display_name}"))
                    self._recording_svc.set_session(session)

                    # Run the async process_session synchronously
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    try:
                        loop.run_until_complete(self._recording_svc.process_session())
                    finally:
                        loop.close()

                    self._session_svc.save(session)
                    logger.info(f"Bulk-processed: {session.display_name}")
                except Exception as e:
                    logger.error(f"Bulk process failed for {sid}: {e}")
            self.after(0, lambda: self._set_status(
                f"Bulk process complete ({len(session_ids)} sessions)"))
            self.after(0, lambda: messagebox.showinfo(
                "Bulk Process Complete",
                f"Processed {len(session_ids)} sessions."))

        threading.Thread(target=_worker, daemon=True).start()

    def _open_logs_folder(self) -> None:
        logs_dir = os.path.abspath(self._settings.recordings_dir)
        try:
            os.startfile(logs_dir)
        except Exception:
            messagebox.showinfo("Logs", f"Session logs are in:\n{logs_dir}")

    def _show_about(self) -> None:
        messagebox.showinfo(
            "About Meeting Recorder",
            "Meeting Recorder v1.0\n\n"
            "AI-powered meeting transcription,\n"
            "speaker diarization, and summarization.\n\n"
            "Powered by Whisper, Pyannote, and Claude.")

    def _on_close(self) -> None:
        if self._recording_svc and self._recording_svc.is_recording:
            self._recording_svc.stop_recording()
        if self._calendar_monitor:
            self._calendar_monitor.stop()
        self.destroy()
