"""
Main application window — Material You dark theme.
"""

import asyncio
import os
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from typing import Optional
import uuid
from datetime import datetime

from config.settings import Settings
from core.diarization import DiarizationEngine
from core.summarizer import Summarizer
from core.transcription import TranscriptionEngine
from models.session import Session
from services.export_service import ExportService
from services.recording_service import RecordingService
from services.session_service import SessionService
from ui import styles
from ui.device_panel import DevicePanel
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

        self._transcription = TranscriptionEngine(settings.whisper_model)
        self._diarization = DiarizationEngine(settings.hf_token, settings.max_speakers)
        self._summarizer = Summarizer(settings.anthropic_api_key)
        self._session_svc = SessionService(settings.recordings_dir)
        self._export_svc = ExportService(settings.recordings_dir)
        self._recording_svc = RecordingService(
            settings=settings,
            transcription_engine=self._transcription,
            diarization_engine=self._diarization,
            on_status=self._thread_safe_status,
        )

        self._build_window()
        self._build_layout()

    def _build_window(self) -> None:
        self.title("Meeting Recorder")
        self.geometry("920x780")
        self.minsize(800, 640)
        self.configure(bg=styles.BG_DARK)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_layout(self) -> None:
        outer = tk.Frame(self, bg=styles.BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=styles.PAD_LG, pady=styles.PAD_LG)

        # ── Top bar ──────────────────────────────────────────────────
        topbar = tk.Frame(outer, bg=styles.BG_DARK)
        topbar.pack(fill=tk.X, pady=(0, styles.PAD))

        icon_frame = tk.Frame(topbar, bg=styles.ACCENT_BG, width=42, height=42)
        icon_frame.pack(side=tk.LEFT)
        icon_frame.pack_propagate(False)
        tk.Label(icon_frame, text="🎙", bg=styles.ACCENT_BG, font=("Segoe UI", 16)).place(relx=0.5, rely=0.5, anchor="center")

        title_block = tk.Frame(topbar, bg=styles.BG_DARK)
        title_block.pack(side=tk.LEFT, padx=(10, 0))
        tk.Label(title_block, text="Meeting Recorder", bg=styles.BG_DARK,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_HEADER).pack(anchor="w")

        self._status_var = tk.StringVar(value="Ready")
        tk.Label(topbar, textvariable=self._status_var, bg=styles.BG_DARK,
                 fg=styles.ACCENT, font=styles.FONT_SMALL).pack(side=tk.RIGHT, anchor="s")

        # ── Device card ───────────────────────────────────────────────
        device_card = tk.Frame(outer, bg=styles.BG_PANEL,
                                highlightbackground=styles.BORDER,
                                highlightthickness=1)
        device_card.pack(fill=tk.X, pady=(0, styles.PAD))
        self._apply_radius(device_card)

        tk.Label(device_card, text="AUDIO DEVICES", bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=("Segoe UI", 9)).pack(anchor="w", padx=14, pady=(10, 4))

        self._device_panel = DevicePanel(device_card)
        self._device_panel.pack(fill=tk.X, padx=8, pady=(0, 8))

        # ── Action row 1 — recording ──────────────────────────────────
        row1 = tk.Frame(outer, bg=styles.BG_DARK)
        row1.pack(fill=tk.X, pady=(0, 8))

        self._rec_btn = self._pill_button(row1, "⏺  Start Recording", styles.DANGER, self._toggle_recording)
        self._rec_btn.pack(side=tk.LEFT, padx=(0, 8))

        self._load_btn = self._pill_button(row1, "📂  Load File", styles.ACCENT_DIM, self._load_audio_file)
        self._load_btn.pack(side=tk.LEFT, padx=(0, 8))

        self._process_btn = self._pill_button(row1, "⚙  Process", styles.BG_INPUT, self._process, outline=True)
        self._process_btn.pack(side=tk.LEFT)
        self._process_btn.config(state=tk.DISABLED)

        # ── Action row 2 — post-processing ────────────────────────────
        row2 = tk.Frame(outer, bg=styles.BG_DARK)
        row2.pack(fill=tk.X, pady=(0, styles.PAD))

        self._summarize_btn = self._pill_button(row2, "✨  Summarize", styles.BG_INPUT, self._summarize, outline=True)
        self._summarize_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._summarize_btn.config(state=tk.DISABLED)

        self._export_btn = self._pill_button(row2, "💾  Export", styles.BG_INPUT, self._export, outline=True)
        self._export_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._export_btn.config(state=tk.DISABLED)

        self._folder_btn = self._pill_button(row2, "📁  Recordings", styles.BG_INPUT, self._open_recordings, outline=True)
        self._folder_btn.pack(side=tk.LEFT)

        # ── Speaker card ──────────────────────────────────────────────
        speaker_card = tk.Frame(outer, bg=styles.BG_PANEL,
                                 highlightbackground=styles.BORDER,
                                 highlightthickness=1)
        speaker_card.pack(fill=tk.X, pady=(0, styles.PAD))
        self._apply_radius(speaker_card)

        tk.Label(speaker_card, text="SPEAKERS", bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=("Segoe UI", 9)).pack(anchor="w", padx=14, pady=(10, 4))

        self._speaker_panel = SpeakerPanel(speaker_card, on_rename=self._on_rename_speaker)
        self._speaker_panel.pack(fill=tk.X, padx=8, pady=(0, 8))

        # ── Transcript card ───────────────────────────────────────────
        transcript_card = tk.Frame(outer, bg=styles.BG_PANEL,
                                    highlightbackground=styles.BORDER,
                                    highlightthickness=1)
        transcript_card.pack(fill=tk.BOTH, expand=True)
        self._apply_radius(transcript_card)

        tk.Label(transcript_card, text="TRANSCRIPT", bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=("Segoe UI", 9)).pack(anchor="w", padx=14, pady=(10, 4))

        self._transcript_panel = TranscriptPanel(transcript_card)
        self._transcript_panel.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

    # ------------------------------------------------------------------ #
    # Button factory
    # ------------------------------------------------------------------ #

    def _pill_button(self, parent, text, color, command, outline=False) -> tk.Button:
        if outline:
            btn = tk.Button(
                parent, text=text, bg=styles.BG_DARK, fg=styles.TEXT_MUTED,
                font=styles.FONT_BODY, relief=tk.FLAT, padx=16, pady=8,
                cursor="hand2", activebackground=styles.BG_INPUT,
                activeforeground=styles.TEXT_PRIMARY, command=command,
                highlightbackground=styles.BORDER, highlightthickness=1,
            )
        else:
            btn = tk.Button(
                parent, text=text, bg=color, fg=styles.TEXT_PRIMARY,
                font=styles.FONT_BODY, relief=tk.FLAT, padx=16, pady=8,
                cursor="hand2", activebackground=color,
                activeforeground=styles.TEXT_PRIMARY, command=command,
            )
        return btn

    def _apply_radius(self, frame) -> None:
        pass

    # ------------------------------------------------------------------ #
    # Recording
    # ------------------------------------------------------------------ #

    def _toggle_recording(self) -> None:
        if not self._recording_svc.is_recording:
            mic_idx = self._device_panel.get_mic_index()
            out_idx = self._device_panel.get_output_index()
            try:
                self._session = self._recording_svc.start_recording(mic_idx, out_idx)
            except Exception as e:
                messagebox.showerror("Recording Error", str(e))
                return
            self._transcript_panel.clear()
            self._rec_btn.config(text="⏹  Stop Recording", bg=styles.BG_INPUT,
                                  fg=styles.DANGER_DIM)
            self._load_btn.config(state=tk.DISABLED)
            self._process_btn.config(state=tk.DISABLED)
            self._summarize_btn.config(state=tk.DISABLED)
            self._export_btn.config(state=tk.DISABLED)
            self._set_status("Recording...")
        else:
            self._session = self._recording_svc.stop_recording()
            self._rec_btn.config(text="⏺  Start Recording", bg=styles.DANGER,
                                  fg=styles.TEXT_PRIMARY)
            self._load_btn.config(state=tk.NORMAL)
            if self._session and self._session.audio_path:
                self._process_btn.config(state=tk.NORMAL,
                                          bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY)
            self._set_status("Recording saved. Ready to process.")

    # ------------------------------------------------------------------ #
    # Load file
    # ------------------------------------------------------------------ #

    def _load_audio_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Audio File",
            filetypes=[
                ("Audio Files", "*.wav *.mp3 *.m4a *.flac *.ogg *.aac"),
                ("WAV files", "*.wav"),
                ("MP3 files", "*.mp3"),
                ("All files", "*.*"),
            ]
        )
        if not file_path:
            return

        session_id = uuid.uuid4().hex[:8].upper()
        self._session = Session(session_id=session_id)
        self._session.started_at = datetime.now()
        self._session.ended_at = datetime.now()
        self._session.audio_path = file_path

        self._transcript_panel.clear()
        self._summarize_btn.config(state=tk.DISABLED)
        self._export_btn.config(state=tk.DISABLED)
        self._process_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                  fg=styles.TEXT_PRIMARY)

        filename = os.path.basename(file_path)
        self._set_status(f"Loaded: {filename}")
        logger.info(f"Loaded external audio file: {file_path}")

    # ------------------------------------------------------------------ #
    # Processing
    # ------------------------------------------------------------------ #

    def _process(self) -> None:
        if not self._session or not self._session.audio_path:
            messagebox.showwarning("No Audio", "Please record or load an audio file first.")
            return

        self._process_btn.config(state=tk.DISABLED, bg=styles.BG_INPUT,
                                  fg=styles.TEXT_MUTED)
        self._set_status("Processing...")

        audio_path = self._session.audio_path
        session = self._session
        transcription = self._transcription
        diarization = self._diarization

        def _coro():
            async def _run():
                from core.diarization import DiarizationEngine
                from models.segment import Segment

                transcription_result = await transcription.transcribe(audio_path)
                if not transcription_result:
                    return session
                diarization_turns = await diarization.diarize(audio_path)
                attributed = DiarizationEngine.assign_speakers(
                    transcription_result, diarization_turns)
                for raw in attributed:
                    speaker = session.get_or_create_speaker(raw["speaker_id"])
                    segment = Segment(
                        speaker_id=speaker.speaker_id,
                        start=raw["start"],
                        end=raw["end"],
                        text=raw["text"],
                    )
                    session.segments.append(segment)
                return session
            return _run()

        def _on_success(result):
            self.after(0, lambda: self._on_process_complete(result))

        def _on_error(e):
            self.after(0, lambda: messagebox.showerror("Processing Error", str(e)))
            self.after(0, lambda: self._set_status("Processing failed."))
            self.after(0, lambda: self._process_btn.config(
                state=tk.NORMAL, bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY))

        _run_async_in_thread(_coro, _on_success, _on_error)

    def _on_process_complete(self, session: Session) -> None:
        self._session = session
        self._speaker_panel.populate(session)
        self._transcript_panel.set_text(session.full_transcript())
        self._summarize_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                    fg=styles.TEXT_PRIMARY)
        self._export_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                 fg=styles.TEXT_PRIMARY)
        self._set_status("Processing complete.")
        try:
            self._session_svc.save(session)
        except Exception as e:
            logger.error(f"Failed to save session JSON: {e}")

    # ------------------------------------------------------------------ #
    # Summarize
    # ------------------------------------------------------------------ #

    def _summarize(self) -> None:
        if not self._session or not self._session.segments:
            messagebox.showwarning("No Transcript", "Please process a recording first.")
            return

        transcript_snapshot = self._session.full_transcript()
        session_ref = self._session
        self._summarize_btn.config(state=tk.DISABLED, bg=styles.BG_INPUT,
                                    fg=styles.TEXT_MUTED)
        self._set_status("Generating AI summary...")
        summarizer = self._summarizer

        def _coro():
            return summarizer.summarize(transcript_snapshot)

        def _on_success(summary):
            def _apply():
                session_ref.summary = summary
                self._transcript_panel.set_text(
                    session_ref.full_transcript() + "\n\n── SUMMARY ──\n\n" + summary
                )
                self._set_status("Summary complete.")
                self._summarize_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                            fg=styles.TEXT_PRIMARY)
            self.after(0, _apply)

        def _on_error(e):
            self.after(0, lambda: messagebox.showerror("Summarization Error", str(e)))
            self.after(0, lambda: self._set_status("Summarization failed."))
            self.after(0, lambda: self._summarize_btn.config(
                state=tk.NORMAL, bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY))

        _run_async_in_thread(_coro, _on_success, _on_error)

    # ------------------------------------------------------------------ #
    # Export & utilities
    # ------------------------------------------------------------------ #

    def _export(self) -> None:
        if not self._session:
            return
        try:
            t_path = self._export_svc.export_transcript(self._session)
            s_path = (
                self._export_svc.export_summary(self._session)
                if self._session.summary else None
            )
            msg = f"Transcript saved:\n{t_path}"
            if s_path:
                msg += f"\n\nSummary saved:\n{s_path}"
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
        self.after(0, lambda: self._set_status(message))

    def _on_close(self) -> None:
        if self._recording_svc.is_recording:
            self._recording_svc.stop_recording()
        self.destroy()