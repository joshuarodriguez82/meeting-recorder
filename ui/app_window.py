"""
Main application window — Material You dark theme.
"""

import asyncio
import datetime
import os
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
from typing import Optional
import uuid

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
        self._active_meeting: Optional[dict] = None
        self._progress_after = None

        self._transcription = TranscriptionEngine(settings.whisper_model)
        self._diarization   = DiarizationEngine(settings.hf_token, settings.max_speakers)
        self._summarizer    = Summarizer(settings.anthropic_api_key)
        self._session_svc   = SessionService(settings.recordings_dir)
        self._export_svc    = ExportService(settings.recordings_dir)
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
        self.geometry("920x820")
        self.minsize(800, 680)
        self.configure(bg=styles.BG_DARK)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        try:
            self.iconbitmap("meeting_recorder.ico")
        except Exception:
            pass

    def _build_layout(self) -> None:
        outer = tk.Frame(self, bg=styles.BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=styles.PAD_LG, pady=styles.PAD_LG)

        # ── Top bar ──────────────────────────────────────────────────
        topbar = tk.Frame(outer, bg=styles.BG_DARK)
        topbar.pack(fill=tk.X, pady=(0, styles.PAD))

        icon_frame = tk.Frame(topbar, bg=styles.ACCENT_BG, width=42, height=42)
        icon_frame.pack(side=tk.LEFT)
        icon_frame.pack_propagate(False)
        tk.Label(icon_frame, text="🎙", bg=styles.ACCENT_BG,
                 font=("Segoe UI", 16)).place(relx=0.5, rely=0.5, anchor="center")

        title_block = tk.Frame(topbar, bg=styles.BG_DARK)
        title_block.pack(side=tk.LEFT, padx=(10, 0))
        tk.Label(title_block, text="Meeting Recorder", bg=styles.BG_DARK,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_HEADER).pack(anchor="w")

        self._status_var = tk.StringVar(value="Ready")
        tk.Label(topbar, textvariable=self._status_var,
                 bg=styles.BG_DARK, fg=styles.ACCENT,
                 font=styles.FONT_SMALL).pack(side=tk.RIGHT, anchor="s")

        # ── Meeting name card ─────────────────────────────────────────
        name_card = tk.Frame(outer, bg=styles.BG_PANEL,
                              highlightbackground=styles.BORDER,
                              highlightthickness=1)
        name_card.pack(fill=tk.X, pady=(0, styles.PAD))
        tk.Label(name_card, text="MEETING NAME",
                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=14, pady=(10, 4))

        name_row = tk.Frame(name_card, bg=styles.BG_PANEL)
        name_row.pack(fill=tk.X, padx=8, pady=(0, 10))

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
        self._name_entry.pack(fill=tk.X, ipady=8)
        self._meeting_name_var.trace_add("write", self._on_name_change)

        # ── Device card ───────────────────────────────────────────────
        device_card = tk.Frame(outer, bg=styles.BG_PANEL,
                                highlightbackground=styles.BORDER,
                                highlightthickness=1)
        device_card.pack(fill=tk.X, pady=(0, styles.PAD))
        tk.Label(device_card, text="AUDIO DEVICES",
                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=14, pady=(10, 4))
        self._device_panel = DevicePanel(device_card)
        self._device_panel.pack(fill=tk.X, padx=8, pady=(0, 8))

        # ── Progress stages ───────────────────────────────────────────
        prog_card = tk.Frame(outer, bg=styles.BG_PANEL,
                              highlightbackground=styles.BORDER,
                              highlightthickness=1)
        prog_card.pack(fill=tk.X, pady=(0, styles.PAD))

        prog_inner = tk.Frame(prog_card, bg=styles.BG_PANEL)
        prog_inner.pack(fill=tk.X, padx=14, pady=10)

        self._stage_labels = {}
        stages = [
            ("transcribe", "Transcribe"),
            ("diarize",    "Diarize"),
            ("speakers",   "ID Speakers"),
            ("complete",   "Complete"),
        ]
        for key, label in stages:
            col = tk.Frame(prog_inner, bg=styles.BG_PANEL)
            col.pack(side=tk.LEFT, expand=True, fill=tk.X)
            dot = tk.Label(col, text="○", bg=styles.BG_PANEL,
                           fg=styles.TEXT_HINT, font=("Segoe UI", 14))
            dot.pack()
            lbl = tk.Label(col, text=label, bg=styles.BG_PANEL,
                           fg=styles.TEXT_HINT, font=styles.FONT_SMALL)
            lbl.pack()
            self._stage_labels[key] = (dot, lbl)

        # ── Action row 1 ──────────────────────────────────────────────
        row1 = tk.Frame(outer, bg=styles.BG_DARK)
        row1.pack(fill=tk.X, pady=(0, 8))

        self._rec_btn = self._pill_button(
            row1, "⏺  Start Recording", styles.DANGER, self._toggle_recording)
        self._rec_btn.pack(side=tk.LEFT, padx=(0, 8))

        self._load_btn = self._pill_button(
            row1, "📂  Load File", styles.ACCENT_DIM, self._load_audio_file)
        self._load_btn.pack(side=tk.LEFT, padx=(0, 8))

        self._process_btn = self._pill_button(
            row1, "⚙  Process", styles.BG_INPUT, self._process, outline=True)
        self._process_btn.pack(side=tk.LEFT)
        self._process_btn.config(state=tk.DISABLED)

        # ── Action row 2 ──────────────────────────────────────────────
        row2 = tk.Frame(outer, bg=styles.BG_DARK)
        row2.pack(fill=tk.X, pady=(0, styles.PAD))

        self._summarize_btn = self._pill_button(
            row2, "✨  Summarize", styles.BG_INPUT, self._summarize, outline=True)
        self._summarize_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._summarize_btn.config(state=tk.DISABLED)

        self._export_btn = self._pill_button(
            row2, "💾  Export", styles.BG_INPUT, self._export, outline=True)
        self._export_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._export_btn.config(state=tk.DISABLED)

        self._email_btn = self._pill_button(
            row2, "✉  Email Summary", styles.BG_INPUT, self._email_summary, outline=True)
        self._email_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._email_btn.config(state=tk.DISABLED)

        self._folder_btn = self._pill_button(
            row2, "📁  Recordings", styles.BG_INPUT, self._open_recordings, outline=True)
        self._folder_btn.pack(side=tk.LEFT)

        # ── Speaker card ──────────────────────────────────────────────
        speaker_card = tk.Frame(outer, bg=styles.BG_PANEL,
                                 highlightbackground=styles.BORDER,
                                 highlightthickness=1)
        speaker_card.pack(fill=tk.X, pady=(0, styles.PAD))
        tk.Label(speaker_card, text="SPEAKERS",
                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=14, pady=(10, 4))
        self._speaker_panel = SpeakerPanel(speaker_card, on_rename=self._on_rename_speaker)
        self._speaker_panel.pack(fill=tk.X, padx=8, pady=(0, 8))

        # ── Transcript card ───────────────────────────────────────────
        transcript_card = tk.Frame(outer, bg=styles.BG_PANEL,
                                    highlightbackground=styles.BORDER,
                                    highlightthickness=1)
        transcript_card.pack(fill=tk.BOTH, expand=True)
        tk.Label(transcript_card, text="TRANSCRIPT",
                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=14, pady=(10, 4))
        self._transcript_panel = TranscriptPanel(transcript_card)
        self._transcript_panel.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

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
                                  bg=styles.BG_INPUT, fg=styles.DANGER_DIM)
            self._load_btn.config(state=tk.DISABLED)
            self._process_btn.config(state=tk.DISABLED)
            self._summarize_btn.config(state=tk.DISABLED)
            self._export_btn.config(state=tk.DISABLED)
            self._email_btn.config(state=tk.DISABLED)
            self._set_status("Recording...")
        else:
            self._session = self._recording_svc.stop_recording()
            self._rec_btn.config(text="⏺  Start Recording",
                                  bg=styles.DANGER, fg=styles.TEXT_PRIMARY)
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
        self._set_stage("transcribe", "done")
        self._set_stage("diarize", "done")
        self._set_stage("speakers", "active")
        self._set_status("Identifying speakers...")
        self._speaker_panel.populate(session)
        self._transcript_panel.set_text(session.full_transcript())
        self._summarize_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                    fg=styles.TEXT_PRIMARY)
        self._export_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,
                                 fg=styles.TEXT_PRIMARY)
        try:
            self._session_svc.save(session)
            self._export_svc.export_transcript(session)
        except Exception as e:
            logger.error(f"Failed to save session: {e}")
        self._auto_identify_speakers(session)

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
                    session_ref.full_transcript() + "\n\n── SUMMARY ──\n\n" + summary)
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

                formatted_summary = self._summarizer.summary_to_html(session_summary)
                html = f"""
<html><body style="font-family:Segoe UI,sans-serif;color:#1a1a1a;max-width:680px;">
<div style="background:#003a57;padding:20px 24px;border-radius:8px;margin-bottom:20px;">
  <h2 style="color:#4fc3f7;margin:0;font-size:18px;">Meeting Recorder Summary</h2>
  <p style="color:#90caf9;margin:4px 0 0;font-size:13px;">{title} &mdash; {date_str}</p>
</div>
<div style="background:#f5f5f5;padding:20px 24px;border-radius:8px;
            font-size:14px;line-height:1.7;color:#222;">
{formatted_summary}
</div>
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
                mail.To         = ns.CurrentUser.Address
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
            t_path = self._export_svc.export_transcript(self._session)
            s_path = (self._export_svc.export_summary(self._session)
                      if self._session.summary else None)
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

    def _pill_button(self, parent, text, color, command, outline=False) -> tk.Button:
        if outline:
            return tk.Button(
                parent, text=text, bg=styles.BG_DARK, fg=styles.TEXT_MUTED,
                font=styles.FONT_BODY, relief=tk.FLAT, padx=16, pady=8,
                cursor="hand2", activebackground=styles.BG_INPUT,
                activeforeground=styles.TEXT_PRIMARY, command=command,
                highlightbackground=styles.BORDER, highlightthickness=1,
            )
        return tk.Button(
            parent, text=text, bg=color, fg=styles.TEXT_PRIMARY,
            font=styles.FONT_BODY, relief=tk.FLAT, padx=16, pady=8,
            cursor="hand2", activebackground=color,
            activeforeground=styles.TEXT_PRIMARY, command=command,
        )

    def _on_close(self) -> None:
        if self._recording_svc.is_recording:
            self._recording_svc.stop_recording()
        self.destroy()
