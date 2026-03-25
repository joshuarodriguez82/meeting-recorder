"""
Today's meetings panel — shows Outlook appointments with Record buttons.
"""

import datetime
import tkinter as tk
from typing import Callable, List
from ui import styles
from utils.logger import get_logger

logger = get_logger(__name__)


class CalendarPanel(tk.Frame):

    def __init__(self, parent, on_record: Callable[[dict], None], **kwargs):
        super().__init__(parent, bg=styles.BG_PANEL, **kwargs)
        self._on_record = on_record
        self._meetings: List[dict] = []

        self._placeholder = tk.Label(
            self,
            text="Loading calendar...",
            bg=styles.BG_PANEL, fg=styles.TEXT_HINT,
            font=styles.FONT_SMALL,
        )
        self._placeholder.pack(pady=8)

    def load(self, meetings: List[dict]) -> None:
        """Populate the panel with today's meetings."""
        for w in self.winfo_children():
            w.destroy()
        self._meetings = meetings

        if not meetings:
            tk.Label(
                self,
                text="No meetings scheduled for today.",
                bg=styles.BG_PANEL, fg=styles.TEXT_HINT,
                font=styles.FONT_SMALL,
            ).pack(pady=8, anchor="w")
            return

        now = datetime.datetime.now()

        for meeting in meetings:
            self._build_meeting_row(meeting, now)

    def _build_meeting_row(self, meeting: dict, now: datetime.datetime) -> None:
        start    = meeting["start"]
        end      = meeting["end"]
        is_now   = start <= now <= end
        is_past  = end < now

        row = tk.Frame(self, bg=styles.BG_INPUT if is_now else styles.BG_PANEL,
                       pady=6, padx=10)
        row.pack(fill=tk.X, pady=3)

        # Time badge
        time_str = f"{start.strftime('%H:%M')} – {end.strftime('%H:%M')}"
        time_color = styles.ACCENT if is_now else (styles.TEXT_HINT if is_past else styles.TEXT_MUTED)
        tk.Label(row, text=time_str, bg=row.cget("bg"),
                 fg=time_color, font=styles.FONT_SMALL, width=13,
                 anchor="w").pack(side=tk.LEFT)

        # Live indicator dot
        if is_now:
            dot = tk.Label(row, text="●", bg=row.cget("bg"),
                           fg=styles.DANGER, font=("Segoe UI", 8))
            dot.pack(side=tk.LEFT, padx=(0, 6))

        # Meeting title
        title_color = styles.TEXT_PRIMARY if not is_past else styles.TEXT_HINT
        title = meeting["subject"]
        if len(title) > 38:
            title = title[:36] + "…"
        tk.Label(row, text=title, bg=row.cget("bg"),
                 fg=title_color, font=styles.FONT_BODY,
                 anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Duration
        tk.Label(row, text=f"{meeting['duration']}m", bg=row.cget("bg"),
                 fg=styles.TEXT_HINT, font=styles.FONT_SMALL).pack(side=tk.LEFT, padx=(4, 8))

        # Record button
        if not is_past:
            btn_text = "⏺ Record" if not is_now else "⏺ Recording now"
            btn = tk.Button(
                row,
                text=btn_text,
                bg=styles.DANGER if is_now else styles.ACCENT_DIM,
                fg=styles.TEXT_PRIMARY,
                font=styles.FONT_SMALL,
                relief=tk.FLAT,
                padx=10, pady=3,
                cursor="hand2",
                command=lambda m=meeting: self._on_record(m),
            )
            btn.pack(side=tk.RIGHT)
        else:
            tk.Label(row, text="Done", bg=row.cget("bg"),
                     fg=styles.TEXT_HINT, font=styles.FONT_SMALL).pack(side=tk.RIGHT)

    def mark_recording(self, meeting: dict) -> None:
        """Refresh to show recording state."""
        pass
