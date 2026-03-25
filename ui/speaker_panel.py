"""
Speaker naming panel — appears after diarization.
Displays detected speakers and allows renaming them.
"""

import tkinter as tk
from typing import Callable, Dict

from models.session import Session
from ui import styles


class SpeakerPanel(tk.LabelFrame):
    """Widget for assigning human names to detected speaker IDs."""

    def __init__(self, parent, on_rename: Callable[[str, str], None], **kwargs):
        super().__init__(
            parent,
            text=" Speaker Names ",
            bg=styles.BG_PANEL,
            fg=styles.TEXT_PRIMARY,
            font=styles.FONT_BODY,
            bd=1,
            relief=tk.SOLID,
            **kwargs,
        )
        self._on_rename = on_rename
        self._entries: Dict[str, tk.Entry] = {}
        self._placeholder = tk.Label(
            self,
            text="Speakers will appear here after processing.",
            bg=styles.BG_PANEL,
            fg=styles.TEXT_MUTED,
            font=styles.FONT_SMALL,
        )
        self._placeholder.pack(pady=styles.PAD)

    def populate(self, session: Session) -> None:
        """Render one row per detected speaker with a rename entry."""
        for widget in self.winfo_children():
            widget.destroy()
        self._entries.clear()

        if not session.speakers:
            tk.Label(self, text="No speakers detected.", bg=styles.BG_PANEL,
                     fg=styles.TEXT_MUTED, font=styles.FONT_SMALL).pack(pady=styles.PAD)
            return

        for row, (speaker_id, speaker) in enumerate(session.speakers.items()):
            frame = tk.Frame(self, bg=styles.BG_PANEL)
            frame.pack(fill=tk.X, padx=styles.PAD, pady=3)

            tk.Label(frame, text=f"{speaker_id}:", bg=styles.BG_PANEL,
                     fg=styles.TEXT_MUTED, font=styles.FONT_SMALL, width=18, anchor="w").pack(
                side=tk.LEFT
            )

            entry = tk.Entry(frame, bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,
                             insertbackground=styles.TEXT_PRIMARY,
                             font=styles.FONT_BODY, relief=tk.FLAT)
            entry.insert(0, speaker.display_name)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
            self._entries[speaker_id] = entry

            entry.bind("<FocusOut>", lambda e, sid=speaker_id: self._commit(sid))
            entry.bind("<Return>", lambda e, sid=speaker_id: self._commit(sid))

    def _commit(self, speaker_id: str) -> None:
        entry = self._entries.get(speaker_id)
        if entry:
            name = entry.get().strip()
            if name:
                self._on_rename(speaker_id, name)