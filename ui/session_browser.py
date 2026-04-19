"""
Session browser — lists all past recordings with processing status.
Click a row to load the session. Bulk process unprocessed meetings.
"""

import datetime
import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, List, Optional

from services.session_service import SessionService
from ui import styles


def _fmt_date(iso_str: Optional[str]) -> str:
    if not iso_str:
        return "—"
    try:
        dt = datetime.datetime.fromisoformat(iso_str)
        return dt.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return iso_str[:16]


def _fmt_duration(seconds: int) -> str:
    if not seconds:
        return "—"
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    if h:
        return f"{h}h {m}m"
    return f"{m}m {s}s" if m else f"{s}s"


class SessionBrowser(tk.Toplevel):
    """Modal dialog showing all past sessions."""

    def __init__(
        self,
        parent,
        session_service: SessionService,
        on_open: Callable[[str], None],
        on_bulk_process: Callable[[List[str]], None],
        recordings_dir: str,
    ):
        super().__init__(parent)
        self._session_service = session_service
        self._on_open = on_open
        self._on_bulk_process = on_bulk_process
        self._recordings_dir = recordings_dir

        self.title("Session History")
        self.geometry("820x540")
        self.minsize(700, 400)
        self.configure(bg=styles.BG_DARK)
        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - 820) // 2
        py = parent.winfo_y() + (parent.winfo_height() - 540) // 2
        self.geometry(f"+{px}+{py}")

        self._sessions: List[dict] = []
        self._build()
        self._refresh()

    def _build(self):
        outer = tk.Frame(self, bg=styles.BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=16, pady=14)

        # Header
        header = tk.Frame(outer, bg=styles.BG_DARK)
        header.pack(fill=tk.X, pady=(0, 10))
        tk.Label(header, text="Session History", bg=styles.BG_DARK,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_HEADER).pack(
                     side=tk.LEFT)
        tk.Button(header, text="Refresh", bg=styles.BG_PANEL,
                  fg=styles.ACCENT, font=styles.FONT_SMALL,
                  relief=tk.FLAT, cursor="hand2",
                  command=self._refresh).pack(side=tk.RIGHT, padx=(6, 0))

        # Treeview of sessions
        tree_frame = tk.Frame(outer, bg=styles.BG_PANEL,
                               highlightbackground=styles.BORDER,
                               highlightthickness=1)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure("Session.Treeview",
                        background=styles.BG_PANEL,
                        foreground=styles.TEXT_PRIMARY,
                        fieldbackground=styles.BG_PANEL,
                        rowheight=26,
                        borderwidth=0,
                        font=styles.FONT_BODY)
        style.configure("Session.Treeview.Heading",
                        background=styles.BG_INPUT,
                        foreground=styles.TEXT_MUTED,
                        font=styles.FONT_SMALL,
                        relief="flat")
        style.map("Session.Treeview",
                  background=[("selected", styles.ACCENT_BG)],
                  foreground=[("selected", styles.ACCENT)])

        columns = ("date", "meeting", "duration", "status")
        self._tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            style="Session.Treeview", selectmode="browse")
        self._tree.heading("date", text="Date")
        self._tree.heading("meeting", text="Meeting")
        self._tree.heading("duration", text="Duration")
        self._tree.heading("status", text="Status")
        self._tree.column("date", width=140, anchor="w")
        self._tree.column("meeting", width=380, anchor="w")
        self._tree.column("duration", width=80, anchor="w")
        self._tree.column("status", width=160, anchor="w")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical",
                                   command=self._tree.yview)
        self._tree.configure(yscrollcommand=scrollbar.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._tree.bind("<Double-1>", self._on_double_click)

        # Footer buttons
        footer = tk.Frame(outer, bg=styles.BG_DARK)
        footer.pack(fill=tk.X, pady=(10, 0))

        self._count_label = tk.Label(footer, text="", bg=styles.BG_DARK,
                                      fg=styles.TEXT_HINT,
                                      font=styles.FONT_SMALL)
        self._count_label.pack(side=tk.LEFT)

        tk.Button(footer, text="Close", bg=styles.BG_PANEL,
                  fg=styles.TEXT_MUTED, font=styles.FONT_BODY,
                  relief=tk.FLAT, padx=18, pady=6, cursor="hand2",
                  command=self.destroy).pack(side=tk.RIGHT)

        tk.Button(footer, text="Open Folder", bg=styles.BG_PANEL,
                  fg=styles.ACCENT, font=styles.FONT_BODY,
                  relief=tk.FLAT, padx=14, pady=6, cursor="hand2",
                  command=self._open_folder).pack(side=tk.RIGHT, padx=(0, 6))

        self._delete_btn = tk.Button(
            footer, text="Delete", bg=styles.BG_PANEL, fg=styles.DANGER,
            font=styles.FONT_BODY, relief=tk.FLAT, padx=14, pady=6,
            cursor="hand2", command=self._delete_selected)
        self._delete_btn.pack(side=tk.RIGHT, padx=(0, 6))

        self._bulk_btn = tk.Button(
            footer, text="Bulk Process", bg=styles.ACCENT, fg="#ffffff",
            font=styles.FONT_BODY, relief=tk.FLAT, padx=14, pady=6,
            cursor="hand2", command=self._bulk_process)
        self._bulk_btn.pack(side=tk.RIGHT, padx=(0, 6))

        self._open_btn = tk.Button(
            footer, text="Open Session", bg=styles.ACCENT, fg="#ffffff",
            font=styles.FONT_BODY, relief=tk.FLAT, padx=14, pady=6,
            cursor="hand2", command=self._open_selected)
        self._open_btn.pack(side=tk.RIGHT, padx=(0, 6))

    def _status_icons(self, s: dict) -> str:
        parts = []
        if s.get("audio_exists"):
            parts.append("🎤")
        if s.get("has_transcript"):
            parts.append("⚙")
        if s.get("has_summary"):
            parts.append("✨")
        if s.get("has_action_items"):
            parts.append("📋")
        if s.get("has_requirements"):
            parts.append("📝")
        return "  ".join(parts) or "—"

    def _refresh(self):
        for row in self._tree.get_children():
            self._tree.delete(row)
        self._sessions = self._session_service.list_sessions()
        for i, s in enumerate(self._sessions):
            self._tree.insert("", "end", iid=str(i), values=(
                _fmt_date(s.get("started_at")),
                s.get("display_name", ""),
                _fmt_duration(s.get("duration_s", 0)),
                self._status_icons(s),
            ))
        unprocessed = sum(1 for s in self._sessions
                          if s.get("audio_exists") and not s.get("has_transcript"))
        self._count_label.config(
            text=f"{len(self._sessions)} sessions • {unprocessed} unprocessed")

    def _selected_session(self) -> Optional[dict]:
        sel = self._tree.selection()
        if not sel:
            return None
        idx = int(sel[0])
        return self._sessions[idx] if 0 <= idx < len(self._sessions) else None

    def _on_double_click(self, event):
        self._open_selected()

    def _open_selected(self):
        s = self._selected_session()
        if not s:
            messagebox.showinfo("No Selection",
                                "Select a session to open.", parent=self)
            return
        self._on_open(s["session_id"])
        self.destroy()

    def _bulk_process(self):
        ids = [s["session_id"] for s in self._sessions
               if s.get("audio_exists") and not s.get("has_transcript")]
        if not ids:
            messagebox.showinfo("Nothing to Process",
                                "All sessions with audio have already been processed.",
                                parent=self)
            return
        if not messagebox.askyesno(
                "Bulk Process",
                f"Process {len(ids)} unprocessed sessions?\n\n"
                "This will run transcription and speaker detection on each. "
                "May take a while depending on audio length.",
                parent=self):
            return
        self._on_bulk_process(ids)
        self.destroy()

    def _delete_selected(self):
        s = self._selected_session()
        if not s:
            messagebox.showinfo("No Selection",
                                "Select a session to delete.", parent=self)
            return
        if not messagebox.askyesno(
                "Delete Session",
                f"Permanently delete this session?\n\n"
                f"{s.get('display_name')}\n"
                f"({_fmt_date(s.get('started_at'))})\n\n"
                "This removes the JSON, WAV, and log files. Cannot be undone.",
                parent=self):
            return
        self._session_service.delete(s["session_id"])
        self._refresh()

    def _open_folder(self):
        folder = os.path.abspath(self._recordings_dir)
        try:
            subprocess.Popen(f'explorer "{folder}"')
        except Exception as e:
            messagebox.showerror("Open Folder", str(e), parent=self)
