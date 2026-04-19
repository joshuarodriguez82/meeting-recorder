"""
Transcript Search — search across every transcript in your recordings.
"""

import datetime
import re
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, List, Optional

from services.session_service import SessionService
from ui import styles


CONTEXT_CHARS = 120  # characters of context around each match


def _fmt_date(iso_str: Optional[str]) -> str:
    if not iso_str:
        return ""
    try:
        return datetime.datetime.fromisoformat(iso_str).strftime("%Y-%m-%d")
    except ValueError:
        return iso_str[:10]


class TranscriptSearch(tk.Toplevel):

    def __init__(
        self,
        parent,
        session_service: SessionService,
        on_open_session: Callable[[str], None],
    ):
        super().__init__(parent)
        self._session_service = session_service
        self._on_open_session = on_open_session
        self._matches: List[dict] = []

        self.title("Transcript Search")
        self.geometry("920x580")
        self.minsize(700, 420)
        self.configure(bg=styles.BG_DARK)
        self.transient(parent)

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - 920) // 2
        py = parent.winfo_y() + (parent.winfo_height() - 580) // 2
        self.geometry(f"+{px}+{py}")

        self._build()

    def _build(self):
        outer = tk.Frame(self, bg=styles.BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        tk.Label(outer, text="Transcript Search", bg=styles.BG_DARK,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_HEADER).pack(
                     anchor="w", pady=(0, 8))

        # Search bar
        bar = tk.Frame(outer, bg=styles.BG_DARK)
        bar.pack(fill=tk.X, pady=(0, 10))

        tk.Label(bar, text="Search", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL).pack(
                     side=tk.LEFT, padx=(0, 6))
        self._query_var = tk.StringVar()
        entry = tk.Entry(
            bar, textvariable=self._query_var,
            bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,
            font=styles.FONT_BODY, relief=tk.FLAT,
            highlightbackground=styles.BORDER, highlightthickness=1)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        entry.focus_set()
        entry.bind("<Return>", lambda e: self._search())

        tk.Button(bar, text="Search", bg=styles.ACCENT, fg="#ffffff",
                  font=styles.FONT_BODY, relief=tk.FLAT, padx=16, pady=5,
                  cursor="hand2", command=self._search).pack(
                      side=tk.LEFT, padx=(6, 0))

        # Results tree
        tree_wrap = tk.Frame(outer, bg=styles.BG_PANEL,
                              highlightbackground=styles.BORDER,
                              highlightthickness=1)
        tree_wrap.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure("TSearch.Treeview",
                        background=styles.BG_PANEL,
                        foreground=styles.TEXT_PRIMARY,
                        fieldbackground=styles.BG_PANEL,
                        rowheight=36, borderwidth=0,
                        font=styles.FONT_SMALL)
        style.configure("TSearch.Treeview.Heading",
                        background=styles.BG_INPUT,
                        foreground=styles.TEXT_MUTED,
                        font=styles.FONT_SMALL, relief="flat")
        style.map("TSearch.Treeview",
                  background=[("selected", styles.ACCENT_BG)],
                  foreground=[("selected", styles.ACCENT)])

        cols = ("date", "meeting", "snippet")
        self._tree = ttk.Treeview(
            tree_wrap, columns=cols, show="headings",
            style="TSearch.Treeview", selectmode="browse")
        self._tree.heading("date", text="Date")
        self._tree.heading("meeting", text="Meeting")
        self._tree.heading("snippet", text="Context")
        self._tree.column("date", width=90, anchor="w")
        self._tree.column("meeting", width=220, anchor="w")
        self._tree.column("snippet", width=560, anchor="w")

        vsb = ttk.Scrollbar(tree_wrap, orient="vertical",
                             command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._tree.bind("<Double-1>", lambda e: self._open_selected())

        # Footer
        footer = tk.Frame(outer, bg=styles.BG_DARK)
        footer.pack(fill=tk.X, pady=(10, 0))
        self._count_label = tk.Label(footer, text="Type a query and hit Enter.",
                                      bg=styles.BG_DARK, fg=styles.TEXT_HINT,
                                      font=styles.FONT_SMALL)
        self._count_label.pack(side=tk.LEFT)
        tk.Button(footer, text="Close", bg=styles.BG_PANEL,
                  fg=styles.TEXT_MUTED, font=styles.FONT_BODY,
                  relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
                  command=self.destroy).pack(side=tk.RIGHT)
        tk.Button(footer, text="Open Meeting", bg=styles.ACCENT,
                  fg="#ffffff", font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=16, pady=6, cursor="hand2",
                  command=self._open_selected).pack(side=tk.RIGHT, padx=(0, 6))

    def _search(self):
        query = self._query_var.get().strip()
        if not query:
            return

        sessions = self._session_service.list_sessions()
        matches = []
        pattern = re.compile(re.escape(query), re.IGNORECASE)

        for s in sessions:
            try:
                session = self._session_service.load_full(s["session_id"])
            except Exception:
                continue
            if not session:
                continue
            transcript = session.full_transcript()
            if not transcript:
                # Also search summaries/action items so non-processed searches work
                text = "\n".join(filter(None, [
                    s.get("summary", ""), s.get("action_items", ""),
                    s.get("decisions", ""), s.get("requirements", "")]))
                if not text or not pattern.search(text):
                    continue
                # Use first match in whatever text exists
                m = pattern.search(text)
                snippet = self._make_snippet(text, m.start(), query)
                matches.append({
                    "session_id": s["session_id"],
                    "display_name": s.get("display_name", ""),
                    "date": _fmt_date(s.get("started_at")),
                    "snippet": snippet,
                })
                continue

            # Find all matches in transcript, keep the best (first for now)
            m = pattern.search(transcript)
            if m:
                snippet = self._make_snippet(transcript, m.start(), query)
                matches.append({
                    "session_id": s["session_id"],
                    "display_name": s.get("display_name", ""),
                    "date": _fmt_date(s.get("started_at")),
                    "snippet": snippet,
                })

        self._matches = matches
        for row in self._tree.get_children():
            self._tree.delete(row)
        for i, m in enumerate(matches):
            self._tree.insert("", "end", iid=str(i), values=(
                m["date"], m["display_name"][:48], m["snippet"]))
        self._count_label.config(
            text=f"{len(matches)} matches across {len(sessions)} sessions")

    def _make_snippet(self, text: str, idx: int, query: str) -> str:
        start = max(0, idx - CONTEXT_CHARS // 2)
        end = min(len(text), idx + len(query) + CONTEXT_CHARS // 2)
        snippet = text[start:end].replace("\n", " ").strip()
        prefix = "…" if start > 0 else ""
        suffix = "…" if end < len(text) else ""
        return f"{prefix}{snippet}{suffix}"

    def _open_selected(self):
        sel = self._tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self._matches):
            self._on_open_session(self._matches[idx]["session_id"])
            self.destroy()
