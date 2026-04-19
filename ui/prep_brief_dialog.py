"""
Meeting Prep Brief — generates a pre-meeting brief from prior session notes.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, List, Optional

from ui import styles


class PrepBriefDialog(tk.Toplevel):

    def __init__(
        self,
        parent,
        subject: str,
        related_sessions: List[dict],
        generate_brief: Callable[[str, str], str],
    ):
        """
        related_sessions: list of session summary dicts (from list_sessions)
                          — must include display_name, summary, action_items,
                          requirements fields
        generate_brief: async function (prior_notes, upcoming_subject) -> brief
        """
        super().__init__(parent)
        self._subject = subject
        self._related = related_sessions
        self._generate_brief = generate_brief

        self.title("Meeting Prep Brief")
        self.geometry("680x540")
        self.minsize(600, 400)
        self.configure(bg=styles.BG_DARK)
        self.transient(parent)

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - 680) // 2
        py = parent.winfo_y() + (parent.winfo_height() - 540) // 2
        self.geometry(f"+{px}+{py}")

        self._build()
        self._generate()

    def _build(self):
        outer = tk.Frame(self, bg=styles.BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        tk.Label(outer, text="Meeting Prep Brief",
                 bg=styles.BG_DARK, fg=styles.TEXT_PRIMARY,
                 font=styles.FONT_HEADER).pack(anchor="w")
        tk.Label(outer, text=self._subject, bg=styles.BG_DARK,
                 fg=styles.ACCENT, font=styles.FONT_TITLE).pack(anchor="w")
        tk.Label(outer,
                 text=f"Based on {len(self._related)} prior meeting(s)",
                 bg=styles.BG_DARK, fg=styles.TEXT_HINT,
                 font=styles.FONT_SMALL).pack(anchor="w", pady=(0, 10))

        text_frame = tk.Frame(outer, bg=styles.BG_PANEL,
                               highlightbackground=styles.BORDER,
                               highlightthickness=1)
        text_frame.pack(fill=tk.BOTH, expand=True)

        self._text = tk.Text(
            text_frame, wrap=tk.WORD, bg=styles.BG_PANEL,
            fg=styles.TEXT_PRIMARY, font=styles.FONT_BODY,
            relief=tk.FLAT, padx=14, pady=10)
        vsb = ttk.Scrollbar(text_frame, orient="vertical",
                             command=self._text.yview)
        self._text.configure(yscrollcommand=vsb.set)
        self._text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._text.insert("1.0", "Generating brief from prior meetings...\n")
        self._text.configure(state=tk.DISABLED)

        footer = tk.Frame(outer, bg=styles.BG_DARK)
        footer.pack(fill=tk.X, pady=(10, 0))
        tk.Button(footer, text="Copy to Clipboard", bg=styles.BG_PANEL,
                  fg=styles.ACCENT, font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=14, pady=6, cursor="hand2",
                  command=self._copy).pack(side=tk.LEFT)
        tk.Button(footer, text="Close", bg=styles.BG_PANEL,
                  fg=styles.TEXT_MUTED, font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=16, pady=6, cursor="hand2",
                  command=self.destroy).pack(side=tk.RIGHT)

    def _build_prior_notes(self) -> str:
        parts = []
        for s in self._related[:8]:  # cap to avoid huge prompts
            block = [f"### {s.get('display_name', 'Meeting')} "
                     f"({(s.get('started_at') or '')[:10]})"]
            if s.get("summary"):
                block.append(f"**Summary:**\n{s['summary']}")
            if s.get("action_items"):
                block.append(f"**Action Items:**\n{s['action_items']}")
            if s.get("requirements"):
                block.append(f"**Requirements:**\n{s['requirements']}")
            parts.append("\n\n".join(block))
        return "\n\n---\n\n".join(parts)

    def _generate(self):
        if not self._related:
            self._set_text("No prior meetings found for this client/project "
                           "yet. Nothing to brief on.")
            return

        prior_notes = self._build_prior_notes()

        # Fire the generator. generate_brief is a callable that handles
        # async execution and calls back with the result. Caller owns threading.
        def on_result(brief: str):
            self.after(0, lambda: self._set_text(brief))

        def on_error(err: Exception):
            self.after(0, lambda: self._set_text(f"Error generating brief:\n{err}"))

        # Use the provided generator (which takes care of threading)
        self._generate_brief(prior_notes, self._subject, on_result, on_error)

    def _set_text(self, content: str):
        self._text.configure(state=tk.NORMAL)
        self._text.delete("1.0", tk.END)
        self._text.insert("1.0", content)
        self._text.configure(state=tk.DISABLED)

    def _copy(self):
        try:
            self.clipboard_clear()
            self.clipboard_append(self._text.get("1.0", tk.END).strip())
        except Exception as e:
            messagebox.showerror("Copy Failed", str(e), parent=self)
