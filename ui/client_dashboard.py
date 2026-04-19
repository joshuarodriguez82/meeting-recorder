"""
Client Dashboard — per-client overview: meetings, open action items,
recent decisions, stats.
"""

import datetime
import tkinter as tk
from tkinter import ttk
from typing import Callable, Dict, List, Optional

from services.session_service import SessionService
from ui import styles
from ui.follow_up_tracker import parse_action_items
from ui.decision_log import parse_decisions


def _fmt_date(iso_str: Optional[str]) -> str:
    if not iso_str:
        return "—"
    try:
        return datetime.datetime.fromisoformat(iso_str).strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return iso_str[:16]


def _fmt_hours(seconds: int) -> str:
    if seconds <= 0:
        return "0h"
    hours = seconds / 3600
    return f"{hours:.1f}h"


class ClientDashboard(tk.Toplevel):

    def __init__(
        self,
        parent,
        session_service: SessionService,
        on_open_session: Callable[[str], None],
    ):
        super().__init__(parent)
        self._session_service = session_service
        self._on_open_session = on_open_session
        self._all_sessions: List[dict] = []

        self.title("Client Dashboard")
        self.geometry("1020x640")
        self.minsize(820, 480)
        self.configure(bg=styles.BG_DARK)
        self.transient(parent)

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - 1020) // 2
        py = parent.winfo_y() + (parent.winfo_height() - 640) // 2
        self.geometry(f"+{px}+{py}")

        self._build()
        self._refresh()

    def _build(self):
        outer = tk.Frame(self, bg=styles.BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        header = tk.Frame(outer, bg=styles.BG_DARK)
        header.pack(fill=tk.X, pady=(0, 10))
        tk.Label(header, text="Client Dashboard", bg=styles.BG_DARK,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_HEADER).pack(
                     side=tk.LEFT)
        tk.Button(header, text="Refresh", bg=styles.BG_PANEL,
                  fg=styles.ACCENT, font=styles.FONT_SMALL,
                  relief=tk.FLAT, cursor="hand2",
                  command=self._refresh).pack(side=tk.RIGHT)

        selector = tk.Frame(outer, bg=styles.BG_DARK)
        selector.pack(fill=tk.X, pady=(0, 10))
        tk.Label(selector, text="Client", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL).pack(
                     side=tk.LEFT, padx=(0, 6))
        self._client_var = tk.StringVar()
        self._client_combo = ttk.Combobox(
            selector, textvariable=self._client_var,
            values=[], state="readonly", width=28)
        self._client_combo.pack(side=tk.LEFT)
        self._client_combo.bind("<<ComboboxSelected>>",
                                 lambda e: self._render())

        # Stats strip
        self._stats_frame = tk.Frame(outer, bg=styles.BG_DARK)
        self._stats_frame.pack(fill=tk.X, pady=(0, 10))

        # Main body: two columns (left meetings, right action items + decisions)
        body = tk.Frame(outer, bg=styles.BG_DARK)
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(0, weight=1, uniform="col")
        body.columnconfigure(1, weight=1, uniform="col")
        body.rowconfigure(0, weight=1)

        # Left: meetings tree
        left = tk.Frame(body, bg=styles.BG_PANEL,
                         highlightbackground=styles.BORDER,
                         highlightthickness=1)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 4))

        tk.Label(left, text="MEETINGS", bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=styles.FONT_SMALL).pack(
                     anchor="w", padx=10, pady=(8, 4))

        mt_frame = tk.Frame(left, bg=styles.BG_PANEL)
        mt_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        style = ttk.Style()
        style.configure("CD.Treeview",
                        background=styles.BG_PANEL,
                        foreground=styles.TEXT_PRIMARY,
                        fieldbackground=styles.BG_PANEL,
                        rowheight=24, borderwidth=0,
                        font=styles.FONT_SMALL)
        style.configure("CD.Treeview.Heading",
                        background=styles.BG_INPUT,
                        foreground=styles.TEXT_MUTED,
                        font=styles.FONT_SMALL, relief="flat")
        style.map("CD.Treeview",
                  background=[("selected", styles.ACCENT_BG)],
                  foreground=[("selected", styles.ACCENT)])

        self._meetings_tree = ttk.Treeview(
            mt_frame, columns=("date", "meeting", "project"), show="headings",
            style="CD.Treeview", selectmode="browse")
        self._meetings_tree.heading("date", text="Date")
        self._meetings_tree.heading("meeting", text="Meeting")
        self._meetings_tree.heading("project", text="Project")
        self._meetings_tree.column("date", width=110, anchor="w")
        self._meetings_tree.column("meeting", width=240, anchor="w")
        self._meetings_tree.column("project", width=100, anchor="w")
        mvsb = ttk.Scrollbar(mt_frame, orient="vertical",
                              command=self._meetings_tree.yview)
        self._meetings_tree.configure(yscrollcommand=mvsb.set)
        self._meetings_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        mvsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._meetings_tree.bind("<Double-1>",
                                  lambda e: self._open_meeting())

        # Right: action items + decisions stacked
        right = tk.Frame(body, bg=styles.BG_DARK)
        right.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        right.rowconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        ai_card = tk.Frame(right, bg=styles.BG_PANEL,
                            highlightbackground=styles.BORDER,
                            highlightthickness=1)
        ai_card.grid(row=0, column=0, sticky="nsew", pady=(0, 4))
        tk.Label(ai_card, text="OPEN ACTION ITEMS", bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=styles.FONT_SMALL).pack(
                     anchor="w", padx=10, pady=(8, 4))
        ai_frame = tk.Frame(ai_card, bg=styles.BG_PANEL)
        ai_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        self._ai_text = tk.Text(ai_frame, wrap=tk.WORD, bg=styles.BG_PANEL,
                                  fg=styles.TEXT_PRIMARY, font=styles.FONT_SMALL,
                                  relief=tk.FLAT, padx=8, pady=6)
        aivsb = ttk.Scrollbar(ai_frame, orient="vertical",
                               command=self._ai_text.yview)
        self._ai_text.configure(yscrollcommand=aivsb.set)
        self._ai_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        aivsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._ai_text.configure(state=tk.DISABLED)

        dec_card = tk.Frame(right, bg=styles.BG_PANEL,
                             highlightbackground=styles.BORDER,
                             highlightthickness=1)
        dec_card.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        tk.Label(dec_card, text="RECENT DECISIONS", bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=styles.FONT_SMALL).pack(
                     anchor="w", padx=10, pady=(8, 4))
        dec_frame = tk.Frame(dec_card, bg=styles.BG_PANEL)
        dec_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        self._dec_text = tk.Text(dec_frame, wrap=tk.WORD, bg=styles.BG_PANEL,
                                   fg=styles.TEXT_PRIMARY, font=styles.FONT_SMALL,
                                   relief=tk.FLAT, padx=8, pady=6)
        dvsb = ttk.Scrollbar(dec_frame, orient="vertical",
                              command=self._dec_text.yview)
        self._dec_text.configure(yscrollcommand=dvsb.set)
        self._dec_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        dvsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._dec_text.configure(state=tk.DISABLED)

        # Footer
        footer = tk.Frame(outer, bg=styles.BG_DARK)
        footer.pack(fill=tk.X, pady=(10, 0))
        tk.Button(footer, text="Close", bg=styles.BG_PANEL,
                  fg=styles.TEXT_MUTED, font=styles.FONT_BODY,
                  relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
                  command=self.destroy).pack(side=tk.RIGHT)
        tk.Button(footer, text="Open Selected Meeting", bg=styles.ACCENT,
                  fg="#ffffff", font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=16, pady=6, cursor="hand2",
                  command=self._open_meeting).pack(side=tk.RIGHT, padx=(0, 6))

    def _refresh(self):
        self._all_sessions = self._session_service.list_sessions()
        clients = sorted({s.get("client", "") for s in self._all_sessions
                           if s.get("client", "").strip()})
        if not clients:
            clients = ["(no clients tagged yet)"]
        self._client_combo["values"] = clients
        if self._client_var.get() not in clients:
            self._client_combo.current(0)
        self._render()

    def _render(self):
        client = self._client_var.get()
        client_sessions = [s for s in self._all_sessions
                            if s.get("client", "") == client]

        # Stats
        for w in self._stats_frame.winfo_children():
            w.destroy()
        total_sessions = len(client_sessions)
        total_seconds = sum(s.get("duration_s", 0) for s in client_sessions)
        projects = len({s.get("project", "") for s in client_sessions
                         if s.get("project", "")})
        self._stat_pill(self._stats_frame, "Meetings", str(total_sessions))
        self._stat_pill(self._stats_frame, "Total Time", _fmt_hours(total_seconds))
        self._stat_pill(self._stats_frame, "Projects", str(projects))

        # Open action items (tagged with this client)
        open_items = []
        for s in client_sessions:
            if not s.get("action_items"):
                continue
            for it in parse_action_items(s["action_items"]):
                if not it["done"]:
                    it["meeting"] = s.get("display_name", "")
                    it["session_id"] = s.get("session_id", "")
                    it["session_date"] = _fmt_date(s.get("started_at"))
                    open_items.append(it)
        self._stat_pill(self._stats_frame, "Open Actions", str(len(open_items)))

        # Decisions (across this client's sessions)
        decisions = []
        for s in client_sessions:
            if not s.get("decisions"):
                continue
            for d in parse_decisions(s["decisions"]):
                d["meeting"] = s.get("display_name", "")
                d["session_id"] = s.get("session_id", "")
                d["session_date"] = _fmt_date(s.get("started_at"))
                decisions.append(d)
        self._stat_pill(self._stats_frame, "Decisions", str(len(decisions)))

        # Meetings tree
        for row in self._meetings_tree.get_children():
            self._meetings_tree.delete(row)
        self._client_sessions = client_sessions
        for i, s in enumerate(client_sessions):
            self._meetings_tree.insert("", "end", iid=str(i), values=(
                _fmt_date(s.get("started_at")),
                s.get("display_name", "")[:48],
                s.get("project", ""),
            ))

        # Action items text
        self._ai_text.configure(state=tk.NORMAL)
        self._ai_text.delete("1.0", tk.END)
        if open_items:
            for it in open_items[:50]:
                owner = f"[{it['owner']}] " if it['owner'] else ""
                due = f" (Due: {it['due']})" if it['due'] else ""
                self._ai_text.insert(tk.END,
                                       f"• {owner}{it['description']}{due}\n")
                self._ai_text.insert(tk.END,
                                       f"    — {it['meeting']} ({it['session_date']})\n\n")
        else:
            self._ai_text.insert(tk.END, "No open action items.")
        self._ai_text.configure(state=tk.DISABLED)

        # Decisions text (most recent first)
        self._dec_text.configure(state=tk.NORMAL)
        self._dec_text.delete("1.0", tk.END)
        if decisions:
            for d in decisions[:20]:
                self._dec_text.insert(tk.END, f"• {d['title']}\n")
                if d.get("decided"):
                    self._dec_text.insert(tk.END, f"    {d['decided'][:180]}\n")
                self._dec_text.insert(tk.END,
                                        f"    — {d['meeting']} ({d['session_date']})\n\n")
        else:
            self._dec_text.insert(tk.END, "No decisions logged for this client.")
        self._dec_text.configure(state=tk.DISABLED)

    def _stat_pill(self, parent, label: str, value: str):
        pill = tk.Frame(parent, bg=styles.BG_PANEL,
                         highlightbackground=styles.BORDER,
                         highlightthickness=1)
        pill.pack(side=tk.LEFT, padx=(0, 8), ipadx=10, ipady=4)
        tk.Label(pill, text=value, bg=styles.BG_PANEL,
                 fg=styles.ACCENT, font=styles.FONT_TITLE).pack(anchor="w", padx=8)
        tk.Label(pill, text=label, bg=styles.BG_PANEL,
                 fg=styles.TEXT_HINT, font=styles.FONT_SMALL).pack(
                     anchor="w", padx=8, pady=(0, 2))

    def _open_meeting(self):
        sel = self._meetings_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self._client_sessions):
            sid = self._client_sessions[idx]["session_id"]
            self._on_open_session(sid)
            self.destroy()
