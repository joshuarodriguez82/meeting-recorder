"""
Decision Log — aggregates all decisions made across all meetings.
Auto-generated ADR (Architecture Decision Record) database.
"""

import datetime
import re
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, List, Optional

from services.session_service import SessionService
from ui import styles


# Pattern: "## Decision: [title]" with following bullet lines
DECISION_BLOCK = re.compile(
    r"##\s*(?:Decision:?\s*)?(?P<title>.+?)\n(?P<body>(?:[-*].*(?:\n|$))+)",
    re.IGNORECASE,
)
BULLET = re.compile(r"^[-*]\s*\*\*(?P<key>[^:*]+)(?::|\*\*:)\s*\*?\*?\s*(?P<value>.*)$",
                     re.MULTILINE)


def parse_decisions(text: str) -> List[dict]:
    """Parse Claude's decision markdown into structured rows."""
    if not text:
        return []
    if "no decisions made" in text.lower()[:100]:
        return []
    decisions = []
    for m in DECISION_BLOCK.finditer(text):
        title = m.group("title").strip().strip("*#:")
        body = m.group("body")
        fields = {}
        for bm in BULLET.finditer(body):
            key = bm.group("key").strip().lower()
            value = bm.group("value").strip()
            fields[key] = value
        decisions.append({
            "title": title,
            "decided": fields.get("decided", ""),
            "rationale": fields.get("rationale", ""),
            "alternatives": fields.get("alternatives considered", ""),
            "owner": fields.get("owner", ""),
            "impact": fields.get("impact", ""),
        })
    return decisions


def _fmt_date(iso_str: Optional[str]) -> str:
    if not iso_str:
        return ""
    try:
        return datetime.datetime.fromisoformat(iso_str).strftime("%Y-%m-%d")
    except ValueError:
        return iso_str[:10]


class DecisionLog(tk.Toplevel):
    """Dialog showing every decision made across every meeting."""

    def __init__(
        self,
        parent,
        session_service: SessionService,
        on_open_session: Callable[[str], None],
    ):
        super().__init__(parent)
        self._session_service = session_service
        self._on_open_session = on_open_session
        self._all_decisions: List[dict] = []

        self.title("Decision Log")
        self.geometry("1040x620")
        self.minsize(800, 480)
        self.configure(bg=styles.BG_DARK)
        self.transient(parent)

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - 1040) // 2
        py = parent.winfo_y() + (parent.winfo_height() - 620) // 2
        self.geometry(f"+{px}+{py}")

        self._build()
        self._refresh()

    def _build(self):
        outer = tk.Frame(self, bg=styles.BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        # Header
        header = tk.Frame(outer, bg=styles.BG_DARK)
        header.pack(fill=tk.X, pady=(0, 8))
        tk.Label(header, text="Decision Log", bg=styles.BG_DARK,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_HEADER).pack(
                     side=tk.LEFT)
        tk.Button(header, text="Refresh", bg=styles.BG_PANEL,
                  fg=styles.ACCENT, font=styles.FONT_SMALL,
                  relief=tk.FLAT, cursor="hand2",
                  command=self._refresh).pack(side=tk.RIGHT)

        # Filters
        filters = tk.Frame(outer, bg=styles.BG_DARK)
        filters.pack(fill=tk.X, pady=(0, 8))
        tk.Label(filters, text="Client", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL).pack(
                     side=tk.LEFT, padx=(0, 4))
        self._client_var = tk.StringVar(value="All")
        self._client_combo = ttk.Combobox(
            filters, textvariable=self._client_var,
            values=["All"], state="readonly", width=14)
        self._client_combo.pack(side=tk.LEFT, padx=(0, 12))
        self._client_combo.bind("<<ComboboxSelected>>",
                                 lambda e: self._apply_filters())

        tk.Label(filters, text="Search", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL).pack(
                     side=tk.LEFT, padx=(0, 4))
        self._search_var = tk.StringVar()
        search_entry = tk.Entry(
            filters, textvariable=self._search_var,
            bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,
            font=styles.FONT_BODY, relief=tk.FLAT,
            highlightbackground=styles.BORDER, highlightthickness=1)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        self._search_var.trace_add("write", lambda *a: self._apply_filters())

        # Split: Tree on left, detail view on right
        paned = tk.PanedWindow(outer, orient=tk.HORIZONTAL,
                                bg=styles.BG_DARK, sashwidth=6,
                                sashrelief=tk.FLAT, bd=0)
        paned.pack(fill=tk.BOTH, expand=True)

        tree_wrap = tk.Frame(paned, bg=styles.BG_PANEL,
                              highlightbackground=styles.BORDER,
                              highlightthickness=1)
        paned.add(tree_wrap, minsize=300, width=520)

        style = ttk.Style()
        style.configure("Dec.Treeview",
                        background=styles.BG_PANEL,
                        foreground=styles.TEXT_PRIMARY,
                        fieldbackground=styles.BG_PANEL,
                        rowheight=26, borderwidth=0,
                        font=styles.FONT_SMALL)
        style.configure("Dec.Treeview.Heading",
                        background=styles.BG_INPUT,
                        foreground=styles.TEXT_MUTED,
                        font=styles.FONT_SMALL, relief="flat")
        style.map("Dec.Treeview",
                  background=[("selected", styles.ACCENT_BG)],
                  foreground=[("selected", styles.ACCENT)])

        cols = ("title", "client", "date")
        self._tree = ttk.Treeview(
            tree_wrap, columns=cols, show="headings",
            style="Dec.Treeview", selectmode="browse")
        self._tree.heading("title", text="Decision")
        self._tree.heading("client", text="Client")
        self._tree.heading("date", text="Date")
        self._tree.column("title", width=280, anchor="w")
        self._tree.column("client", width=110, anchor="w")
        self._tree.column("date", width=90, anchor="w")

        vsb = ttk.Scrollbar(tree_wrap, orient="vertical",
                             command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._tree.bind("<<TreeviewSelect>>", self._on_select)
        self._tree.bind("<Double-1>", self._on_double_click)

        # Detail pane
        detail_wrap = tk.Frame(paned, bg=styles.BG_PANEL,
                                highlightbackground=styles.BORDER,
                                highlightthickness=1)
        paned.add(detail_wrap, minsize=300)

        self._detail = tk.Text(
            detail_wrap, wrap=tk.WORD, bg=styles.BG_PANEL,
            fg=styles.TEXT_PRIMARY, font=styles.FONT_BODY,
            relief=tk.FLAT, padx=14, pady=10)
        dvsb = ttk.Scrollbar(detail_wrap, orient="vertical",
                              command=self._detail.yview)
        self._detail.configure(yscrollcommand=dvsb.set)
        self._detail.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        dvsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._detail.insert("1.0", "Select a decision to see details.")
        self._detail.configure(state=tk.DISABLED)

        # Footer
        footer = tk.Frame(outer, bg=styles.BG_DARK)
        footer.pack(fill=tk.X, pady=(8, 0))
        self._count_label = tk.Label(footer, text="", bg=styles.BG_DARK,
                                      fg=styles.TEXT_HINT,
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

    def _refresh(self):
        sessions = self._session_service.list_sessions()
        all_decisions = []
        clients = set()
        for s in sessions:
            if not s.get("decisions"):
                continue
            for d in parse_decisions(s["decisions"]):
                d["client"] = s.get("client", "")
                d["meeting"] = s.get("display_name", "")
                d["session_id"] = s.get("session_id", "")
                d["session_date"] = _fmt_date(s.get("started_at"))
                all_decisions.append(d)
                if d["client"]:
                    clients.add(d["client"])

        self._all_decisions = all_decisions
        self._client_combo["values"] = ["All"] + sorted(clients)
        self._apply_filters()

    def _apply_filters(self):
        client = self._client_var.get()
        search = self._search_var.get().strip().lower()
        filtered = []
        for d in self._all_decisions:
            if client != "All" and d["client"] != client:
                continue
            if search:
                blob = " ".join([d["title"], d["decided"], d["rationale"],
                                  d["alternatives"], d["owner"], d["impact"],
                                  d["meeting"], d["client"]]).lower()
                if search not in blob:
                    continue
            filtered.append(d)

        for row in self._tree.get_children():
            self._tree.delete(row)
        for i, d in enumerate(filtered):
            self._tree.insert("", "end", iid=str(i), values=(
                d["title"][:60], d["client"], d["session_date"]))

        self._filtered = filtered
        self._count_label.config(
            text=f"{len(filtered)} shown of {len(self._all_decisions)} total")

    def _on_select(self, event):
        sel = self._tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if idx < 0 or idx >= len(self._filtered):
            return
        d = self._filtered[idx]
        text = f"{d['title']}\n{'=' * len(d['title'])}\n\n"
        if d.get("decided"):
            text += f"DECIDED:\n{d['decided']}\n\n"
        if d.get("rationale"):
            text += f"RATIONALE:\n{d['rationale']}\n\n"
        if d.get("alternatives"):
            text += f"ALTERNATIVES CONSIDERED:\n{d['alternatives']}\n\n"
        if d.get("owner"):
            text += f"OWNER: {d['owner']}\n"
        if d.get("impact"):
            text += f"IMPACT: {d['impact']}\n"
        text += f"\n───\nFrom meeting: {d.get('meeting', '')}  ({d.get('session_date', '')})"
        self._detail.configure(state=tk.NORMAL)
        self._detail.delete("1.0", tk.END)
        self._detail.insert("1.0", text)
        self._detail.configure(state=tk.DISABLED)

    def _on_double_click(self, event):
        self._open_selected()

    def _open_selected(self):
        sel = self._tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self._filtered):
            sid = self._filtered[idx].get("session_id")
            if sid:
                self._on_open_session(sid)
                self.destroy()
