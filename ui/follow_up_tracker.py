"""
Follow-Up Tracker — aggregates action items across all meetings.
Parses the markdown action items written by Claude and presents them
in a filterable list with owner/status/date/client/project tags.
"""

import datetime
import re
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, List, Optional

from services.session_service import SessionService
from ui import styles


# Pattern matches: "- [ ] **Owner**: Task description (Due: date)"
# or variations with just "- [ ] Task" etc.
CHECKBOX_LINE = re.compile(
    r"^\s*-\s*\[(?P<status>[ xX])\]\s*(?P<rest>.+)$",
    re.MULTILINE,
)
OWNER_BOLD = re.compile(r"\*\*(?P<owner>[^*]+)\*\*\s*:\s*(?P<desc>.+)")
DUE_DATE = re.compile(r"\(Due:\s*(?P<due>[^)]+)\)", re.IGNORECASE)


def parse_action_items(text: str) -> List[dict]:
    """Parse Claude's markdown action items into structured rows."""
    if not text:
        return []
    items = []
    for m in CHECKBOX_LINE.finditer(text):
        status = m.group("status").strip().lower()
        rest = m.group("rest").strip()
        done = status == "x"

        # Extract owner from "**Owner**: Desc"
        owner = ""
        desc = rest
        owner_match = OWNER_BOLD.search(rest)
        if owner_match:
            owner = owner_match.group("owner").strip().strip("[]")
            desc = owner_match.group("desc").strip()

        # Extract due date
        due = ""
        due_match = DUE_DATE.search(desc)
        if due_match:
            due = due_match.group("due").strip()
            desc = DUE_DATE.sub("", desc).strip()

        items.append({
            "done": done,
            "owner": owner,
            "description": desc,
            "due": due,
        })
    return items


def _fmt_date(iso_str: Optional[str]) -> str:
    if not iso_str:
        return ""
    try:
        return datetime.datetime.fromisoformat(iso_str).strftime("%Y-%m-%d")
    except ValueError:
        return iso_str[:10]


class FollowUpTracker(tk.Toplevel):
    """Dialog aggregating all action items across every session."""

    def __init__(
        self,
        parent,
        session_service: SessionService,
        on_open_session: Callable[[str], None],
    ):
        super().__init__(parent)
        self._session_service = session_service
        self._on_open_session = on_open_session
        self._all_items: List[dict] = []

        self.title("Follow-Up Tracker")
        self.geometry("980x620")
        self.minsize(800, 480)
        self.configure(bg=styles.BG_DARK)
        self.transient(parent)

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - 980) // 2
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
        tk.Label(header, text="Follow-Up Tracker", bg=styles.BG_DARK,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_HEADER).pack(
                     side=tk.LEFT)
        tk.Button(header, text="Refresh", bg=styles.BG_PANEL,
                  fg=styles.ACCENT, font=styles.FONT_SMALL,
                  relief=tk.FLAT, cursor="hand2",
                  command=self._refresh).pack(side=tk.RIGHT)

        # Filter bar
        filters = tk.Frame(outer, bg=styles.BG_DARK)
        filters.pack(fill=tk.X, pady=(0, 8))

        tk.Label(filters, text="Status", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL).pack(
                     side=tk.LEFT, padx=(0, 4))
        self._status_var = tk.StringVar(value="Open")
        status_combo = ttk.Combobox(
            filters, textvariable=self._status_var,
            values=["Open", "Done", "All"], state="readonly", width=8)
        status_combo.pack(side=tk.LEFT, padx=(0, 12))
        status_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

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

        tk.Label(filters, text="Owner", bg=styles.BG_DARK,
                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL).pack(
                     side=tk.LEFT, padx=(0, 4))
        self._owner_var = tk.StringVar(value="All")
        self._owner_combo = ttk.Combobox(
            filters, textvariable=self._owner_var,
            values=["All"], state="readonly", width=14)
        self._owner_combo.pack(side=tk.LEFT, padx=(0, 12))
        self._owner_combo.bind("<<ComboboxSelected>>",
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

        # Treeview
        tree_wrap = tk.Frame(outer, bg=styles.BG_PANEL,
                              highlightbackground=styles.BORDER,
                              highlightthickness=1)
        tree_wrap.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure("Track.Treeview",
                        background=styles.BG_PANEL,
                        foreground=styles.TEXT_PRIMARY,
                        fieldbackground=styles.BG_PANEL,
                        rowheight=24,
                        borderwidth=0,
                        font=styles.FONT_SMALL)
        style.configure("Track.Treeview.Heading",
                        background=styles.BG_INPUT,
                        foreground=styles.TEXT_MUTED,
                        font=styles.FONT_SMALL,
                        relief="flat")
        style.map("Track.Treeview",
                  background=[("selected", styles.ACCENT_BG)],
                  foreground=[("selected", styles.ACCENT)])

        cols = ("status", "owner", "description", "due", "client", "meeting", "date")
        self._tree = ttk.Treeview(
            tree_wrap, columns=cols, show="headings",
            style="Track.Treeview", selectmode="browse")
        self._tree.heading("status", text="")
        self._tree.heading("owner", text="Owner")
        self._tree.heading("description", text="Action")
        self._tree.heading("due", text="Due")
        self._tree.heading("client", text="Client")
        self._tree.heading("meeting", text="Meeting")
        self._tree.heading("date", text="Date")
        self._tree.column("status", width=36, anchor="center")
        self._tree.column("owner", width=120, anchor="w")
        self._tree.column("description", width=340, anchor="w")
        self._tree.column("due", width=100, anchor="w")
        self._tree.column("client", width=110, anchor="w")
        self._tree.column("meeting", width=180, anchor="w")
        self._tree.column("date", width=90, anchor="w")

        vsb = ttk.Scrollbar(tree_wrap, orient="vertical",
                             command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._tree.bind("<Double-1>", self._on_double_click)

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
        all_items = []
        clients = set()
        owners = set()
        for s in sessions:
            if not s.get("action_items"):
                continue
            items = parse_action_items(s["action_items"])
            for it in items:
                it["client"] = s.get("client", "") or ""
                it["meeting"] = s.get("display_name", "")
                it["session_id"] = s.get("session_id", "")
                it["session_date"] = _fmt_date(s.get("started_at"))
                all_items.append(it)
                if it["client"]:
                    clients.add(it["client"])
                if it["owner"]:
                    owners.add(it["owner"])

        self._all_items = all_items
        self._client_combo["values"] = ["All"] + sorted(clients)
        self._owner_combo["values"] = ["All"] + sorted(owners)
        self._apply_filters()

    def _apply_filters(self):
        status = self._status_var.get()
        client = self._client_var.get()
        owner = self._owner_var.get()
        search = self._search_var.get().strip().lower()

        filtered = []
        for it in self._all_items:
            if status == "Open" and it["done"]:
                continue
            if status == "Done" and not it["done"]:
                continue
            if client != "All" and it["client"] != client:
                continue
            if owner != "All" and it["owner"] != owner:
                continue
            if search:
                blob = " ".join([
                    it["description"], it["owner"], it["meeting"],
                    it["client"], it["due"]
                ]).lower()
                if search not in blob:
                    continue
            filtered.append(it)

        for row in self._tree.get_children():
            self._tree.delete(row)
        for i, it in enumerate(filtered):
            self._tree.insert("", "end", iid=str(i), values=(
                "✓" if it["done"] else "◯",
                it["owner"],
                it["description"][:120],
                it["due"],
                it["client"],
                it["meeting"][:40],
                it["session_date"],
            ))

        total = len(self._all_items)
        shown = len(filtered)
        open_count = sum(1 for i in self._all_items if not i["done"])
        self._count_label.config(
            text=f"{shown} shown • {open_count} open of {total} total")

        # Keep filtered list mapped by tree iid for click handling
        self._filtered = filtered

    def _on_double_click(self, event):
        self._open_selected()

    def _open_selected(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("No Selection",
                                "Select an action item first.", parent=self)
            return
        idx = int(sel[0])
        if 0 <= idx < len(self._filtered):
            sid = self._filtered[idx].get("session_id")
            if sid:
                self._on_open_session(sid)
                self.destroy()
