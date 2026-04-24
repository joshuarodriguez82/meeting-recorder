"""
Client folder manager dialog.

Lets the user map each client (tag) to a designated folder. When the main
window runs an export, sessions tagged with that client are written to the
mapped folder instead of the default recordings directory.

Configuration persists in <recordings_dir>/clients.json via ClientService.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable, List, Optional

from services.client_service import ClientService
from ui import styles


class ClientManagerDialog(tk.Toplevel):

    def __init__(
        self,
        parent,
        client_service: ClientService,
        known_clients: List[str],
        on_change: Callable[[], None],
    ):
        super().__init__(parent)
        self._svc = client_service
        self._on_change = on_change
        # Seed the editor with every client we already know about, even if
        # not yet in clients.json — easier to assign folders to existing tags.
        stored = self._svc.load()
        stored_names = {c["name"].lower() for c in stored}
        rows: List[dict] = list(stored)
        for name in known_clients:
            if name and name.lower() not in stored_names:
                rows.append({"name": name, "folder": ""})
                stored_names.add(name.lower())
        self._rows = rows

        self.title("Manage Client Folders")
        self.geometry("720x460")
        self.minsize(600, 360)
        self.configure(bg=styles.BG_DARK)
        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - 720) // 2
        py = parent.winfo_y() + (parent.winfo_height() - 460) // 2
        self.geometry(f"+{px}+{py}")

        self._build()
        self._refresh()

    def _build(self) -> None:
        outer = tk.Frame(self, bg=styles.BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        tk.Label(outer, text="Client Folders", bg=styles.BG_DARK,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_HEADER).pack(
                     anchor="w")
        tk.Label(outer,
                 text=("When a recording is tagged with a client, its "
                       "transcript, summary, action items, decisions and "
                       "requirements are saved to that client's folder."),
                 bg=styles.BG_DARK, fg=styles.TEXT_HINT,
                 font=styles.FONT_SMALL, wraplength=680, justify="left").pack(
                     anchor="w", pady=(2, 10))

        frame = tk.Frame(outer, bg=styles.BG_PANEL,
                         highlightbackground=styles.BORDER,
                         highlightthickness=1)
        frame.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure("Client.Treeview",
                        background=styles.BG_PANEL,
                        foreground=styles.TEXT_PRIMARY,
                        fieldbackground=styles.BG_PANEL,
                        rowheight=26, borderwidth=0,
                        font=styles.FONT_BODY)
        style.configure("Client.Treeview.Heading",
                        background=styles.BG_INPUT,
                        foreground=styles.TEXT_MUTED,
                        font=styles.FONT_SMALL, relief="flat")
        style.map("Client.Treeview",
                  background=[("selected", styles.ACCENT_BG)],
                  foreground=[("selected", styles.ACCENT)])

        self._tree = ttk.Treeview(
            frame, columns=("name", "folder"), show="headings",
            style="Client.Treeview", selectmode="browse")
        self._tree.heading("name", text="Client")
        self._tree.heading("folder", text="Designated Folder")
        self._tree.column("name", width=180, anchor="w")
        self._tree.column("folder", width=490, anchor="w")
        scrollbar = ttk.Scrollbar(frame, orient="vertical",
                                  command=self._tree.yview)
        self._tree.configure(yscrollcommand=scrollbar.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._tree.bind("<Double-1>", lambda e: self._browse_for_selected())

        # Buttons row
        btns = tk.Frame(outer, bg=styles.BG_DARK)
        btns.pack(fill=tk.X, pady=(10, 0))

        tk.Button(btns, text="+ Add Client", bg=styles.BG_PANEL,
                  fg=styles.ACCENT, font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=14, pady=6, cursor="hand2",
                  command=self._add_client).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(btns, text="Browse Folder…", bg=styles.BG_PANEL,
                  fg=styles.ACCENT, font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=14, pady=6, cursor="hand2",
                  command=self._browse_for_selected).pack(
                      side=tk.LEFT, padx=(0, 6))

        tk.Button(btns, text="Clear Folder", bg=styles.BG_PANEL,
                  fg=styles.TEXT_MUTED, font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=14, pady=6, cursor="hand2",
                  command=self._clear_folder).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(btns, text="Remove", bg=styles.BG_PANEL,
                  fg=styles.DANGER, font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=14, pady=6, cursor="hand2",
                  command=self._remove_selected).pack(side=tk.LEFT)

        tk.Button(btns, text="Save", bg=styles.ACCENT, fg="#ffffff",
                  font=styles.FONT_BODY, relief=tk.FLAT, padx=18, pady=6,
                  cursor="hand2",
                  command=self._save_and_close).pack(side=tk.RIGHT)
        tk.Button(btns, text="Cancel", bg=styles.BG_PANEL,
                  fg=styles.TEXT_MUTED, font=styles.FONT_BODY, relief=tk.FLAT,
                  padx=14, pady=6, cursor="hand2",
                  command=self.destroy).pack(side=tk.RIGHT, padx=(0, 6))

    def _refresh(self) -> None:
        for row in self._tree.get_children():
            self._tree.delete(row)
        self._rows.sort(key=lambda r: r.get("name", "").lower())
        for i, entry in enumerate(self._rows):
            folder = entry.get("folder", "") or "(default recordings folder)"
            self._tree.insert("", "end", iid=str(i),
                              values=(entry.get("name", ""), folder))

    def _selected_index(self) -> Optional[int]:
        sel = self._tree.selection()
        if not sel:
            return None
        idx = int(sel[0])
        return idx if 0 <= idx < len(self._rows) else None

    def _add_client(self) -> None:
        dlg = tk.Toplevel(self)
        dlg.title("Add Client")
        dlg.configure(bg=styles.BG_PANEL)
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)
        tk.Label(dlg, text="Client name:", bg=styles.BG_PANEL,
                 fg=styles.TEXT_PRIMARY, font=styles.FONT_BODY).pack(
                     anchor="w", padx=16, pady=(14, 4))
        name_var = tk.StringVar()
        entry = tk.Entry(dlg, textvariable=name_var,
                         bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,
                         insertbackground=styles.TEXT_PRIMARY,
                         font=styles.FONT_BODY, relief=tk.FLAT, width=34)
        entry.pack(padx=16, ipady=6)
        entry.focus_set()

        def _commit(*_):
            name = name_var.get().strip()
            if not name:
                dlg.destroy()
                return
            if any(r.get("name", "").lower() == name.lower()
                   for r in self._rows):
                messagebox.showerror(
                    "Duplicate", f"Client '{name}' already exists.",
                    parent=dlg)
                return
            self._rows.append({"name": name, "folder": ""})
            self._refresh()
            for i, r in enumerate(self._rows):
                if r.get("name") == name:
                    self._tree.selection_set(str(i))
                    break
            dlg.destroy()

        entry.bind("<Return>", _commit)
        btn_row = tk.Frame(dlg, bg=styles.BG_PANEL)
        btn_row.pack(fill=tk.X, padx=16, pady=(10, 14))
        tk.Button(btn_row, text="Cancel", bg=styles.BG_PANEL,
                  fg=styles.TEXT_MUTED, relief=tk.FLAT, font=styles.FONT_BODY,
                  padx=12, pady=5, cursor="hand2",
                  command=dlg.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        tk.Button(btn_row, text="Add", bg=styles.ACCENT, fg="#ffffff",
                  relief=tk.FLAT, font=styles.FONT_BODY, padx=14, pady=5,
                  cursor="hand2", command=_commit).pack(side=tk.RIGHT)

    def _browse_for_selected(self) -> None:
        idx = self._selected_index()
        if idx is None:
            messagebox.showinfo(
                "No Selection",
                "Select a client first (or click + Add Client).",
                parent=self)
            return
        start_dir = self._rows[idx].get("folder") or os.path.expanduser("~")
        if not os.path.isdir(start_dir):
            start_dir = os.path.expanduser("~")
        chosen = filedialog.askdirectory(
            parent=self,
            title=f"Choose folder for {self._rows[idx]['name']}",
            initialdir=start_dir,
            mustexist=True,
        )
        if not chosen:
            return
        self._rows[idx]["folder"] = chosen
        self._refresh()
        self._tree.selection_set(str(idx))

    def _clear_folder(self) -> None:
        idx = self._selected_index()
        if idx is None:
            return
        self._rows[idx]["folder"] = ""
        self._refresh()
        self._tree.selection_set(str(idx))

    def _remove_selected(self) -> None:
        idx = self._selected_index()
        if idx is None:
            return
        name = self._rows[idx].get("name", "")
        if not messagebox.askyesno(
                "Remove Client",
                f"Remove '{name}' from the client folder list?\n\n"
                "This only removes the folder mapping — existing "
                "session tags are untouched.",
                parent=self):
            return
        del self._rows[idx]
        self._refresh()

    def _save_and_close(self) -> None:
        try:
            self._svc.save(self._rows)
        except Exception as e:
            messagebox.showerror("Save Failed", str(e), parent=self)
            return
        try:
            self._on_change()
        except Exception:
            pass
        self.destroy()
