import tkinter as tk
from ui import styles


class TranscriptPanel(tk.Frame):

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=styles.BG_PANEL, **kwargs)
        self._build()

    def _build(self):
        scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL,
                                  bg=styles.BG_INPUT,
                                  troughcolor=styles.BG_PANEL,
                                  relief=tk.FLAT)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._text = tk.Text(
            self,
            bg=styles.BG_PANEL,
            fg=styles.TEXT_PRIMARY,
            font=styles.FONT_MONO,
            relief=tk.FLAT,
            wrap=tk.WORD,
            state=tk.DISABLED,
            yscrollcommand=scrollbar.set,
            padx=6,
            pady=6,
            selectbackground=styles.ACCENT_BG,
            selectforeground=styles.ACCENT,
            insertbackground=styles.ACCENT,
            spacing1=4,
            spacing3=4,
        )
        self._text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self._text.yview)

        self._text.tag_configure(
            "speaker", foreground=styles.ACCENT,
            font=("Segoe UI", 10, "bold"))
        self._text.tag_configure(
            "timestamp", foreground=styles.TEXT_HINT,
            font=("Segoe UI", 9))
        self._text.tag_configure(
            "divider", foreground=styles.ACCENT_DIM,
            font=("Segoe UI", 10, "bold"))

    def set_text(self, content: str) -> None:
        self._text.config(state=tk.NORMAL)
        self._text.delete("1.0", tk.END)
        for line in content.splitlines():
            if line.startswith("──"):
                self._text.insert(tk.END, line + "\n", "divider")
            elif "]" in line and "→" in line:
                parts = line.split("] ", 1)
                if len(parts) == 2:
                    ts_part = parts[0] + "] "
                    rest = parts[1]
                    if ": " in rest:
                        speaker, text = rest.split(": ", 1)
                        self._text.insert(tk.END, ts_part, "timestamp")
                        self._text.insert(tk.END, speaker + ": ", "speaker")
                        self._text.insert(tk.END, text + "\n")
                    else:
                        self._text.insert(tk.END, line + "\n")
                else:
                    self._text.insert(tk.END, line + "\n")
            else:
                self._text.insert(tk.END, line + "\n")
        self._text.config(state=tk.DISABLED)
        self._text.see(tk.END)

    def append_line(self, line: str) -> None:
        self._text.config(state=tk.NORMAL)
        self._text.insert(tk.END, line + "\n")
        self._text.config(state=tk.DISABLED)
        self._text.see(tk.END)

    def clear(self) -> None:
        self._text.config(state=tk.NORMAL)
        self._text.delete("1.0", tk.END)
        self._text.config(state=tk.DISABLED)
