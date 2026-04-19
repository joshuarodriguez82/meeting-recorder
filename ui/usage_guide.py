"""
Usage Guide — in-app help dialog with a walkthrough of every feature.
"""

import tkinter as tk
from tkinter import ttk

from ui import styles


GUIDE_TEXT = """
Meeting Recorder — Usage Guide

═══════════════════════════════════════════════════════════
  RECORD A MEETING
═══════════════════════════════════════════════════════════

1. Pick a meeting from TODAY'S MEETINGS, or type a name in the
   Meeting field at the top.
2. (Optional) Set Client and Project tags — these power the
   Client Dashboard, Follow-Up Tracker, and Meeting Prep Brief.
3. Pick a Template: General, Requirements Gathering, Design
   Review, Sprint Planning, or Stakeholder Update. This tailors
   the AI summary to that meeting type.
4. Click "Start Recording" (red button).
5. When done, click "Stop Recording". Audio is saved to the
   recordings folder.

  Tip: System Audio must be set to a loopback device (Stereo Mix,
  VB-Cable, etc.) for the recorder to capture the other meeting
  participants. Mic-only will only capture your voice.

═══════════════════════════════════════════════════════════
  PROCESS & EXTRACT
═══════════════════════════════════════════════════════════

After recording stops, click the AI buttons in order (or turn on
Auto-Process in Settings to do it all automatically):

  ⚙  Process        — Transcribes audio + identifies speakers
  ✨  Summarize      — AI summary based on the template
  📋  Action Items   — Who needs to do what, by when
  📝  Requirements   — FR/NFR tables with priority + owner
  🎯  Decisions      — Decision log with rationale (ADR style)

Each output shows up in the transcript panel stacked under the
raw transcript. Results are exported to text files in the
recordings folder automatically.

═══════════════════════════════════════════════════════════
  SHARING
═══════════════════════════════════════════════════════════

  💾  Export       — Saves transcript, summary, action items,
                    requirements, and decisions as text files.
  ✉  Email        — Drafts a structured email in Outlook with
                    summary, action items, and decisions.
  📁  Recordings   — Opens the recordings folder in Explorer.

  If "Auto-draft follow-up email" is on in Settings, an email
  draft is created in Outlook after processing completes — addressed
  to all calendar attendees.

═══════════════════════════════════════════════════════════
  CALENDAR INTEGRATION
═══════════════════════════════════════════════════════════

Today's meetings from Outlook appear in the TODAY'S MEETINGS card.
Click "Record" on any row to start a recording with the meeting
name and attendees pre-filled.

A popup notification appears 2 minutes before each meeting (adjust
in Settings > Calendar > Notify Before). Click "Start Recording"
in the popup to begin; click "Dismiss" to skip that meeting.

═══════════════════════════════════════════════════════════
  SESSIONS MENU
═══════════════════════════════════════════════════════════

  Session History      — All past recordings with status icons
                         (🎤 audio, ⚙ transcript, ✨ summary,
                         📋 actions, 📝 requirements, 🎯 decisions).
                         Double-click a row to load it. Bulk Process
                         runs all unprocessed meetings at once.

  Follow-Up Tracker    — Every action item from every meeting,
                         filterable by owner, client, status, or
                         search text. Double-click an item to jump
                         to its meeting.

  Decision Log         — Auto-generated Architecture Decision Record
                         database. Every decision ever made across
                         your meetings, with rationale and owner.

  Transcript Search    — "What did we decide about auth?" Searches
                         every transcript, returns context snippets.

  Client Dashboard     — Per-client view: all meetings, open action
                         items, recent decisions, total hours
                         recorded, project count.

  Meeting Prep Brief   — Before a scheduled meeting, generates a
                         brief from past meetings tagged with the
                         same client/project: recent context, open
                         items, risks, suggested discussion points.

═══════════════════════════════════════════════════════════
  SETTINGS (File > Settings)
═══════════════════════════════════════════════════════════

  API Keys             — Anthropic API key (Claude), HuggingFace
                         token (pyannote for speaker ID).

  Audio Devices        — Microphone and loopback (System Audio)
                         device selection.

  Model Configuration  — Whisper model size, max speakers, Claude
                         model (Haiku 4.5 is default — ~4x cheaper
                         than Sonnet, good enough for summaries).

  Storage              — Where recordings and session files live.

  Email                — Default recipient for the Email button.

  Calendar             — Minutes before a meeting to notify you.

  Workflow             — Auto-process after stop, auto-draft
                         follow-up email, launch on Windows startup.

═══════════════════════════════════════════════════════════
  COST MANAGEMENT
═══════════════════════════════════════════════════════════

Claude Haiku 4.5 costs about $1/M input + $5/M output tokens.
A typical meeting (5–20k input) running all extractions costs
under $0.05. You can run hundreds of meetings for a few dollars.

For deeper meetings (complex design reviews) switch to Sonnet 4.5
in Settings > Model Configuration — about 4x the cost but better
for nuance. Switch back to Haiku for routine meetings.

═══════════════════════════════════════════════════════════
  TROUBLESHOOTING
═══════════════════════════════════════════════════════════

  • "Only captured my voice" — enable Stereo Mix in Windows Sound
    settings (Recording tab > Show Disabled Devices > Enable), then
    set it as System Audio in Settings.

  • Calendar shows no meetings — requires Classic Outlook, not New
    Outlook. If using New Outlook, switch to Classic or meetings
    won't sync.

  • No desktop shortcut after install — run `python make_shortcut.py`
    from the install folder.

  • Session logs per recording are in the recordings folder as
    session_<ID>.log — include one when reporting a bug.

  • Help > Open Logs Folder jumps straight there.
"""


class UsageGuide(tk.Toplevel):

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Usage Guide")
        self.geometry("780x620")
        self.minsize(600, 400)
        self.configure(bg=styles.BG_DARK)
        self.transient(parent)

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - 780) // 2
        py = parent.winfo_y() + (parent.winfo_height() - 620) // 2
        self.geometry(f"+{px}+{py}")

        outer = tk.Frame(self, bg=styles.BG_DARK)
        outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        tk.Label(outer, text="Usage Guide",
                 bg=styles.BG_DARK, fg=styles.TEXT_PRIMARY,
                 font=styles.FONT_HEADER).pack(anchor="w", pady=(0, 8))

        text_wrap = tk.Frame(outer, bg=styles.BG_PANEL,
                              highlightbackground=styles.BORDER,
                              highlightthickness=1)
        text_wrap.pack(fill=tk.BOTH, expand=True)

        text = tk.Text(text_wrap, wrap=tk.WORD, bg=styles.BG_PANEL,
                        fg=styles.TEXT_PRIMARY, font=styles.FONT_BODY,
                        relief=tk.FLAT, padx=16, pady=12)
        vsb = ttk.Scrollbar(text_wrap, orient="vertical",
                             command=text.yview)
        text.configure(yscrollcommand=vsb.set)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        text.insert("1.0", GUIDE_TEXT.strip())
        text.configure(state=tk.DISABLED)

        footer = tk.Frame(outer, bg=styles.BG_DARK)
        footer.pack(fill=tk.X, pady=(10, 0))
        tk.Button(footer, text="Close", bg=styles.BG_PANEL,
                  fg=styles.TEXT_MUTED, font=styles.FONT_BODY,
                  relief=tk.FLAT, padx=16, pady=6, cursor="hand2",
                  command=self.destroy).pack(side=tk.RIGHT)
