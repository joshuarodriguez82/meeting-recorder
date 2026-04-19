# Meeting Recorder

An AI-powered desktop meeting recorder for Windows that transcribes, identifies speakers, extracts summaries, action items, requirements, and decisions — and organizes everything into a searchable knowledge base built for Solutions Architects.

## Features

### Recording
- **Records mic + system audio** simultaneously via WASAPI loopback (works with headphones)
- **Unlimited duration** — streams to disk, no memory limits
- **Transcribes** locally using faster-whisper (GPU or CPU)
- **Speaker diarization** — identifies who said what (pyannote)
- **Per-session logs** saved alongside each recording for easy debugging

### AI Extraction (Claude)
Each recording can be extracted into structured outputs:

- **Summary** — template-aware (General, Requirements Gathering, Design Review, Sprint Planning, Stakeholder Update)
- **📋 Action Items** — owner, description, due date, plus decisions and open questions
- **📝 Requirements** — FR/NFR tables with priority and owner
- **🎯 Decisions** — auto-generated ADR log (Decided, Rationale, Alternatives, Owner, Impact)

Claude Haiku 4.5 is the default model (~4× cheaper than Sonnet). Switch to Sonnet 4.5 in Settings for complex meetings where quality matters.

### Calendar Integration (Outlook)
- **Today's Meetings** panel pulls from Outlook calendar on startup (requires Classic Outlook)
- **Meeting-start popup** — 2 min before a scheduled meeting, a toast appears with Start Recording / Dismiss buttons
- **Auto-fill meeting name** from calendar subject + date when you click Record
- **Attendee capture** — recipients from the Outlook invite are saved with the session

### Organization & Knowledge Base
- **Session History** — every recording with status icons, double-click to reload, Bulk Process for unprocessed meetings
- **Follow-Up Tracker** — every action item across every meeting, filterable by owner/client/status/search
- **Decision Log** — every decision with full context, searchable across all time
- **Transcript Search** — "what did we say about auth?" across every transcript with context snippets
- **Client Dashboard** — per-client overview: meetings, open action items, decisions, stats
- **Meeting Prep Brief** — Claude generates a pre-meeting brief from all prior meetings tagged with the same client/project
- **Client/Project tags** — auto-complete from previously-used tags, power the dashboard and tracker

### Workflow Automation
- **Auto-process after stop** — optional: runs transcribe → summarize → action items → requirements → decisions automatically
- **Auto-draft follow-up email** — optional: drafts an Outlook email to attendees with summary + action items + decisions after processing
- **Launch on Windows startup** — optional: adds a shortcut to the Startup folder so the app opens on login
- **Retention policy** — optional automatic cleanup of old audio WAV files (transcripts/JSONs are never deleted)

### Ease of Use
- **Settings dialog** for all API keys, audio devices, email, Claude model, calendar, workflow, and retention preferences
- **In-app Usage Guide** (Help > Usage Guide) with full walkthrough
- **Menu bar:** File, Sessions, Help
- **Light blue Material Design UI** with Segoe UI Variable Display font

## Requirements

- Windows 10/11 (Classic Outlook for calendar integration; New Outlook not supported)
- Python 3.11+
- Anthropic API key — [console.anthropic.com](https://console.anthropic.com)
- HuggingFace token — [huggingface.co/settings/tokens](https://huggingface.co/settings/tokens)
- NVIDIA GPU recommended (works on CPU, slower)

## Quick Install (CLI)

For environments where the .exe installer is blocked by security policies:

```powershell
git clone https://github.com/joshuarodriguez82/meeting-recorder.git
cd meeting-recorder
python setup.py
```

Then launch:
```powershell
.venv\Scripts\activate
python main.py
```

Or install a desktop shortcut:
```powershell
python make_shortcut.py
```

The app will prompt you to add your API keys via File > Settings on first run.

### GUI Installer
Download `MeetingRecorderSetup.exe` from [Releases](../../releases). Run it; it walks through API key setup and creates a desktop shortcut automatically.

## System Audio Setup

To capture other meeting participants (not just your mic):

1. Right-click speaker icon > Sound settings > Recording tab
2. Right-click empty space > Show Disabled Devices
3. Enable "Stereo Mix" (right-click > Enable)
4. In Meeting Recorder: File > Settings > set System Audio to your loopback device

If Stereo Mix isn't available, install [VB-Cable](https://vb-audio.com/Cable/) (free).

## HuggingFace Model Access

Accept model terms at these URLs before first use:
- https://huggingface.co/pyannote/speaker-diarization-3.1
- https://huggingface.co/pyannote/segmentation-3.0

## Environment Variables

Settings are stored in `.env`:

```
ANTHROPIC_API_KEY=sk-ant-...
HF_TOKEN=hf_...
WHISPER_MODEL=base
MAX_SPEAKERS=10
RECORDINGS_DIR=recordings
EMAIL_TO=
CLAUDE_MODEL=claude-haiku-4-5
NOTIFY_MINUTES_BEFORE=2
AUTO_PROCESS_AFTER_STOP=false
LAUNCH_ON_STARTUP=false
AUTO_FOLLOW_UP_EMAIL=false
RETENTION_ENABLED=false
RETENTION_PROCESSED_DAYS=7
RETENTION_UNPROCESSED_DAYS=30
```

All editable via File > Settings — you never need to touch `.env` directly.

## Project Structure

```
meeting_recorder/
├── main.py                       # Entry point
├── setup.py                      # CLI installer (venv + deps)
├── make_shortcut.py              # Standalone desktop shortcut creator
├── config/settings.py            # Configuration (.env)
├── core/
│   ├── audio_capture.py          # Mic + WASAPI loopback recording
│   ├── transcription.py          # faster-whisper
│   ├── diarization.py            # pyannote speaker ID
│   └── summarizer.py             # Claude: summary / action items /
│                                 #  requirements / decisions / prep brief
├── services/
│   ├── recording_service.py      # Recording orchestration
│   ├── session_service.py        # Session JSON persistence + list/delete
│   ├── export_service.py         # Text file exports
│   ├── calendar_service.py       # Outlook COM calendar reader
│   ├── calendar_monitor.py       # Background meeting-start notifier
│   └── retention_service.py      # Automatic audio cleanup
├── ui/
│   ├── app_window.py             # Main application window
│   ├── calendar_panel.py         # Today's meetings panel
│   ├── client_dashboard.py       # Per-client overview
│   ├── decision_log.py           # Decision log viewer
│   ├── device_panel.py           # Audio device selection
│   ├── follow_up_tracker.py      # Action items aggregator
│   ├── prep_brief_dialog.py      # Meeting prep brief
│   ├── session_browser.py        # Session history + bulk process
│   ├── settings_dialog.py        # Settings UI
│   ├── speaker_panel.py          # Speaker name editor
│   ├── styles.py                 # Theme colors + fonts
│   ├── transcript_panel.py       # Transcript display
│   ├── transcript_search.py      # Full-text search dialog
│   └── usage_guide.py            # In-app help
├── models/
│   ├── session.py                # Session model
│   ├── segment.py                # Transcript segment
│   └── speaker.py                # Speaker
├── utils/
│   ├── audio_utils.py            # WAV helpers
│   ├── logger.py                 # Logging
│   └── startup_shortcut.py       # Windows startup shortcut manager
└── installer/
    ├── installer.py              # Installer template
    ├── installer_bundled.py      # Self-contained installer (generated)
    ├── bundle.py                 # Embeds source into installer
    └── build.bat                 # Builds MeetingRecorderSetup.exe
```

## License

MIT License — free to use, modify, and distribute.
