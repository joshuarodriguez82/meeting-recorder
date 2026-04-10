# Meeting Recorder

An AI-powered desktop meeting recorder for Windows that transcribes, identifies speakers, and extracts summaries, action items, and requirements — all with a single click.

## Features

- **Records mic + system audio** simultaneously (WASAPI loopback — works with headphones)
- **Unlimited recording duration** — streams to disk, no memory limits
- **Transcribes** using faster-whisper (local, GPU or CPU)
- **Speaker diarization** — identifies who said what using pyannote
- **AI Summarization** — generates meeting summaries via Claude API
- **Action Items extraction** — who, what, by when, plus decisions and open questions
- **Requirements extraction** — structured FR/NFR tables with priority and owner
- **Meeting templates** — General, Requirements Gathering, Design Review, Sprint Planning, Stakeholder Update
- **Email notes** directly via Outlook with structured sections
- **Export** transcripts, summaries, action items, and requirements to text files
- **Settings UI** — change API keys, audio devices, email, and model config without editing files
- **Per-session logs** for debugging
- **Light blue Material Design UI** with modern rounded font

## Requirements

- Windows 10/11
- Python 3.11+
- Anthropic API key (for summarization) — [console.anthropic.com](https://console.anthropic.com)
- HuggingFace token (for speaker detection) — [huggingface.co](https://huggingface.co)
- NVIDIA GPU recommended (works on CPU too, slower)

## Quick Install

Download `MeetingRecorderSetup.exe` from [Releases](../../releases) and run it.
The installer walks you through everything including API key setup.

## System Audio Setup

To capture other meeting participants (not just your mic):

1. Right-click speaker icon > Sound settings > Recording tab
2. Right-click empty space > Show Disabled Devices
3. Enable "Stereo Mix" (right-click > Enable)
4. In Meeting Recorder: File > Settings > set System Audio device

If Stereo Mix isn't available, install [VB-Cable](https://vb-audio.com/Cable/) (free virtual audio cable).

## Manual Setup

```bash
git clone https://github.com/joshuarodriguez82/meeting-recorder
cd meeting-recorder
python -m venv .venv
.venv\Scripts\activate
pip install torch==2.6.0 torchaudio==2.6.0 --index-url https://download.pytorch.org/whl/cu124
pip install faster-whisper pyannote.audio==3.3.2 anthropic sounddevice soundfile scipy matplotlib huggingface_hub==0.23.0 pywin32 python-dotenv pyaudiowpatch numpy==2.1.3
```

Copy `.env.example` to `.env` and add your API keys:
```
ANTHROPIC_API_KEY=sk-ant-...
HF_TOKEN=hf_...
WHISPER_MODEL=base
MAX_SPEAKERS=8
RECORDINGS_DIR=recordings
EMAIL_TO=
```

Accept model terms at:
- https://huggingface.co/pyannote/speaker-diarization-3.1
- https://huggingface.co/pyannote/segmentation-3.0

Then run:
```bash
python main.py
```

## Project Structure

```
meeting_recorder/
├── main.py                    # Entry point
├── config/settings.py         # Configuration (loads from .env)
├── core/
│   ├── audio_capture.py       # Mic + WASAPI loopback recording
│   ├── transcription.py       # faster-whisper transcription
│   ├── diarization.py         # pyannote speaker detection
│   └── summarizer.py          # Claude AI: summaries, action items, requirements
├── services/
│   ├── recording_service.py   # Recording orchestration + session logging
│   ├── export_service.py      # File export (transcript, summary, action items, requirements)
│   └── session_service.py     # Session persistence
├── ui/
│   ├── app_window.py          # Main application window
│   ├── device_panel.py        # Audio device selection
│   ├── settings_dialog.py     # Settings dialog (API keys, devices, email)
│   ├── speaker_panel.py       # Speaker name editor
│   ├── styles.py              # Theme colors and fonts
│   └── transcript_panel.py    # Transcript display
├── models/
│   ├── session.py             # Recording session model
│   ├── segment.py             # Transcript segment model
│   └── speaker.py             # Speaker model
└── installer/
    ├── installer.py           # Installer template
    ├── installer_bundled.py   # Self-contained installer with embedded source
    ├── bundle.py              # Embeds source files into installer
    └── build.bat              # Builds MeetingRecorderSetup.exe
```

## License

MIT License — free to use, modify, and distribute.
