<<<<<<< HEAD
# 🎙 Meeting Recorder

An AI-powered desktop meeting recorder for Windows that transcribes, identifies speakers, and summarizes your meetings — all with a single click.

## Features

- 🎤 **Records mic + system audio** simultaneously
- 📝 **Transcribes** using faster-whisper (local, runs on GPU or CPU)
- 👥 **Speaker diarization** — identifies who said what using pyannote
- ✨ **AI Summarization** — generates meeting summaries via Claude API
- 📧 **Email summary** directly to your Outlook inbox
- 📂 **Load existing audio files** for transcription
- 🗓 **Outlook calendar integration** (Windows)
- 🖥 **Material You dark theme UI**

## Requirements

- Windows 10/11
- Python 3.11
- Anthropic API key (for summarization) — [console.anthropic.com](https://console.anthropic.com)
- HuggingFace token (for speaker detection) — [huggingface.co](https://huggingface.co)
- Nvidia GPU recommended (works on CPU too, slower)

## Quick Install

Download `MeetingRecorderSetup.exe` from [Releases](../../releases) and run it.
The installer walks you through everything including API key setup.

## Manual Setup
```bash
git clone https://github.com/YOUR_USERNAME/meeting-recorder
cd meeting-recorder
python -m venv .venv
.venv\Scripts\activate
pip install torch==2.6.0 torchaudio==2.6.0 --index-url https://download.pytorch.org/whl/cu124
pip install faster-whisper pyannote.audio==3.3.2 anthropic sounddevice soundfile scipy matplotlib huggingface_hub==0.23.0 pywin32 python-dotenv pyaudio numpy==2.1.3
```

Copy `.env.example` to `.env` and add your API keys:
```
ANTHROPIC_API_KEY=sk-ant-...
HF_TOKEN=hf_...
WHISPER_MODEL=base
MAX_SPEAKERS=8
RECORDINGS_DIR=recordings
```

Accept model terms at:
- https://huggingface.co/pyannote/speaker-diarization-3.1
- https://huggingface.co/pyannote/segmentation-3.0

Then run:
```bash
python main.py
```

## Mac Port

We welcome Mac contributions! Key areas to port:
- `core/audio_capture.py` — replace loopback with BlackHole driver
- `services/calendar_service.py` — replace win32com with AppleScript
- `ui/app_window.py` — replace Outlook email with AppleScript
- `core/transcription.py` / `core/diarization.py` — swap `cuda` to `mps` for Apple Silicon
- `main.py` — remove Windows-specific torch patches

## Project Structure
```
meeting_recorder/
├── main.py                    # Entry point
├── config/settings.py         # Configuration
├── core/
│   ├── audio_capture.py       # Mic + system audio recording
│   ├── transcription.py       # faster-whisper transcription
│   ├── diarization.py         # pyannote speaker detection
│   └── summarizer.py          # Claude AI summarization
├── services/
│   ├── recording_service.py   # Recording orchestration
│   ├── export_service.py      # File export
│   ├── session_service.py     # Session management
│   └── calendar_service.py    # Outlook calendar integration
├── ui/
│   ├── app_window.py          # Main application window
│   ├── device_panel.py        # Audio device selection
│   ├── speaker_panel.py       # Speaker name editor
│   └── transcript_panel.py    # Transcript display
└── models/
    ├── session.py             # Recording session model
    ├── segment.py             # Transcript segment model
    └── speaker.py             # Speaker model
```

## Contributing

Pull requests welcome! Please open an issue first to discuss major changes.

## License

MIT License — free to use, modify, and distribute.

=======
# meeting-recorder
AI-powered meeting recorder with transcription, speaker detection, and summarization
>>>>>>> 479fadf9352595e03faa6cf3322c3e5b7dc535b4
