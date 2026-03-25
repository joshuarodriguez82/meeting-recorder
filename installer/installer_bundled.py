"""
Meeting Recorder — Bulletproof Windows Installer
Handles all edge cases for mixed GPU, corporate, and consumer setups.
"""

import os
import sys
import json
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser
from pathlib import Path
import ctypes
import winreg
import shutil

APP_FILES    = {
  "main.py": "\"\"\"\nMeeting Recorder — application entry point.\n\"\"\"\n\nimport sys\nimport torch\nimport numpy as np\nfrom torch.torch_version import TorchVersion\n\n# Patch NumPy 2.0 compatibility with pyannote\nif not hasattr(np, 'NaN'):\n    np.NaN = np.nan\nif not hasattr(np, 'NAN'):\n    np.NAN = np.nan\n\n# Patch PyTorch 2.6 compatibility with pyannote\ntorch.serialization.add_safe_globals([TorchVersion])\n_original_torch_load = torch.load\ndef _patched_torch_load(f, *args, **kwargs):\n    kwargs['weights_only'] = False\n    return _original_torch_load(f, *args, **kwargs)\ntorch.load = _patched_torch_load\n\nfrom config.settings import Settings\nfrom ui.app_window import AppWindow\nfrom utils.logger import get_logger\n\nlogger = get_logger(__name__)\n\n\ndef main() -> None:\n    try:\n        settings = Settings.from_env()\n    except EnvironmentError as e:\n        print(f\"\\n[CONFIG ERROR]\\n{e}\\n\")\n        sys.exit(1)\n    try:\n        app.iconbitmap(\"meeting_recorder.ico\")\n    except Exception:\n        pass\n\n    logger.info(\"Starting Meeting Recorder...\")\n    app = AppWindow(settings)\n    app.mainloop()\n\n\nif __name__ == \"__main__\":\n    main()",
  "config/__init__.py": "",
  "config/settings.py": "\"\"\"\nApplication configuration loaded from environment variables.\nAll secrets are sourced from .env — never hardcoded.\n\"\"\"\n\nimport os\nfrom dataclasses import dataclass\nfrom dotenv import load_dotenv\n\nload_dotenv()\n\n\n@dataclass(frozen=True)\nclass Settings:\n    \"\"\"Immutable application settings resolved at startup.\"\"\"\n\n    anthropic_api_key: str\n    hf_token: str\n    whisper_model: str\n    max_speakers: int\n    recordings_dir: str\n\n    @classmethod\n    def from_env(cls) -> \"Settings\":\n        \"\"\"\n        Load settings from environment variables.\n\n        Raises:\n            EnvironmentError: If a required variable is missing.\n        \"\"\"\n        required = {\n            \"ANTHROPIC_API_KEY\": os.getenv(\"ANTHROPIC_API_KEY\"),\n            \"HF_TOKEN\": os.getenv(\"HF_TOKEN\"),\n        }\n\n        missing = [k for k, v in required.items() if not v]\n        if missing:\n            raise EnvironmentError(\n                f\"Missing required environment variables: {', '.join(missing)}\\n\"\n                \"Copy .env.example to .env and fill in your keys.\"\n            )\n\n        return cls(\n            anthropic_api_key=required[\"ANTHROPIC_API_KEY\"],\n            hf_token=required[\"HF_TOKEN\"],\n            whisper_model=os.getenv(\"WHISPER_MODEL\", \"base\"),\n            max_speakers=int(os.getenv(\"MAX_SPEAKERS\", \"10\")),\n            recordings_dir=os.getenv(\"RECORDINGS_DIR\", \"recordings\"),\n        )",
  "core/__init__.py": "",
  "core/audio_capture.py": "import threading\nfrom typing import Callable, List, Optional\nimport numpy as np\nimport sounddevice as sd\nfrom utils.logger import get_logger\n\nlogger = get_logger(__name__)\n\nSAMPLE_RATE = 16000\nBLOCK_SIZE = 1024\n\n\ndef list_input_devices() -> List[dict]:\n    devices = []\n    for idx, dev in enumerate(sd.query_devices()):\n        if dev[\"max_input_channels\"] > 0:\n            devices.append({\n                \"index\": idx,\n                \"name\": dev[\"name\"],\n                \"max_input_channels\": dev[\"max_input_channels\"],\n                \"default_samplerate\": dev[\"default_samplerate\"],\n            })\n    return devices\n\n\ndef list_output_devices() -> List[dict]:\n    devices = []\n    for idx, dev in enumerate(sd.query_devices()):\n        if dev[\"max_output_channels\"] > 0:\n            devices.append({\n                \"index\": idx,\n                \"name\": dev[\"name\"],\n                \"max_output_channels\": dev[\"max_output_channels\"],\n                \"default_samplerate\": dev[\"default_samplerate\"],\n            })\n    return devices\n\n\nclass AudioCapture:\n\n    def __init__(\n        self,\n        mic_device_index: Optional[int],\n        output_device_index: Optional[int],\n        on_chunk: Callable[[np.ndarray], None],\n    ):\n        self._mic_idx = mic_device_index\n        self._out_idx = output_device_index\n        self._on_chunk = on_chunk\n        self._streams: List[sd.InputStream] = []\n        self._lock = threading.Lock()\n        self._mic_buffer: Optional[np.ndarray] = None\n        self._out_buffer: Optional[np.ndarray] = None\n        self._running = False\n        self._chunk_count = 0\n        self.actual_sr = SAMPLE_RATE\n\n    def start(self) -> None:\n        self._running = True\n        self._chunk_count = 0\n        logger.info(f\"Starting capture: mic={self._mic_idx}, output={self._out_idx}\")\n\n        try:\n            if self._mic_idx is not None:\n                dev_info = sd.query_devices(self._mic_idx)\n                native_sr = int(dev_info[\"default_samplerate\"])\n                channels = min(2, int(dev_info[\"max_input_channels\"]))\n                self.actual_sr = native_sr\n                logger.info(f\"Mic device: {dev_info['name']}, sr={native_sr}, ch={channels}\")\n                mic_stream = sd.InputStream(\n                    device=self._mic_idx,\n                    channels=channels,\n                    samplerate=native_sr,\n                    blocksize=BLOCK_SIZE,\n                    dtype=\"float32\",\n                    callback=self._mic_callback,\n                )\n                mic_stream.start()\n                self._streams.append(mic_stream)\n                logger.info(f\"Mic stream started at {native_sr}Hz, {channels}ch\")\n\n            if self._out_idx is not None:\n                try:\n                    dev_info = sd.query_devices(self._out_idx)\n                    native_sr = int(dev_info[\"default_samplerate\"])\n                    channels = min(2, int(dev_info[\"max_output_channels\"]))\n                    logger.info(f\"Output device: {dev_info['name']}, sr={native_sr}, ch={channels}\")\n                    out_stream = sd.InputStream(\n                        device=self._out_idx,\n                        channels=channels,\n                        samplerate=native_sr,\n                        blocksize=BLOCK_SIZE,\n                        dtype=\"float32\",\n                        callback=self._out_callback,\n                    )\n                    out_stream.start()\n                    self._streams.append(out_stream)\n                    logger.info(f\"Output stream started at {native_sr}Hz, {channels}ch\")\n                except Exception as e:\n                    logger.warning(f\"Loopback capture unavailable: {e}. Mic only.\")\n                    self._out_idx = None\n\n        except Exception as e:\n            self._close_all_streams()\n            self._running = False\n            raise\n\n    def stop(self) -> None:\n        self._running = False\n        self._close_all_streams()\n        logger.info(f\"Audio capture stopped. Total chunks captured: {self._chunk_count}\")\n\n    def _close_all_streams(self) -> None:\n        for stream in self._streams:\n            try:\n                stream.stop()\n                stream.close()\n            except Exception as e:\n                logger.warning(f\"Error closing stream: {e}\")\n        self._streams.clear()\n\n    def _mic_callback(self, indata: np.ndarray, frames: int, time, status) -> None:\n        try:\n            if status:\n                logger.warning(f\"Mic stream status: {status}\")\n            chunk = indata.mean(axis=1) if indata.ndim > 1 else indata[:, 0].copy()\n            self._chunk_count += 1\n            if self._chunk_count % 100 == 0:\n                logger.info(f\"Mic chunks received: {self._chunk_count}\")\n            with self._lock:\n                self._mic_buffer = chunk\n                self._try_merge()\n        except Exception as e:\n            logger.error(f\"Error in mic callback: {e}\")\n\n    def _out_callback(self, indata: np.ndarray, frames: int, time, status) -> None:\n        try:\n            if status:\n                logger.warning(f\"Output stream status: {status}\")\n            chunk = indata.mean(axis=1) if indata.ndim > 1 else indata[:, 0].copy()\n            with self._lock:\n                self._out_buffer = chunk\n                self._try_merge()\n        except Exception as e:\n            logger.error(f\"Error in output callback: {e}\")\n\n    def _try_merge(self) -> None:\n        if self._mic_idx is not None and self._out_idx is not None:\n            if self._mic_buffer is not None and self._out_buffer is not None:\n                size = min(len(self._mic_buffer), len(self._out_buffer))\n                merged = (self._mic_buffer[:size] + self._out_buffer[:size]) / 2.0\n                self._mic_buffer = None\n                self._out_buffer = None\n                self._safe_invoke(merged)\n        elif self._mic_buffer is not None:\n            chunk = self._mic_buffer\n            self._mic_buffer = None\n            self._safe_invoke(chunk)\n        elif self._out_buffer is not None:\n            chunk = self._out_buffer\n            self._out_buffer = None\n            self._safe_invoke(chunk)\n\n    def _safe_invoke(self, chunk: np.ndarray) -> None:\n        try:\n            self._on_chunk(chunk)\n        except Exception as e:\n            logger.error(f\"Error in on_chunk callback: {e}\")",
  "core/diarization.py": "\"\"\"\nPyannote speaker diarization — GPU accelerated.\n\"\"\"\n\nimport asyncio\nfrom typing import List, Optional\nfrom utils.logger import get_logger\n\nlogger = get_logger(__name__)\n\n\nclass DiarizationEngine:\n\n    def __init__(self, hf_token: str, max_speakers: int = 8):\n        from pyannote.audio import Pipeline\n        import torch\n        logger.info(\"Loading pyannote diarization pipeline on GPU...\")\n        self._pipeline = Pipeline.from_pretrained(\n            \"pyannote/speaker-diarization-3.1\",\n        )\n        self._pipeline.to(torch.device(\"cuda\"))\n        self._max_speakers = max_speakers\n        logger.info(\"Diarization pipeline loaded on GPU.\")\n\n    async def diarize(self, audio_path: str) -> List[dict]:\n        logger.info(f\"Diarizing: {audio_path}\")\n        loop = asyncio.get_event_loop()\n        try:\n            diarization = await loop.run_in_executor(\n                None,\n                lambda: self._pipeline(\n                    audio_path,\n                    max_speakers=self._max_speakers,\n                )\n            )\n        except Exception as e:\n            raise RuntimeError(\n                f\"Diarization failed: {e}\\n\"\n                \"Check that the audio file is a valid 16kHz mono WAV.\"\n            ) from e\n\n        turns = []\n        for turn, _, speaker in diarization.itertracks(yield_label=True):\n            turns.append({\n                \"start\":   turn.start,\n                \"end\":     turn.end,\n                \"speaker\": speaker,\n            })\n        logger.info(f\"Diarization complete: {len(set(t['speaker'] for t in turns))} speakers detected.\")\n        return turns\n\n    @staticmethod\n    def assign_speakers(\n        segments: List[dict],\n        turns: List[dict],\n    ) -> List[dict]:\n        attributed = []\n        for seg in segments:\n            seg_mid = (seg[\"start\"] + seg[\"end\"]) / 2\n            speaker = \"SPEAKER_UNKNOWN\"\n            best_overlap = 0.0\n            for turn in turns:\n                overlap = min(seg[\"end\"], turn[\"end\"]) - max(seg[\"start\"], turn[\"start\"])\n                if overlap > best_overlap:\n                    best_overlap = overlap\n                    speaker = turn[\"speaker\"]\n            attributed.append({**seg, \"speaker_id\": speaker})\n        return attributed\n",
  "core/summarizer.py": "\"\"\"\nClaude-powered meeting summarizer and speaker identifier.\n\"\"\"\n\nimport asyncio\nimport json\nimport re\nfrom typing import Dict\nfrom anthropic import AsyncAnthropic\nfrom utils.logger import get_logger\n\nlogger = get_logger(__name__)\n\n\ndef _markdown_to_html(text: str) -> str:\n    \"\"\"Convert basic markdown to HTML for email display.\"\"\"\n    lines = text.split(\"\\n\")\n    html_lines = []\n    in_list = False\n\n    for line in lines:\n        # Headers\n        if line.startswith(\"### \"):\n            if in_list:\n                html_lines.append(\"</ul>\")\n                in_list = False\n            html_lines.append(\n                f'<h3 style=\"color:#1a1a1a;font-size:15px;margin:16px 0 6px;\">'\n                f'{line[4:]}</h3>')\n        elif line.startswith(\"## \"):\n            if in_list:\n                html_lines.append(\"</ul>\")\n                in_list = False\n            html_lines.append(\n                f'<h2 style=\"color:#003a57;font-size:17px;margin:20px 0 8px;'\n                f'border-bottom:1px solid #ddd;padding-bottom:4px;\">'\n                f'{line[3:]}</h2>')\n        elif line.startswith(\"# \"):\n            if in_list:\n                html_lines.append(\"</ul>\")\n                in_list = False\n            html_lines.append(\n                f'<h1 style=\"color:#003a57;font-size:20px;margin:20px 0 10px;\">'\n                f'{line[2:]}</h1>')\n        # Bullet points\n        elif line.startswith(\"- \") or line.startswith(\"* \"):\n            if not in_list:\n                html_lines.append(\n                    '<ul style=\"margin:6px 0;padding-left:20px;\">')\n                in_list = True\n            content = _inline_markdown(line[2:])\n            html_lines.append(\n                f'<li style=\"margin:4px 0;color:#333;\">{content}</li>')\n        # Numbered list\n        elif re.match(r\"^\\d+\\. \", line):\n            if not in_list:\n                html_lines.append(\n                    '<ol style=\"margin:6px 0;padding-left:20px;\">')\n                in_list = True\n            content = _inline_markdown(re.sub(r\"^\\d+\\. \", \"\", line))\n            html_lines.append(\n                f'<li style=\"margin:4px 0;color:#333;\">{content}</li>')\n        # Empty line\n        elif line.strip() == \"\":\n            if in_list:\n                html_lines.append(\"</ul>\")\n                in_list = False\n            html_lines.append('<div style=\"height:8px;\"></div>')\n        # Regular paragraph\n        else:\n            if in_list:\n                html_lines.append(\"</ul>\")\n                in_list = False\n            content = _inline_markdown(line)\n            html_lines.append(\n                f'<p style=\"margin:4px 0;color:#333;line-height:1.6;\">'\n                f'{content}</p>')\n\n    if in_list:\n        html_lines.append(\"</ul>\")\n\n    return \"\\n\".join(html_lines)\n\n\ndef _inline_markdown(text: str) -> str:\n    \"\"\"Convert inline markdown (bold, italic, code) to HTML.\"\"\"\n    # Bold\n    text = re.sub(r\"\\*\\*(.+?)\\*\\*\", r'<strong>\\1</strong>', text)\n    text = re.sub(r\"__(.+?)__\",     r'<strong>\\1</strong>', text)\n    # Italic\n    text = re.sub(r\"\\*(.+?)\\*\",     r'<em>\\1</em>', text)\n    text = re.sub(r\"_(.+?)_\",       r'<em>\\1</em>', text)\n    # Inline code\n    text = re.sub(r\"`(.+?)`\",\n                  r'<code style=\"background:#f0f0f0;padding:1px 4px;'\n                  r'border-radius:3px;font-family:monospace;\">\\1</code>',\n                  text)\n    return text\n\n\nclass Summarizer:\n\n    def __init__(self, api_key: str):\n        self._client = AsyncAnthropic(api_key=api_key)\n\n    async def summarize(self, transcript: str) -> str:\n        logger.info(\"Requesting meeting summary from Claude...\")\n        try:\n            message = await asyncio.wait_for(\n                self._client.messages.create(\n                    model=\"claude-sonnet-4-20250514\",\n                    max_tokens=1024,\n                    messages=[{\n                        \"role\": \"user\",\n                        \"content\": (\n                            \"Please summarize this meeting transcript. \"\n                            \"Include: key topics discussed, decisions made, \"\n                            \"action items, and any follow-ups needed.\\n\\n\"\n                            f\"{transcript}\"\n                        )\n                    }]\n                ),\n                timeout=60.0\n            )\n            summary = message.content[0].text\n            logger.info(\"Summary received.\")\n            return summary\n        except Exception as e:\n            raise RuntimeError(f\"Summarization API call failed: {e}\") from e\n\n    async def identify_speakers(self, transcript: str) -> Dict[str, str]:\n        logger.info(\"Requesting speaker identification from Claude...\")\n        try:\n            message = await asyncio.wait_for(\n                self._client.messages.create(\n                    model=\"claude-sonnet-4-20250514\",\n                    max_tokens=512,\n                    messages=[{\n                        \"role\": \"user\",\n                        \"content\": (\n                            \"Analyze this meeting transcript and identify any speakers \"\n                            \"who introduced themselves by name. Return ONLY a JSON object \"\n                            \"mapping speaker IDs to their real names. \"\n                            \"Only include speakers where you are confident of their name \"\n                            \"from an explicit introduction like 'Hi I'm X', 'My name is X', \"\n                            \"'This is X speaking', etc. \"\n                            \"If no introductions are found, return an empty JSON object {}.\\n\\n\"\n                            \"Example response: \"\n                            \"{\\\"SPEAKER_00\\\": \\\"John Smith\\\", \\\"SPEAKER_02\\\": \\\"Sarah Jones\\\"}\\n\\n\"\n                            f\"Transcript:\\n{transcript}\"\n                        )\n                    }]\n                ),\n                timeout=30.0\n            )\n            raw = message.content[0].text.strip()\n            logger.info(f\"Speaker identification response: {raw}\")\n\n            if raw.startswith(\"```\"):\n                lines = raw.split(\"\\n\")\n                raw = \"\\n\".join(\n                    line for line in lines\n                    if not line.startswith(\"```\")\n                ).strip()\n\n            result = json.loads(raw)\n            if not isinstance(result, dict):\n                return {}\n\n            filtered = {\n                k: v for k, v in result.items()\n                if isinstance(k, str) and isinstance(v, str)\n                and k.startswith(\"SPEAKER\") and v.strip()\n            }\n            logger.info(f\"Identified {len(filtered)} speakers by name\")\n            return filtered\n\n        except json.JSONDecodeError:\n            logger.warning(\"Speaker ID response was not valid JSON\")\n            return {}\n        except Exception as e:\n            logger.warning(f\"Speaker identification failed: {e}\")\n            return {}\n\n    def summary_to_html(self, summary: str) -> str:\n        \"\"\"Convert a markdown summary to formatted HTML for email.\"\"\"\n        return _markdown_to_html(summary)",
  "core/transcription.py": "import asyncio\nfrom pathlib import Path\nfrom faster_whisper import WhisperModel\nfrom utils.logger import get_logger\nlogger = get_logger(__name__)\nclass TranscriptionEngine:\n    def __init__(self, model_name=\"base\"):\n        logger.info(f\"Loading faster-whisper model: {model_name}\")\n        self._model = WhisperModel(model_name, device=\"cpu\", compute_type=\"int8\")\n        logger.info(\"faster-whisper model loaded.\")\n    async def transcribe(self, audio_path):\n        if not Path(audio_path).exists():\n            raise FileNotFoundError(f\"Audio file not found: {audio_path}\")\n        loop = asyncio.get_event_loop()\n        try:\n            segments, info = await loop.run_in_executor(None, lambda: self._model.transcribe(audio_path, language=\"en\", vad_filter=True))\n            segment_list = await loop.run_in_executor(None, lambda: [{\"start\": s.start, \"end\": s.end, \"text\": s.text.strip()} for s in segments if s.text.strip()])\n        except Exception as e:\n            raise RuntimeError(f\"Transcription failed: {e}\") from e\n        return segment_list\n",
  "models/__init__.py": "",
  "models/segment.py": "\"\"\"Transcript segment model — a single spoken utterance.\"\"\"\n\nfrom dataclasses import dataclass\n\n\n@dataclass\nclass Segment:\n    \"\"\"A time-bounded spoken segment attributed to one speaker.\"\"\"\n\n    speaker_id: str\n    start: float        # seconds\n    end: float          # seconds\n    text: str\n\n    def to_dict(self) -> dict:\n        return {\n            \"speaker_id\": self.speaker_id,\n            \"start\": round(self.start, 3),\n            \"end\": round(self.end, 3),\n            \"text\": self.text,\n        }\n\n    def formatted(self, display_name: str) -> str:\n        \"\"\"Human-readable line for display and export.\"\"\"\n        start_str = self._format_time(self.start)\n        end_str = self._format_time(self.end)\n        return f\"[{start_str} → {end_str}] {display_name}: {self.text}\"\n\n    @staticmethod\n    def _format_time(seconds: float) -> str:\n        m, s = divmod(int(seconds), 60)\n        return f\"{m:02d}:{s:02d}\"",
  "models/session.py": "from __future__ import annotations\nimport datetime\nfrom typing import Dict, List, Optional\nfrom models.speaker import Speaker\nfrom models.segment import Segment\n\n\nclass Session:\n\n    def __init__(self, session_id: str):\n        self.session_id: str = session_id\n        self.display_name: str = \"\"\n        self.started_at: datetime.datetime = datetime.datetime.now()\n        self.ended_at: Optional[datetime.datetime] = None\n        self.audio_path: Optional[str] = None\n        self.speakers: Dict[str, Speaker] = {}\n        self.segments: List[Segment] = []\n        self.summary: Optional[str] = None\n\n    def get_or_create_speaker(self, speaker_id: str) -> Speaker:\n        if speaker_id not in self.speakers:\n            self.speakers[speaker_id] = Speaker(speaker_id=speaker_id)\n        return self.speakers[speaker_id]\n\n    def rename_speaker(self, speaker_id: str, name: str) -> None:\n        if speaker_id in self.speakers:\n            self.speakers[speaker_id].display_name = name\n\n    def full_transcript(self) -> str:\n        if not self.segments:\n            return \"\"\n        lines = []\n        for seg in self.segments:\n            speaker = self.speakers.get(seg.speaker_id)\n            name = speaker.display_name if speaker else seg.speaker_id\n            start = _fmt_time(seg.start)\n            end = _fmt_time(seg.end)\n            lines.append(f\"[{start} → {end}] {name}: {seg.text}\")\n        return \"\\n\".join(lines)\n\n    def to_dict(self) -> dict:\n        return {\n            \"session_id\": self.session_id,\n            \"display_name\": self.display_name,\n            \"started_at\": self.started_at.isoformat() if self.started_at else None,\n            \"ended_at\": self.ended_at.isoformat() if self.ended_at else None,\n            \"audio_path\": self.audio_path,\n            \"speakers\": {k: v.to_dict() for k, v in self.speakers.items()},\n            \"segments\": [s.to_dict() for s in self.segments],\n            \"summary\": self.summary,\n        }\n\n\ndef _fmt_time(seconds: float) -> str:\n    m, s = divmod(int(seconds), 60)\n    return f\"{m:02d}:{s:02d}\"\n",
  "models/speaker.py": "\"\"\"Speaker identity model.\"\"\"\n\nfrom dataclasses import dataclass, field\nimport uuid\n\n\n@dataclass\nclass Speaker:\n    \"\"\"Represents a detected speaker in a meeting session.\"\"\"\n\n    speaker_id: str = field(default_factory=lambda: f\"SPEAKER_{uuid.uuid4().hex[:4].upper()}\")\n    display_name: str = \"\"\n\n    def __post_init__(self):\n        if not self.display_name:\n            self.display_name = self.speaker_id\n\n    def to_dict(self) -> dict:\n        return {\"speaker_id\": self.speaker_id, \"display_name\": self.display_name}",
  "services/__init__.py": "",
  "services/export_service.py": "\"\"\"\nExports transcripts and summaries to text files.\nUses meeting display name if available for clean filenames.\n\"\"\"\n\nimport os\nfrom pathlib import Path\nfrom models.session import Session\nfrom utils.logger import get_logger\n\nlogger = get_logger(__name__)\n\n\nclass ExportService:\n\n    def __init__(self, recordings_dir: str):\n        self._dir = Path(recordings_dir)\n        self._dir.mkdir(parents=True, exist_ok=True)\n\n    def _base_name(self, session: Session) -> str:\n        if session.display_name:\n            safe = \"\".join(\n                c if c.isalnum() or c in \" -_\" else \"\" \n                for c in session.display_name\n            ).strip()\n            return safe or session.session_id\n        return f\"session_{session.session_id}\"\n\n    def export_transcript(self, session: Session) -> str:\n        name = self._base_name(session)\n        path = self._dir / f\"transcript_{name}.txt\"\n        lines = []\n        if session.display_name:\n            lines.append(f\"Meeting: {session.display_name}\")\n            lines.append(\"=\" * 60)\n            lines.append(\"\")\n        lines.append(session.full_transcript())\n        path.write_text(\"\\n\".join(lines), encoding=\"utf-8\")\n        logger.info(f\"Transcript exported: {path}\")\n        return str(path)\n\n    def export_summary(self, session: Session) -> str:\n        if not session.summary:\n            raise ValueError(\"No summary to export.\")\n        name = self._base_name(session)\n        path = self._dir / f\"summary_{name}.txt\"\n        lines = []\n        if session.display_name:\n            lines.append(f\"Meeting: {session.display_name}\")\n            lines.append(\"=\" * 60)\n            lines.append(\"\")\n        lines.append(session.summary)\n        path.write_text(\"\\n\".join(lines), encoding=\"utf-8\")\n        logger.info(f\"Summary exported: {path}\")\n        return str(path)\n",
  "services/recording_service.py": "\"\"\"\nOrchestrates the full recording lifecycle.\n\"\"\"\n\nimport asyncio\nimport threading\nimport uuid\nfrom datetime import datetime\nfrom pathlib import Path\nfrom typing import Callable, List, Optional\n\nimport numpy as np\nfrom math import gcd\nfrom scipy.signal import resample_poly\n\nfrom config.settings import Settings\nfrom core.audio_capture import AudioCapture\nfrom core.diarization import DiarizationEngine\nfrom core.transcription import TranscriptionEngine\nfrom models.segment import Segment\nfrom models.session import Session\nfrom utils.audio_utils import save_wav\nfrom utils.logger import get_logger\n\nlogger = get_logger(__name__)\n\nTARGET_SR = 16000\nMAX_CHUNKS = int((90 * 60 * TARGET_SR) / 1024)\n\n\ndef _resample(audio: np.ndarray, orig_sr: int) -> np.ndarray:\n    if orig_sr == TARGET_SR:\n        return audio.astype(np.float32)\n    divisor = gcd(TARGET_SR, orig_sr)\n    up = TARGET_SR // divisor\n    down = orig_sr // divisor\n    resampled = resample_poly(audio.astype(np.float64), up, down)\n    return np.clip(resampled, -1.0, 1.0).astype(np.float32)\n\n\nclass RecordingService:\n\n    def __init__(\n        self,\n        settings: Settings,\n        transcription_engine: TranscriptionEngine,\n        diarization_engine: DiarizationEngine,\n        on_status: Optional[Callable[[str], None]] = None,\n    ):\n        self._settings = settings\n        self._transcription = transcription_engine\n        self._diarization = diarization_engine\n        self._on_status = on_status or (lambda _: None)\n        self._session: Optional[Session] = None\n        self._capture: Optional[AudioCapture] = None\n        self._audio_chunks: List[np.ndarray] = []\n        self._chunks_lock = threading.Lock()\n        self._recording = False\n        self._capture_sr = TARGET_SR\n\n    @property\n    def current_session(self) -> Optional[Session]:\n        return self._session\n\n    @property\n    def is_recording(self) -> bool:\n        return self._recording\n\n    def set_session(self, session: Session) -> None:\n        \"\"\"Allow an externally created session (e.g. loaded file) to be processed.\"\"\"\n        self._session = session\n\n    def start_recording(\n        self,\n        mic_device_index: Optional[int],\n        output_device_index: Optional[int],\n    ) -> Session:\n        if self._recording:\n            raise RuntimeError(\"A recording is already in progress.\")\n\n        session_id = uuid.uuid4().hex[:8].upper()\n        self._session = Session(session_id=session_id)\n\n        with self._chunks_lock:\n            self._audio_chunks = []\n\n        self._recording = True\n        self._capture = AudioCapture(\n            mic_device_index=mic_device_index,\n            output_device_index=output_device_index,\n            on_chunk=self._on_audio_chunk,\n        )\n\n        try:\n            self._capture.start()\n            if hasattr(self._capture, 'actual_sr'):\n                self._capture_sr = self._capture.actual_sr\n            else:\n                self._capture_sr = TARGET_SR\n        except Exception as e:\n            self._recording = False\n            self._capture = None\n            raise RuntimeError(f\"Failed to start audio capture: {e}\") from e\n\n        self._on_status(f\"Recording started — Session {session_id}\")\n        logger.info(f\"Session {session_id} recording started.\")\n        return self._session\n\n    def stop_recording(self) -> Optional[Session]:\n        if not self._recording or not self._capture:\n            return self._session\n\n        self._recording = False\n        self._capture.stop()\n\n        capture_sr = self._capture_sr\n        self._capture = None\n\n        with self._chunks_lock:\n            chunks_snapshot = list(self._audio_chunks)\n            self._audio_chunks = []\n\n        if self._session and chunks_snapshot:\n            try:\n                audio = np.concatenate(chunks_snapshot, axis=0)\n                if audio.ndim == 2:\n                    audio = audio.mean(axis=1)\n                audio = _resample(audio, capture_sr)\n                path = self._build_audio_path(self._session.session_id)\n                save_wav(path, audio, TARGET_SR)\n                self._session.audio_path = path\n                self._session.ended_at = datetime.now()\n                self._on_status(\"Recording saved. Ready to process.\")\n                logger.info(f\"Audio saved to {path} ({len(audio)/TARGET_SR:.1f}s)\")\n            except Exception as e:\n                logger.error(f\"Failed to save audio: {e}\")\n                self._on_status(f\"Error saving audio: {e}\")\n        elif self._session and not chunks_snapshot:\n            logger.warning(\"Recording stopped with no audio chunks captured.\")\n            self._on_status(\"No audio was captured. Try again.\")\n\n        return self._session\n\n    async def process_session(self) -> Session:\n        if not self._session or not self._session.audio_path:\n            raise RuntimeError(\"No recorded session to process.\")\n\n        self._on_status(\"__stage:transcribe:active__\")\n        raw_segments = await self._transcription.transcribe(self._session.audio_path)\n\n        if not raw_segments:\n            self._on_status(\"Transcription produced no output. Check audio quality.\")\n            return self._session\n\n        self._on_status(\"__stage:transcribe:done____stage:diarize:active__\")\n        diarization_turns = await self._diarization.diarize(self._session.audio_path)\n\n        self._on_status(\"__stage:diarize:done____stage:speakers:active__\")\n        attributed = DiarizationEngine.assign_speakers(raw_segments, diarization_turns)\n\n        for raw in attributed:\n            speaker = self._session.get_or_create_speaker(raw[\"speaker_id\"])\n            segment = Segment(\n                speaker_id=speaker.speaker_id,\n                start=raw[\"start\"],\n                end=raw[\"end\"],\n                text=raw[\"text\"],\n            )\n            self._session.segments.append(segment)\n\n        self._on_status(\"Processing complete.\")\n        logger.info(f\"Session {self._session.session_id} processing complete.\")\n        return self._session\n\n    def _on_audio_chunk(self, chunk: np.ndarray) -> None:\n        if not self._recording:\n            return\n        with self._chunks_lock:\n            if len(self._audio_chunks) >= MAX_CHUNKS:\n                self._audio_chunks.pop(0)\n            self._audio_chunks.append(chunk)\n\n    def _build_audio_path(self, session_id: str) -> str:\n        recordings_dir = Path(self._settings.recordings_dir)\n        recordings_dir.mkdir(parents=True, exist_ok=True)\n        return str(recordings_dir / f\"session_{session_id}.wav\")\n",
  "services/session_service.py": "\"\"\"\nPersists and loads session data as JSON.\nUses atomic write (temp file + rename) to prevent corrupt JSON on crash.\n\"\"\"\n\nimport json\nimport os\nimport tempfile\nfrom pathlib import Path\nfrom typing import Optional\n\nfrom models.session import Session\nfrom utils.logger import get_logger\n\nlogger = get_logger(__name__)\n\n\nclass SessionService:\n    \"\"\"Handles JSON serialization of Session objects.\"\"\"\n\n    def __init__(self, recordings_dir: str):\n        self._recordings_dir = Path(recordings_dir)\n        self._recordings_dir.mkdir(parents=True, exist_ok=True)\n\n    def save(self, session: Session) -> str:\n        \"\"\"\n        Serialize a session to a JSON file using an atomic write.\n\n        Writes to a temporary file first, then renames it to the final path.\n        This ensures the target file is never left in a half-written state\n        if the process is interrupted mid-write.\n\n        Args:\n            session: The completed Session object.\n\n        Returns:\n            The path of the saved JSON file.\n\n        Raises:\n            OSError: If writing or renaming fails.\n        \"\"\"\n        final_path = self._recordings_dir / f\"session_{session.session_id}.json\"\n        data = json.dumps(session.to_dict(), indent=2, ensure_ascii=False)\n\n        # FIX #10: write to temp file in same directory, then atomic rename\n        try:\n            fd, tmp_path = tempfile.mkstemp(\n                dir=self._recordings_dir,\n                suffix=\".json.tmp\",\n            )\n            try:\n                with os.fdopen(fd, \"w\", encoding=\"utf-8\") as f:\n                    f.write(data)\n                    f.flush()\n                    os.fsync(f.fileno())  # Flush OS buffers before rename\n                os.replace(tmp_path, final_path)  # Atomic on POSIX & Windows\n            except Exception:\n                # Clean up temp file if rename or write fails\n                try:\n                    os.unlink(tmp_path)\n                except OSError:\n                    pass\n                raise\n        except Exception as e:\n            raise OSError(f\"Failed to save session {session.session_id}: {e}\") from e\n\n        logger.info(f\"Session atomically saved: {final_path}\")\n        return str(final_path)\n\n    def load(self, session_id: str) -> Optional[dict]:\n        \"\"\"\n        Load a session JSON by ID.\n\n        Args:\n            session_id: The session identifier.\n\n        Returns:\n            Parsed session dict, or None if not found.\n\n        Raises:\n            ValueError: If the file exists but contains invalid JSON.\n        \"\"\"\n        path = self._recordings_dir / f\"session_{session_id}.json\"\n        if not path.exists():\n            logger.warning(f\"Session file not found: {path}\")\n            return None\n        try:\n            with open(path, \"r\", encoding=\"utf-8\") as f:\n                return json.load(f)\n        except json.JSONDecodeError as e:\n            raise ValueError(f\"Corrupt session file {path}: {e}\") from e\n",
  "ui/__init__.py": "",
  "ui/app_window.py": "\"\"\"\nMain application window — Material You dark theme.\n\"\"\"\n\nimport asyncio\nimport datetime\nimport os\nimport subprocess\nimport threading\nimport tkinter as tk\nfrom tkinter import messagebox, filedialog\nfrom typing import Optional\nimport uuid\n\nfrom config.settings import Settings\nfrom core.diarization import DiarizationEngine\nfrom core.summarizer import Summarizer\nfrom core.transcription import TranscriptionEngine\nfrom models.session import Session\nfrom services.export_service import ExportService\nfrom services.recording_service import RecordingService\nfrom services.session_service import SessionService\nfrom ui import styles\nfrom ui.device_panel import DevicePanel\nfrom ui.speaker_panel import SpeakerPanel\nfrom ui.transcript_panel import TranscriptPanel\nfrom utils.logger import get_logger\n\nlogger = get_logger(__name__)\n\n\ndef _run_async_in_thread(coro_factory, on_success, on_error):\n    def _worker():\n        loop = asyncio.new_event_loop()\n        asyncio.set_event_loop(loop)\n        try:\n            result = loop.run_until_complete(coro_factory())\n            on_success(result)\n        except Exception as e:\n            logger.exception(\"Background async task failed\")\n            on_error(e)\n        finally:\n            loop.close()\n    threading.Thread(target=_worker, daemon=True).start()\n\n\nclass AppWindow(tk.Tk):\n\n    def __init__(self, settings: Settings):\n        super().__init__()\n        self._settings = settings\n        self._session: Optional[Session] = None\n        self._active_meeting: Optional[dict] = None\n        self._progress_after = None\n\n        self._transcription = TranscriptionEngine(settings.whisper_model)\n        self._diarization   = DiarizationEngine(settings.hf_token, settings.max_speakers)\n        self._summarizer    = Summarizer(settings.anthropic_api_key)\n        self._session_svc   = SessionService(settings.recordings_dir)\n        self._export_svc    = ExportService(settings.recordings_dir)\n        self._recording_svc = RecordingService(\n            settings=settings,\n            transcription_engine=self._transcription,\n            diarization_engine=self._diarization,\n            on_status=self._thread_safe_status,\n        )\n\n        self._build_window()\n        self._build_layout()\n\n    def _build_window(self) -> None:\n        self.title(\"Meeting Recorder\")\n        self.geometry(\"920x820\")\n        self.minsize(800, 680)\n        self.configure(bg=styles.BG_DARK)\n        self.protocol(\"WM_DELETE_WINDOW\", self._on_close)\n        try:\n            self.iconbitmap(\"meeting_recorder.ico\")\n        except Exception:\n            pass\n\n    def _build_layout(self) -> None:\n        outer = tk.Frame(self, bg=styles.BG_DARK)\n        outer.pack(fill=tk.BOTH, expand=True, padx=styles.PAD_LG, pady=styles.PAD_LG)\n\n        # ── Top bar ──────────────────────────────────────────────────\n        topbar = tk.Frame(outer, bg=styles.BG_DARK)\n        topbar.pack(fill=tk.X, pady=(0, styles.PAD))\n\n        icon_frame = tk.Frame(topbar, bg=styles.ACCENT_BG, width=42, height=42)\n        icon_frame.pack(side=tk.LEFT)\n        icon_frame.pack_propagate(False)\n        tk.Label(icon_frame, text=\"🎙\", bg=styles.ACCENT_BG,\n                 font=(\"Segoe UI\", 16)).place(relx=0.5, rely=0.5, anchor=\"center\")\n\n        title_block = tk.Frame(topbar, bg=styles.BG_DARK)\n        title_block.pack(side=tk.LEFT, padx=(10, 0))\n        tk.Label(title_block, text=\"Meeting Recorder\", bg=styles.BG_DARK,\n                 fg=styles.TEXT_PRIMARY, font=styles.FONT_HEADER).pack(anchor=\"w\")\n\n        self._status_var = tk.StringVar(value=\"Ready\")\n        tk.Label(topbar, textvariable=self._status_var,\n                 bg=styles.BG_DARK, fg=styles.ACCENT,\n                 font=styles.FONT_SMALL).pack(side=tk.RIGHT, anchor=\"s\")\n\n        # ── Meeting name card ─────────────────────────────────────────\n        name_card = tk.Frame(outer, bg=styles.BG_PANEL,\n                              highlightbackground=styles.BORDER,\n                              highlightthickness=1)\n        name_card.pack(fill=tk.X, pady=(0, styles.PAD))\n        tk.Label(name_card, text=\"MEETING NAME\",\n                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,\n                 font=(\"Segoe UI\", 9)).pack(anchor=\"w\", padx=14, pady=(10, 4))\n\n        name_row = tk.Frame(name_card, bg=styles.BG_PANEL)\n        name_row.pack(fill=tk.X, padx=8, pady=(0, 10))\n\n        self._meeting_name_var = tk.StringVar(\n            value=datetime.datetime.now().strftime(\"%Y-%m-%d Meeting\"))\n        self._name_entry = tk.Entry(\n            name_row,\n            textvariable=self._meeting_name_var,\n            bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,\n            insertbackground=styles.TEXT_PRIMARY,\n            font=styles.FONT_BODY, relief=tk.FLAT,\n            highlightbackground=styles.BORDER, highlightthickness=1,\n        )\n        self._name_entry.pack(fill=tk.X, ipady=8)\n        self._meeting_name_var.trace_add(\"write\", self._on_name_change)\n\n        # ── Device card ───────────────────────────────────────────────\n        device_card = tk.Frame(outer, bg=styles.BG_PANEL,\n                                highlightbackground=styles.BORDER,\n                                highlightthickness=1)\n        device_card.pack(fill=tk.X, pady=(0, styles.PAD))\n        tk.Label(device_card, text=\"AUDIO DEVICES\",\n                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,\n                 font=(\"Segoe UI\", 9)).pack(anchor=\"w\", padx=14, pady=(10, 4))\n        self._device_panel = DevicePanel(device_card)\n        self._device_panel.pack(fill=tk.X, padx=8, pady=(0, 8))\n\n        # ── Progress stages ───────────────────────────────────────────\n        prog_card = tk.Frame(outer, bg=styles.BG_PANEL,\n                              highlightbackground=styles.BORDER,\n                              highlightthickness=1)\n        prog_card.pack(fill=tk.X, pady=(0, styles.PAD))\n\n        prog_inner = tk.Frame(prog_card, bg=styles.BG_PANEL)\n        prog_inner.pack(fill=tk.X, padx=14, pady=10)\n\n        self._stage_labels = {}\n        stages = [\n            (\"transcribe\", \"Transcribe\"),\n            (\"diarize\",    \"Diarize\"),\n            (\"speakers\",   \"ID Speakers\"),\n            (\"complete\",   \"Complete\"),\n        ]\n        for key, label in stages:\n            col = tk.Frame(prog_inner, bg=styles.BG_PANEL)\n            col.pack(side=tk.LEFT, expand=True, fill=tk.X)\n            dot = tk.Label(col, text=\"○\", bg=styles.BG_PANEL,\n                           fg=styles.TEXT_HINT, font=(\"Segoe UI\", 14))\n            dot.pack()\n            lbl = tk.Label(col, text=label, bg=styles.BG_PANEL,\n                           fg=styles.TEXT_HINT, font=styles.FONT_SMALL)\n            lbl.pack()\n            self._stage_labels[key] = (dot, lbl)\n\n        # ── Action row 1 ──────────────────────────────────────────────\n        row1 = tk.Frame(outer, bg=styles.BG_DARK)\n        row1.pack(fill=tk.X, pady=(0, 8))\n\n        self._rec_btn = self._pill_button(\n            row1, \"⏺  Start Recording\", styles.DANGER, self._toggle_recording)\n        self._rec_btn.pack(side=tk.LEFT, padx=(0, 8))\n\n        self._load_btn = self._pill_button(\n            row1, \"📂  Load File\", styles.ACCENT_DIM, self._load_audio_file)\n        self._load_btn.pack(side=tk.LEFT, padx=(0, 8))\n\n        self._process_btn = self._pill_button(\n            row1, \"⚙  Process\", styles.BG_INPUT, self._process, outline=True)\n        self._process_btn.pack(side=tk.LEFT)\n        self._process_btn.config(state=tk.DISABLED)\n\n        # ── Action row 2 ──────────────────────────────────────────────\n        row2 = tk.Frame(outer, bg=styles.BG_DARK)\n        row2.pack(fill=tk.X, pady=(0, styles.PAD))\n\n        self._summarize_btn = self._pill_button(\n            row2, \"✨  Summarize\", styles.BG_INPUT, self._summarize, outline=True)\n        self._summarize_btn.pack(side=tk.LEFT, padx=(0, 8))\n        self._summarize_btn.config(state=tk.DISABLED)\n\n        self._export_btn = self._pill_button(\n            row2, \"💾  Export\", styles.BG_INPUT, self._export, outline=True)\n        self._export_btn.pack(side=tk.LEFT, padx=(0, 8))\n        self._export_btn.config(state=tk.DISABLED)\n\n        self._email_btn = self._pill_button(\n            row2, \"✉  Email Summary\", styles.BG_INPUT, self._email_summary, outline=True)\n        self._email_btn.pack(side=tk.LEFT, padx=(0, 8))\n        self._email_btn.config(state=tk.DISABLED)\n\n        self._folder_btn = self._pill_button(\n            row2, \"📁  Recordings\", styles.BG_INPUT, self._open_recordings, outline=True)\n        self._folder_btn.pack(side=tk.LEFT)\n\n        # ── Speaker card ──────────────────────────────────────────────\n        speaker_card = tk.Frame(outer, bg=styles.BG_PANEL,\n                                 highlightbackground=styles.BORDER,\n                                 highlightthickness=1)\n        speaker_card.pack(fill=tk.X, pady=(0, styles.PAD))\n        tk.Label(speaker_card, text=\"SPEAKERS\",\n                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,\n                 font=(\"Segoe UI\", 9)).pack(anchor=\"w\", padx=14, pady=(10, 4))\n        self._speaker_panel = SpeakerPanel(speaker_card, on_rename=self._on_rename_speaker)\n        self._speaker_panel.pack(fill=tk.X, padx=8, pady=(0, 8))\n\n        # ── Transcript card ───────────────────────────────────────────\n        transcript_card = tk.Frame(outer, bg=styles.BG_PANEL,\n                                    highlightbackground=styles.BORDER,\n                                    highlightthickness=1)\n        transcript_card.pack(fill=tk.BOTH, expand=True)\n        tk.Label(transcript_card, text=\"TRANSCRIPT\",\n                 bg=styles.BG_PANEL, fg=styles.TEXT_HINT,\n                 font=(\"Segoe UI\", 9)).pack(anchor=\"w\", padx=14, pady=(10, 4))\n        self._transcript_panel = TranscriptPanel(transcript_card)\n        self._transcript_panel.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))\n\n    # ------------------------------------------------------------------ #\n    # Meeting name\n    # ------------------------------------------------------------------ #\n\n    def _on_name_change(self, *args) -> None:\n        if self._session:\n            self._session.display_name = self._meeting_name_var.get().strip()\n\n    def _get_meeting_name(self) -> str:\n        name = self._meeting_name_var.get().strip()\n        if not name:\n            name = datetime.datetime.now().strftime(\"%Y-%m-%d Meeting\")\n        return name\n\n    # ------------------------------------------------------------------ #\n    # Progress stages\n    # ------------------------------------------------------------------ #\n\n    def _set_stage(self, stage: str, state: str) -> None:\n        dot, lbl = self._stage_labels[stage]\n        if state == \"active\":\n            dot.config(text=\"●\", fg=styles.ACCENT)\n            lbl.config(fg=styles.ACCENT)\n            self._animate_dot(stage)\n        elif state == \"done\":\n            dot.config(text=\"✓\", fg=styles.SUCCESS_DIM)\n            lbl.config(fg=styles.SUCCESS_DIM)\n        else:\n            dot.config(text=\"○\", fg=styles.TEXT_HINT)\n            lbl.config(fg=styles.TEXT_HINT)\n\n    def _animate_dot(self, stage: str) -> None:\n        frames = [\"●\", \"○\", \"●\", \"○\"]\n        def _tick(i=0):\n            if stage not in self._stage_labels:\n                return\n            dot, _ = self._stage_labels[stage]\n            if dot.cget(\"text\") == \"✓\":\n                return\n            if dot.cget(\"fg\") not in (styles.ACCENT, styles.TEXT_HINT):\n                return\n            dot.config(text=frames[i % len(frames)])\n            self._progress_after = self.after(500, lambda: _tick(i + 1))\n        _tick()\n\n    def _reset_stages(self) -> None:\n        for key in self._stage_labels:\n            self._set_stage(key, \"pending\")\n\n    # ------------------------------------------------------------------ #\n    # Recording\n    # ------------------------------------------------------------------ #\n\n    def _toggle_recording(self) -> None:\n        if not self._recording_svc.is_recording:\n            self._active_meeting = None\n            mic_idx = self._device_panel.get_mic_index()\n            out_idx = self._device_panel.get_output_index()\n            try:\n                self._session = self._recording_svc.start_recording(mic_idx, out_idx)\n                self._session.display_name = self._get_meeting_name()\n            except Exception as e:\n                messagebox.showerror(\"Recording Error\", str(e))\n                return\n            self._reset_stages()\n            self._transcript_panel.clear()\n            self._rec_btn.config(text=\"⏹  Stop Recording\",\n                                  bg=styles.BG_INPUT, fg=styles.DANGER_DIM)\n            self._load_btn.config(state=tk.DISABLED)\n            self._process_btn.config(state=tk.DISABLED)\n            self._summarize_btn.config(state=tk.DISABLED)\n            self._export_btn.config(state=tk.DISABLED)\n            self._email_btn.config(state=tk.DISABLED)\n            self._set_status(\"Recording...\")\n        else:\n            self._session = self._recording_svc.stop_recording()\n            self._rec_btn.config(text=\"⏺  Start Recording\",\n                                  bg=styles.DANGER, fg=styles.TEXT_PRIMARY)\n            self._load_btn.config(state=tk.NORMAL)\n            if self._session and self._session.audio_path:\n                self._process_btn.config(state=tk.NORMAL,\n                                          bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY)\n            self._set_status(\"Recording saved. Ready to process.\")\n\n    # ------------------------------------------------------------------ #\n    # Load file\n    # ------------------------------------------------------------------ #\n\n    def _load_audio_file(self) -> None:\n        file_path = filedialog.askopenfilename(\n            title=\"Select Audio File\",\n            filetypes=[\n                (\"Audio Files\", \"*.wav *.mp3 *.m4a *.flac *.ogg *.aac\"),\n                (\"All files\", \"*.*\"),\n            ]\n        )\n        if not file_path:\n            return\n        session_id = uuid.uuid4().hex[:8].upper()\n        self._session = Session(session_id=session_id)\n        self._session.started_at = datetime.datetime.now()\n        self._session.ended_at   = datetime.datetime.now()\n        self._session.audio_path = file_path\n        self._session.display_name = self._get_meeting_name()\n        self._active_meeting = None\n        self._recording_svc.set_session(self._session)\n        self._reset_stages()\n        self._transcript_panel.clear()\n        self._summarize_btn.config(state=tk.DISABLED)\n        self._export_btn.config(state=tk.DISABLED)\n        self._email_btn.config(state=tk.DISABLED)\n        self._process_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,\n                                  fg=styles.TEXT_PRIMARY)\n        self._set_status(f\"Loaded: {os.path.basename(file_path)}\")\n\n    # ------------------------------------------------------------------ #\n    # Processing\n    # ------------------------------------------------------------------ #\n\n    def _process(self) -> None:\n        if not self._session or not self._session.audio_path:\n            messagebox.showwarning(\"No Audio\", \"Please record or load an audio file first.\")\n            return\n        self._process_btn.config(state=tk.DISABLED, bg=styles.BG_INPUT,\n                                  fg=styles.TEXT_MUTED)\n        self._reset_stages()\n        self._set_stage(\"transcribe\", \"active\")\n        self._set_status(\"Transcribing audio...\")\n        recording_svc = self._recording_svc\n\n        def _coro():\n            return recording_svc.process_session()\n\n        def _on_success(result):\n            self.after(0, lambda: self._on_process_complete(result))\n\n        def _on_error(e):\n            self.after(0, lambda: messagebox.showerror(\"Processing Error\", str(e)))\n            self.after(0, lambda: self._set_status(\"Processing failed.\"))\n            self.after(0, lambda: self._reset_stages())\n            self.after(0, lambda: self._process_btn.config(\n                state=tk.NORMAL, bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY))\n\n        _run_async_in_thread(_coro, _on_success, _on_error)\n\n    def _on_process_complete(self, session: Session) -> None:\n        self._session = session\n        self._session.display_name = self._get_meeting_name()\n        self._set_stage(\"transcribe\", \"done\")\n        self._set_stage(\"diarize\", \"done\")\n        self._set_stage(\"speakers\", \"active\")\n        self._set_status(\"Identifying speakers...\")\n        self._speaker_panel.populate(session)\n        self._transcript_panel.set_text(session.full_transcript())\n        self._summarize_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,\n                                    fg=styles.TEXT_PRIMARY)\n        self._export_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,\n                                 fg=styles.TEXT_PRIMARY)\n        try:\n            self._session_svc.save(session)\n            self._export_svc.export_transcript(session)\n        except Exception as e:\n            logger.error(f\"Failed to save session: {e}\")\n        self._auto_identify_speakers(session)\n\n    def _auto_identify_speakers(self, session: Session) -> None:\n        transcript = session.full_transcript()\n        summarizer = self._summarizer\n\n        def _coro():\n            return summarizer.identify_speakers(transcript)\n\n        def _on_success(suggestions: dict):\n            self.after(0, lambda: self._show_speaker_suggestions(suggestions, session))\n\n        def _on_error(e):\n            logger.warning(f\"Speaker identification failed: {e}\")\n            self.after(0, lambda: self._set_stage(\"speakers\", \"done\"))\n            self.after(0, lambda: self._set_stage(\"complete\", \"done\"))\n            self.after(0, lambda: self._set_status(\"Processing complete.\"))\n\n        _run_async_in_thread(_coro, _on_success, _on_error)\n\n    def _show_speaker_suggestions(self, suggestions: dict, session: Session) -> None:\n        self._set_stage(\"speakers\", \"done\")\n        self._set_stage(\"complete\", \"done\")\n        self._set_status(\"Processing complete.\")\n\n        if not suggestions:\n            return\n\n        overlay = tk.Toplevel(self)\n        overlay.title(\"Speaker Names Detected\")\n        overlay.configure(bg=styles.BG_PANEL)\n        overlay.resizable(False, False)\n        overlay.transient(self)\n\n        overlay.update_idletasks()\n        x = self.winfo_x() + (self.winfo_width()  // 2) - 200\n        y = self.winfo_y() + (self.winfo_height() // 2) - 150\n        overlay.geometry(f\"400x{80 + len(suggestions) * 36 + 80}+{x}+{y}\")\n\n        tk.Label(overlay, text=\"Names detected from introductions\",\n                 bg=styles.BG_PANEL, fg=styles.TEXT_PRIMARY,\n                 font=(\"Segoe UI\", 12, \"bold\")).pack(pady=(16, 4))\n        tk.Label(overlay,\n                 text=\"Auto-applying in 5 seconds — click Apply to confirm now\",\n                 bg=styles.BG_PANEL, fg=styles.TEXT_MUTED,\n                 font=styles.FONT_SMALL).pack(pady=(0, 12))\n\n        for speaker_id, name in suggestions.items():\n            row = tk.Frame(overlay, bg=styles.BG_INPUT, pady=6, padx=12)\n            row.pack(fill=tk.X, padx=16, pady=2)\n            tk.Label(row, text=speaker_id, bg=styles.ACCENT_BG,\n                     fg=styles.ACCENT, font=styles.FONT_SMALL,\n                     padx=8, pady=2).pack(side=tk.LEFT)\n            tk.Label(row, text=\"→\", bg=styles.BG_INPUT,\n                     fg=styles.TEXT_MUTED, font=styles.FONT_BODY).pack(side=tk.LEFT, padx=8)\n            tk.Label(row, text=name, bg=styles.BG_INPUT,\n                     fg=styles.TEXT_PRIMARY, font=styles.FONT_BODY).pack(side=tk.LEFT)\n\n        countdown_var = tk.StringVar(value=\"Apply (5)\")\n        apply_btn = tk.Button(\n            overlay, textvariable=countdown_var,\n            bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY,\n            font=styles.FONT_BODY, relief=tk.FLAT, padx=20, pady=8,\n            cursor=\"hand2\",\n        )\n        apply_btn.pack(pady=16)\n\n        applied = [False]\n\n        def _apply():\n            if applied[0]:\n                return\n            applied[0] = True\n            for speaker_id, name in suggestions.items():\n                session.rename_speaker(speaker_id, name)\n            self._speaker_panel.populate(session)\n            self._transcript_panel.set_text(session.full_transcript())\n            try:\n                self._session_svc.save(session)\n                self._export_svc.export_transcript(session)\n            except Exception as e:\n                logger.error(f\"Failed to save after speaker rename: {e}\")\n            overlay.destroy()\n\n        apply_btn.config(command=_apply)\n\n        def _countdown(n=5):\n            if applied[0]:\n                return\n            if n <= 0:\n                _apply()\n                return\n            countdown_var.set(f\"Apply ({n})\")\n            overlay.after(1000, lambda: _countdown(n - 1))\n\n        _countdown()\n\n        tk.Button(overlay, text=\"Skip\", bg=styles.BG_PANEL,\n                  fg=styles.TEXT_MUTED, font=styles.FONT_SMALL,\n                  relief=tk.FLAT, cursor=\"hand2\",\n                  command=overlay.destroy).pack()\n\n    # ------------------------------------------------------------------ #\n    # Summarize\n    # ------------------------------------------------------------------ #\n\n    def _summarize(self) -> None:\n        if not self._session or not self._session.segments:\n            messagebox.showwarning(\"No Transcript\", \"Please process a recording first.\")\n            return\n        transcript_snapshot = self._session.full_transcript()\n        session_ref = self._session\n        self._summarize_btn.config(state=tk.DISABLED, bg=styles.BG_INPUT,\n                                    fg=styles.TEXT_MUTED)\n        self._set_status(\"Generating AI summary...\")\n        summarizer = self._summarizer\n\n        def _coro():\n            return summarizer.summarize(transcript_snapshot)\n\n        def _on_success(summary):\n            def _apply():\n                session_ref.summary = summary\n                self._transcript_panel.set_text(\n                    session_ref.full_transcript() + \"\\n\\n── SUMMARY ──\\n\\n\" + summary)\n                self._set_status(\"Summary complete.\")\n                self._summarize_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,\n                                            fg=styles.TEXT_PRIMARY)\n                self._email_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,\n                                        fg=styles.TEXT_PRIMARY)\n                try:\n                    self._export_svc.export_summary(session_ref)\n                    self._session_svc.save(session_ref)\n                except Exception as ex:\n                    logger.error(f\"Failed to auto-save summary: {ex}\")\n            self.after(0, _apply)\n\n        def _on_error(e):\n            self.after(0, lambda: messagebox.showerror(\"Summarization Error\", str(e)))\n            self.after(0, lambda: self._set_status(\"Summarization failed.\"))\n            self.after(0, lambda: self._summarize_btn.config(\n                state=tk.NORMAL, bg=styles.ACCENT_DIM, fg=styles.TEXT_PRIMARY))\n\n        _run_async_in_thread(_coro, _on_success, _on_error)\n\n    # ------------------------------------------------------------------ #\n    # Email summary\n    # ------------------------------------------------------------------ #\n\n    def _email_summary(self) -> None:\n        if not self._session or not self._session.summary:\n            messagebox.showwarning(\"No Summary\", \"Please summarize the recording first.\")\n            return\n\n        title    = self._get_meeting_name()\n        date_str = datetime.datetime.now().strftime(\"%B %d, %Y\").replace(\" 0\", \" \")\n        t_path   = None\n        try:\n            t_path = self._export_svc.export_transcript(self._session)\n        except Exception:\n            pass\n\n        self._email_btn.config(state=tk.DISABLED, fg=styles.TEXT_MUTED)\n        self._set_status(\"Sending email...\")\n        session_summary = self._session.summary\n\n        def _send():\n            try:\n                import pythoncom\n                import win32com.client\n                import time\n\n                pythoncom.CoInitialize()\n                outlook = None\n                for attempt in range(4):\n                    try:\n                        outlook = win32com.client.GetActiveObject(\"Outlook.Application\")\n                        break\n                    except Exception:\n                        try:\n                            outlook = win32com.client.Dispatch(\"Outlook.Application\")\n                            break\n                        except Exception:\n                            time.sleep(2)\n\n                if outlook is None:\n                    raise RuntimeError(\n                        \"Could not connect to Outlook after 4 attempts.\\n\"\n                        \"Make sure Outlook is fully open (not just in the tray).\")\n\n                ns   = outlook.GetNamespace(\"MAPI\")\n                mail = outlook.CreateItem(0)\n\n                transcript_line = \"\"\n                if t_path:\n                    transcript_line = (\n                        f\"<p style='font-size:12px;color:#888;margin-top:16px;'>\"\n                        f\"Transcript saved to: {t_path}</p>\"\n                    )\n\n                formatted_summary = self._summarizer.summary_to_html(session_summary)\n                html = f\"\"\"\n<html><body style=\"font-family:Segoe UI,sans-serif;color:#1a1a1a;max-width:680px;\">\n<div style=\"background:#003a57;padding:20px 24px;border-radius:8px;margin-bottom:20px;\">\n  <h2 style=\"color:#4fc3f7;margin:0;font-size:18px;\">Meeting Recorder Summary</h2>\n  <p style=\"color:#90caf9;margin:4px 0 0;font-size:13px;\">{title} &mdash; {date_str}</p>\n</div>\n<div style=\"background:#f5f5f5;padding:20px 24px;border-radius:8px;\n            font-size:14px;line-height:1.7;color:#222;\">\n{formatted_summary}\n</div>\n{transcript_line}\n<p style=\"font-size:11px;color:#aaa;margin-top:24px;\n          border-top:1px solid #eee;padding-top:12px;\">\n  Sent automatically by Meeting Recorder\n</p>\n</body></html>\n\"\"\"\n                mail.Subject    = f\"Meeting Notes: {title} ({date_str})\"\n                mail.BodyFormat = 2\n                mail.HTMLBody   = html\n                mail.To         = ns.CurrentUser.Address\n                mail.Send()\n                pythoncom.CoUninitialize()\n\n                def _done():\n                    self._set_status(\"Summary emailed to you.\")\n                    messagebox.showinfo(\n                        \"Email Sent\",\n                        \"The meeting summary has been sent to your Outlook inbox.\")\n                    self._email_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,\n                                            fg=styles.TEXT_PRIMARY)\n                self.after(0, _done)\n\n            except Exception as e:\n                def _fail(err=e):\n                    self._set_status(\"Email failed.\")\n                    messagebox.showerror(\n                        \"Email Failed\",\n                        f\"Could not send via Outlook:\\n{err}\\n\\n\"\n                        \"Make sure Outlook is fully open and try again.\")\n                    self._email_btn.config(state=tk.NORMAL, bg=styles.ACCENT_DIM,\n                                            fg=styles.TEXT_PRIMARY)\n                self.after(0, _fail)\n\n        threading.Thread(target=_send, daemon=True).start()\n\n    # ------------------------------------------------------------------ #\n    # Export & utilities\n    # ------------------------------------------------------------------ #\n\n    def _export(self) -> None:\n        if not self._session:\n            return\n        try:\n            self._session.display_name = self._get_meeting_name()\n            t_path = self._export_svc.export_transcript(self._session)\n            s_path = (self._export_svc.export_summary(self._session)\n                      if self._session.summary else None)\n            msg = f\"Transcript saved:\\n{t_path}\"\n            if s_path:\n                msg += f\"\\n\\nSummary saved:\\n{s_path}\"\n            messagebox.showinfo(\"Export Complete\", msg)\n        except Exception as e:\n            messagebox.showerror(\"Export Failed\", str(e))\n\n    def _open_recordings(self) -> None:\n        folder = os.path.abspath(self._settings.recordings_dir)\n        os.makedirs(folder, exist_ok=True)\n        subprocess.Popen(f'explorer \"{folder}\"')\n\n    def _on_rename_speaker(self, speaker_id: str, name: str) -> None:\n        if self._session:\n            self._session.rename_speaker(speaker_id, name)\n            self._transcript_panel.set_text(self._session.full_transcript())\n\n    def _set_status(self, message: str) -> None:\n        self._status_var.set(message)\n\n    def _thread_safe_status(self, message: str) -> None:\n        if \"__stage:\" in message:\n            parts = message.split(\"__stage:\")\n            for part in parts:\n                if not part:\n                    continue\n                part = part.rstrip(\"_\")\n                bits = part.split(\":\")\n                if len(bits) == 2:\n                    stage, state = bits\n                    self.after(0, lambda s=stage, st=state: self._set_stage(s, st))\n        else:\n            self.after(0, lambda: self._set_status(message))\n\n    def _pill_button(self, parent, text, color, command, outline=False) -> tk.Button:\n        if outline:\n            return tk.Button(\n                parent, text=text, bg=styles.BG_DARK, fg=styles.TEXT_MUTED,\n                font=styles.FONT_BODY, relief=tk.FLAT, padx=16, pady=8,\n                cursor=\"hand2\", activebackground=styles.BG_INPUT,\n                activeforeground=styles.TEXT_PRIMARY, command=command,\n                highlightbackground=styles.BORDER, highlightthickness=1,\n            )\n        return tk.Button(\n            parent, text=text, bg=color, fg=styles.TEXT_PRIMARY,\n            font=styles.FONT_BODY, relief=tk.FLAT, padx=16, pady=8,\n            cursor=\"hand2\", activebackground=color,\n            activeforeground=styles.TEXT_PRIMARY, command=command,\n        )\n\n    def _on_close(self) -> None:\n        if self._recording_svc.is_recording:\n            self._recording_svc.stop_recording()\n        self.destroy()\n",
  "ui/device_panel.py": "import json\nimport os\nimport tkinter as tk\nfrom tkinter import ttk\nfrom typing import Optional\nfrom core.audio_capture import list_input_devices, list_output_devices\nfrom ui import styles\n\nPREFS_FILE = \"device_prefs.json\"\n\n\ndef _load_prefs() -> dict:\n    try:\n        if os.path.exists(PREFS_FILE):\n            with open(PREFS_FILE, \"r\") as f:\n                return json.load(f)\n    except Exception:\n        pass\n    return {}\n\n\ndef _save_prefs(mic_index, out_index) -> None:\n    try:\n        with open(PREFS_FILE, \"w\") as f:\n            json.dump({\"mic_index\": mic_index, \"out_index\": out_index}, f)\n    except Exception:\n        pass\n\n\nclass DevicePanel(tk.Frame):\n\n    def __init__(self, parent, **kwargs):\n        super().__init__(parent, bg=styles.BG_PANEL, **kwargs)\n        self._input_devices = list_input_devices()\n        self._output_devices = list_output_devices()\n        self._prefs = _load_prefs()\n        self._build()\n\n    def _build(self):\n        style = ttk.Style()\n        style.theme_use(\"default\")\n        style.configure(\n            \"M.TCombobox\",\n            fieldbackground=styles.BG_INPUT,\n            background=styles.BG_INPUT,\n            foreground=styles.TEXT_PRIMARY,\n            selectbackground=styles.ACCENT_BG,\n            selectforeground=styles.ACCENT,\n            borderwidth=0,\n            relief=\"flat\",\n            padding=(10, 8),\n        )\n        style.map(\"M.TCombobox\",\n                  fieldbackground=[(\"readonly\", styles.BG_INPUT)],\n                  foreground=[(\"readonly\", styles.TEXT_PRIMARY)])\n\n        mic_row = tk.Frame(self, bg=styles.BG_PANEL)\n        mic_row.pack(fill=tk.X, pady=(0, 8))\n        tk.Label(mic_row, text=\"Microphone\", bg=styles.BG_PANEL,\n                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL, width=12,\n                 anchor=\"w\").pack(side=tk.LEFT, padx=(2, 8))\n\n        self._mic_var = tk.StringVar()\n        mic_names = [f\"[{d['index']}] {d['name']}\" for d in self._input_devices]\n        self._mic_combo = ttk.Combobox(mic_row, textvariable=self._mic_var,\n                                        values=mic_names, state=\"readonly\",\n                                        style=\"M.TCombobox\")\n        self._mic_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)\n        if mic_names:\n            self._mic_combo.current(0)\n            saved = self._prefs.get(\"mic_index\")\n            if saved is not None:\n                for i, d in enumerate(self._input_devices):\n                    if d[\"index\"] == saved:\n                        self._mic_combo.current(i)\n                        break\n        self._mic_combo.bind(\"<<ComboboxSelected>>\", lambda e: self._on_change())\n\n        out_row = tk.Frame(self, bg=styles.BG_PANEL)\n        out_row.pack(fill=tk.X)\n        tk.Label(out_row, text=\"System Audio\", bg=styles.BG_PANEL,\n                 fg=styles.TEXT_MUTED, font=styles.FONT_SMALL, width=12,\n                 anchor=\"w\").pack(side=tk.LEFT, padx=(2, 8))\n\n        self._out_var = tk.StringVar()\n        out_names = [\"[None] — Skip\"] + [\n            f\"[{d['index']}] {d['name']}\" for d in self._output_devices]\n        self._out_combo = ttk.Combobox(out_row, textvariable=self._out_var,\n                                        values=out_names, state=\"readonly\",\n                                        style=\"M.TCombobox\")\n        self._out_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)\n        self._out_combo.current(0)\n        saved_out = self._prefs.get(\"out_index\")\n        if saved_out is not None:\n            for i, d in enumerate(self._output_devices):\n                if d[\"index\"] == saved_out:\n                    self._out_combo.current(i + 1)\n                    break\n        self._out_combo.bind(\"<<ComboboxSelected>>\", lambda e: self._on_change())\n\n    def _on_change(self):\n        _save_prefs(self.get_mic_index(), self.get_output_index())\n\n    def get_mic_index(self) -> Optional[int]:\n        val = self._mic_var.get()\n        if not val:\n            return None\n        return int(val.split(\"]\")[0].replace(\"[\", \"\").strip())\n\n    def get_output_index(self) -> Optional[int]:\n        val = self._out_var.get()\n        if not val or val.startswith(\"[None]\"):\n            return None\n        return int(val.split(\"]\")[0].replace(\"[\", \"\").strip())\n",
  "ui/speaker_panel.py": "\"\"\"\nSpeaker naming panel — appears after diarization.\nDisplays detected speakers and allows renaming them.\n\"\"\"\n\nimport tkinter as tk\nfrom typing import Callable, Dict\n\nfrom models.session import Session\nfrom ui import styles\n\n\nclass SpeakerPanel(tk.LabelFrame):\n    \"\"\"Widget for assigning human names to detected speaker IDs.\"\"\"\n\n    def __init__(self, parent, on_rename: Callable[[str, str], None], **kwargs):\n        super().__init__(\n            parent,\n            text=\" Speaker Names \",\n            bg=styles.BG_PANEL,\n            fg=styles.TEXT_PRIMARY,\n            font=styles.FONT_BODY,\n            bd=1,\n            relief=tk.SOLID,\n            **kwargs,\n        )\n        self._on_rename = on_rename\n        self._entries: Dict[str, tk.Entry] = {}\n        self._placeholder = tk.Label(\n            self,\n            text=\"Speakers will appear here after processing.\",\n            bg=styles.BG_PANEL,\n            fg=styles.TEXT_MUTED,\n            font=styles.FONT_SMALL,\n        )\n        self._placeholder.pack(pady=styles.PAD)\n\n    def populate(self, session: Session) -> None:\n        \"\"\"Render one row per detected speaker with a rename entry.\"\"\"\n        for widget in self.winfo_children():\n            widget.destroy()\n        self._entries.clear()\n\n        if not session.speakers:\n            tk.Label(self, text=\"No speakers detected.\", bg=styles.BG_PANEL,\n                     fg=styles.TEXT_MUTED, font=styles.FONT_SMALL).pack(pady=styles.PAD)\n            return\n\n        for row, (speaker_id, speaker) in enumerate(session.speakers.items()):\n            frame = tk.Frame(self, bg=styles.BG_PANEL)\n            frame.pack(fill=tk.X, padx=styles.PAD, pady=3)\n\n            tk.Label(frame, text=f\"{speaker_id}:\", bg=styles.BG_PANEL,\n                     fg=styles.TEXT_MUTED, font=styles.FONT_SMALL, width=18, anchor=\"w\").pack(\n                side=tk.LEFT\n            )\n\n            entry = tk.Entry(frame, bg=styles.BG_INPUT, fg=styles.TEXT_PRIMARY,\n                             insertbackground=styles.TEXT_PRIMARY,\n                             font=styles.FONT_BODY, relief=tk.FLAT)\n            entry.insert(0, speaker.display_name)\n            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))\n            self._entries[speaker_id] = entry\n\n            entry.bind(\"<FocusOut>\", lambda e, sid=speaker_id: self._commit(sid))\n            entry.bind(\"<Return>\", lambda e, sid=speaker_id: self._commit(sid))\n\n    def _commit(self, speaker_id: str) -> None:\n        entry = self._entries.get(speaker_id)\n        if entry:\n            name = entry.get().strip()\n            if name:\n                self._on_rename(speaker_id, name)",
  "ui/styles.py": "# Material You — Dark theme with Blue accent\n\nBG_DARK = \"#1a1c1e\"\nBG_PANEL = \"#2d2f31\"\nBG_INPUT = \"#3b3d3f\"\nBG_CARD = \"#2d2f31\"\n\nACCENT = \"#4fc3f7\"\nACCENT_DIM = \"#0277bd\"\nACCENT_BG = \"#003a57\"\n\nTEXT_PRIMARY = \"#e6e1e5\"\nTEXT_MUTED = \"#938f99\"\nTEXT_HINT = \"#6b6870\"\n\nDANGER = \"#b5242a\"\nDANGER_DIM = \"#ff8a80\"\nSUCCESS = \"#2e7d32\"\nSUCCESS_DIM = \"#a5d6a7\"\nWARNING = \"#e65100\"\nWARNING_DIM = \"#ffcc80\"\n\nBORDER = \"#4a4d50\"\nBORDER_SUBTLE = \"#3b3d3f\"\n\nRADIUS = 16\nRADIUS_SM = 12\nRADIUS_PILL = 20\n\nFONT_HEADER = (\"Segoe UI\", 18, \"bold\")\nFONT_TITLE = (\"Segoe UI\", 13)\nFONT_BODY = (\"Segoe UI\", 11)\nFONT_SMALL = (\"Segoe UI\", 10)\nFONT_MONO = (\"Cascadia Code\", 10)\n\nPAD = 12\nPAD_LG = 16",
  "ui/transcript_panel.py": "import tkinter as tk\nfrom ui import styles\n\n\nclass TranscriptPanel(tk.Frame):\n\n    def __init__(self, parent, **kwargs):\n        super().__init__(parent, bg=styles.BG_PANEL, **kwargs)\n        self._build()\n\n    def _build(self):\n        scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL,\n                                  bg=styles.BG_INPUT,\n                                  troughcolor=styles.BG_PANEL,\n                                  relief=tk.FLAT)\n        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)\n        self._text = tk.Text(\n            self,\n            bg=styles.BG_PANEL,\n            fg=styles.TEXT_PRIMARY,\n            font=styles.FONT_MONO,\n            relief=tk.FLAT,\n            wrap=tk.WORD,\n            state=tk.DISABLED,\n            yscrollcommand=scrollbar.set,\n            padx=6,\n            pady=6,\n            selectbackground=styles.ACCENT_BG,\n            selectforeground=styles.ACCENT,\n            insertbackground=styles.ACCENT,\n            spacing1=4,\n            spacing3=4,\n        )\n        self._text.pack(fill=tk.BOTH, expand=True)\n        scrollbar.config(command=self._text.yview)\n\n        self._text.tag_configure(\n            \"speaker\", foreground=styles.ACCENT,\n            font=(\"Segoe UI\", 10, \"bold\"))\n        self._text.tag_configure(\n            \"timestamp\", foreground=styles.TEXT_HINT,\n            font=(\"Segoe UI\", 9))\n        self._text.tag_configure(\n            \"divider\", foreground=styles.ACCENT_DIM,\n            font=(\"Segoe UI\", 10, \"bold\"))\n\n    def set_text(self, content: str) -> None:\n        self._text.config(state=tk.NORMAL)\n        self._text.delete(\"1.0\", tk.END)\n        for line in content.splitlines():\n            if line.startswith(\"──\"):\n                self._text.insert(tk.END, line + \"\\n\", \"divider\")\n            elif \"]\" in line and \"→\" in line:\n                parts = line.split(\"] \", 1)\n                if len(parts) == 2:\n                    ts_part = parts[0] + \"] \"\n                    rest = parts[1]\n                    if \": \" in rest:\n                        speaker, text = rest.split(\": \", 1)\n                        self._text.insert(tk.END, ts_part, \"timestamp\")\n                        self._text.insert(tk.END, speaker + \": \", \"speaker\")\n                        self._text.insert(tk.END, text + \"\\n\")\n                    else:\n                        self._text.insert(tk.END, line + \"\\n\")\n                else:\n                    self._text.insert(tk.END, line + \"\\n\")\n            else:\n                self._text.insert(tk.END, line + \"\\n\")\n        self._text.config(state=tk.DISABLED)\n        self._text.see(tk.END)\n\n    def append_line(self, line: str) -> None:\n        self._text.config(state=tk.NORMAL)\n        self._text.insert(tk.END, line + \"\\n\")\n        self._text.config(state=tk.DISABLED)\n        self._text.see(tk.END)\n\n    def clear(self) -> None:\n        self._text.config(state=tk.NORMAL)\n        self._text.delete(\"1.0\", tk.END)\n        self._text.config(state=tk.DISABLED)\n",
  "utils/__init__.py": "",
  "utils/audio_utils.py": "\"\"\"Audio file helpers: WAV writing, resampling, mixing.\"\"\"\n\nimport numpy as np\nimport soundfile as sf\nfrom scipy.signal import resample_poly\nfrom math import gcd\nfrom pathlib import Path\n\nfrom utils.logger import get_logger\n\nlogger = get_logger(__name__)\n\nTARGET_SAMPLE_RATE = 16000\n\n\ndef save_wav(path: str, audio: np.ndarray, samplerate: int) -> None:\n    Path(path).parent.mkdir(parents=True, exist_ok=True)\n    audio_clipped = np.clip(audio, -1.0, 1.0)\n    sf.write(path, audio_clipped, samplerate, subtype=\"PCM_16\")\n    logger.info(f\"Saved WAV: {path} ({len(audio_clipped)/samplerate:.1f}s @ {samplerate}Hz)\")\n\n\ndef resample_to_16k(audio: np.ndarray, orig_sr: int) -> np.ndarray:\n    \"\"\"\n    High quality resample to 16kHz using polyphase filtering.\n    Much better than scipy.signal.resample for audio — no aliasing artifacts.\n    \"\"\"\n    if orig_sr == TARGET_SAMPLE_RATE:\n        return audio.astype(np.float32)\n    divisor = gcd(TARGET_SAMPLE_RATE, orig_sr)\n    up = TARGET_SAMPLE_RATE // divisor\n    down = orig_sr // divisor\n    resampled = resample_poly(audio, up, down)\n    return np.clip(resampled, -1.0, 1.0).astype(np.float32)\n\n\ndef mix_stereo_to_mono(audio: np.ndarray) -> np.ndarray:\n    if audio.ndim == 2:\n        return audio.mean(axis=1).astype(np.float32)\n    return audio.astype(np.float32)",
  "utils/logger.py": "\"\"\"Structured application logger.\"\"\"\n\nimport logging\nimport sys\n\n\ndef get_logger(name: str) -> logging.Logger:\n    \"\"\"\n    Return a configured logger for the given module name.\n\n    Args:\n        name: Typically __name__ from the calling module.\n\n    Returns:\n        Configured Logger instance.\n    \"\"\"\n    logger = logging.getLogger(name)\n    if not logger.handlers:\n        handler = logging.StreamHandler(sys.stdout)\n        handler.setFormatter(\n            logging.Formatter(\"%(asctime)s [%(levelname)s] %(name)s: %(message)s\")\n        )\n        logger.addHandler(handler)\n        logger.setLevel(logging.INFO)\n    return logger"
}
APP_ICON_B64 = "AAABAAYAEBAAAAAAIADhAAAAZgAAACAgAAAAACAAWgEAAEcBAAAwMAAAAAAgALcBAAChAgAAQEAAAAAAIAAyAgAAWAQAAICAAAAAACAALQQAAIoGAAAAAAAAAAAgAJ4IAAC3CgAAiVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAqElEQVR4nGNkgAIpGbn/DCSAZ08eMTIwMDAwkaMZWQ8jOZqRARM+SdPlNxlMl98kzwBkjfgMwesCYgCGAabLb24lpAlZDUogomn2xqF/KwMDA8PmLccYJJYmesNd8CJ6PorNpyPVsbJhmmF6qOcCqE0wTVg1n45Uh6uRWJrozcBAREKCRSG6N2AApwHY4h6bIRSnA0YGBvIyEwMDJEcywRjkaGZgYGAAAEBSRn2KlynOAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAABIUlEQVR4nO2XSw6CMBCGp42HMMa45QRsvAJ6AMLKk7EyHEC5ghtOwJYQwi101aQdpg8KtJr4JywY2vl+pg8oA0KH4+lNxZdq6DuGY0pgK7DJCA8NxyweGo5NcFvDrcVivL2snW/HtGqV+yZPvPJ4DQGG62KbGDCBfExEn4R/A99tIK3aeg2IKY/WgOi01IQtD7kRjUWpNBbLy3Wz0S3HsSjr/f2WybFJBTAcy2TC9OzxfJH5vSYhBXKNYU0M4BLJkkvb5IlyUW1kXS9nMj9ZAdzIBUAZtOUFsHyO06qtmzzJpHstXBaG4zzOBjSmZsFtWvxDMneJYkXfilcfAoB51fi9CqwtTh2XQmnoOxZ9CLhwEhosmBwHQsIB0OlYKOTx/AMkZYQDQu2k0QAAAABJRU5ErkJggolQTkcNChoKAAAADUlIRFIAAAAwAAAAMAgGAAAAVwL5hwAAAX5JREFUeJztWUuOwjAMdaI5BBohtj1BT8EJqlnNyVihnmBO0RN0ixCaWwwbMgrBjpPiKDH0SagoH/s924nSxgCDz+3ujxtTEpfzycT6yc7axENQQizW2Bp5AJqTSRnUGvxsoBnQhH8BWqIPcM/Vhg1a4DirLyGjMfo+PkoZ7sf5oW0aOnE/RUoIIx9rfwbiAjiS0iJEBaSSkxShfhdaBdTGKqA2VgG1kSWgH+efUkSW+kkW4IyWFpHr5/VL6PfrgEaiH2eRI4GzQ9mi/DtEBbjJnBGH1ONy6rgU/+IlxJGj+pe+KxRZA9IkY4gK2By/9/4TA1W709A9/FLn5vhnM0BN9gktWcz+nFhmYuQBBEsoR0SVF5pp6Pb+8/Y/m1g4JrSB+YlB7LNKblSlFrRYCeUQktyNin/YcpkpsYUCvMNZqHWoF1BtFwKQWRdrBmrDcvewLeNyPhn1JWQB+NvwFuE427BBA17qnhiNeqs7E1YlaAZaLCeKE0u0dja4YF4BSwmo22SJCaMAAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAAAQAAAAEAIBgAAAKppcd4AAAH5SURBVHic7VtLbsMgFMRWD1FVVbc+gU+RE0Rd9WRdVT5BT+ETeBtFUW/RrqgowfAenzc1MJsoDsEz8waMcTIoBp6eX7457VG4XS8DtW2w4VFE7yFkxuj78OjilQprcLpTg3AXXGm4S0Ct4pVyaxtDDWqDrdE7B7SAXwNaqL6GqXW0D7QCrbkPATQBNIYW42/iAXXiedn+vF/PE4SHeAJs4TakjRCdA0LiqW1yQswAjjBJE5q/CogYEFNRqRT0BKAJoNENQBNAoxuAJoBGNwBNAI1oA+Zl+8xJJAUpXKIM0Cf8DyakculDAE0ADbIBX6/v3ojNy7Z7AxOzybH3HX0e381SiKsJkgG6Q07HKHC5ig0BTgoobXNtnYnOAZLCqMhuAGXT0yVy7zin7xhk2xZfzxOLYEqlc6aElIDHj7eT+RpCiUpR++RyJQ8BSodmZXKaYPZFqT5VvFKFJ8EcJpTeHM1ugD2ZpQiwK1/iChFlwHqeTuZrCKGVW0p7LhcbxZ8N+oRQrhyl1wViD0e5Q0FqQQT/fYA2BvV4vN8Oowmg0Q1AE0CjeQOg6wAqSl4hegLQ6wA0mk/AyPmDUW24XS9DTwCaABqjUrz/2dUCrXm0D7QAU2sfAuabFlJga7xLQM0muLR5xdaySvQV1TsH1JCGkAaWwKMkglO4H9vT0yoMu6OSAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAD9ElEQVR4nO2dzY3bMBCFx0aKCIIgV1fgKlKBkVMqy2nhClyFK/DVMIx0kRwCbWStfqkZambe+04LLERRfB+HlOzV7sSYL1+//bE+R2aej/vOsn31xhm4LdpCqDXG4OuiJcLqRhj8tqwVofhgBu+LUhH2JQcxfH+UZrLIGgYfgyXVYHYFYPhxWJLVLAEYfjzmZjYpAMOPy5zsRgVg+PGZynBQAIafh7Esi24DSR56BeDsz8dQph8EYPh56cuWSwA4LwJw9uenmzErADjvAnD249DOmhUAnL0IZz8iTeasAOBQAHAoADg7tPX/eL6N/v56OlTqiQ9gBJgKvguKCOkFWBp8l+wipN4DrA1fqw3PpBVAM7jMEqQUwCKwrBKkE8AyqIwSpBOALCOVADVmaLYqkEoAspw0AtScmZmqQBoBSBkUABwKAA4FAIcCgEMBwKEA4FAAcCgAOBQAHAoADgUAhwKAQwHAoQDgUABwKAA4FACc6gIcz7fL8Xy71D6vd7Yal6oCtC+QEvxny3GpJkDfhVGC7ceFewBwKAA4FAAcCgBOGgFqvsmj5Fxe/5rok3aDv3/8uoiIfH77+b3k+O5AZXhFS3NNx/Ot+HrWjusQqhWg6WT351rUkGXpObpCl1QCy3FVE6CvY1tI4J2lAlmPa5o9QINlFciwHHVJJ4CITVAZwxdJKoCIbmBZwxdJLICITnCZwxdxKEB3wDXe9FkSYulxbSLc0qo/B/BKM/h8WfQrLgW4ng4vQa15gNLXdg0izH4RxSWg7wmV9lMrRKzHVXUP0O4Yw9fDclzVlwCtDlouA9ZYlH+rCeXuLmAMr5+otYnQxzauBeibOZ4HuK9v3quWawFE4kgQMXyRAAKI+JcgavgiQQQQ8StB5PBFKv/TqO733a+nw+Kd7VDotQddsx8a41JK1QrQvrDSixwaYA9vCy+VUGNcSgn7b+PGAreqBluc05qwAjTMmfml4Vi27YXwAoiUl/+5nxBOHR+ZFAI01NoHZAi+IZUADVYiZAq+IaUAXZBL/BQQAgwR5UsbloR5EkhsoADgUABwKAA4FAAcCgAOBQCHAoBDAcAJ/yTQw9fCIj9BZAUAhwKAQwHACb8HIOtgBQCHAoCzfz7uu607Qbbh+bjvWAHAoQDgUABw9iL/1oKtO0Lq0mTOCgDOuwCsAji0s2YFAOdFAFaB/HQzZgUA54MArAJ56cu2twJQgnwMZcolAJxBAVgF8jCW5WgFoATxmcpwcgmgBHGZk92sPQAliMfczGZvAilBHJZkVRQqv0jqk5JJWnQbyGrgj9JMVgfJarAtayej2kymCHXRqsLqpZwi2KK9/Jqv5RRiHdb7rb9+wJf+7zt3ygAAAABJRU5ErkJggolQTkcNChoKAAAADUlIRFIAAAEAAAABAAgGAAAAXHKoZgAACGVJREFUeJzt3MttHEkWBdAqoY0YCIK2tIBWjAXErMayWTVoQVtBC7glCGK80CwG1aJK9clPfF7EOwfojVrdjMqMe/NlVrGOh4l8/fb9R+81ML+P97dj7zWUMuQLEXQiGrEYhlmw0DOSUcog9CKFnhlELoNwCxN6ZhatDMIsRvDJJEoRdF+E4JNZ7yLo9sMFH37qVQTNf6jgw3Wti+BLyx8m/HBb64w0aRvBh/VaTAPVJwDhh21aZKdqAQg/7FM7Q1VGDMGH8mrcEhSfAIQf6qiRraIFIPxQV+mMFSsA4Yc2SmatSAEIP7RVKnO7C0D4oY8S2dtVAMIPfe3N4OYCEH6IYU8WNxWA8EMsWzPZ9JeBgFhWF4CrP8S0JZurCkD4Iba1GV1cAMIPY1iTVc8AILFFBeDqD2NZmtm7BSD8MKYl2XULAIndLABXfxjbvQybACCxqwXg6g9zuJVlEwAkdrEAXP1hLtcybQKAxH4rAFd/mNOlbJsAIDEFAIn9UgDGf5jbecZNAJCYAoDE/i4A4z/k8DnrJgBITAFAYl8OB+M/ZHPKvAkAElMAkNgfvRdAO4/Pr4v+3svTQ+WVEMXxcPAMYFZLA3+PQpjTx/vbUQFMqFTwzymCuXy8vx2Pwj+PWsE/pwjmoQAm0Cr45xTB+LwLMLhe4e/9sylDAQwsQgAjrIHtFMCgIgUv0lpYRwEMKGLgIq6J+xTAYCIHLfLauEwBDGSEgI2wRn5SAIMYKVgjrTU7BTCAEQM14pozUgCQmAIIbuQr6chrz0IBQGIKILAZrqAzvIaZKQBITAEENdOVc6bXMhsFAIkpAEhMAUBiCiCgGe+ZZ3xNM1AAkJgCgMQUACSmACAxBQCJKQBITAFAYgoAElMAkJgCgMQUACSmACAxBQCJKQBITAFAYgoAElMAkJgCgMQUACSmACAxBQCJKQBITAFAYgoAElMAkJgCgMQUACSmACAxBQCJKQBITAFAYgoAElMAkJgCgMQUACT2R+8F1Pb4/PrX+Z+9PD38s8daiC3jXpl6Arh0Qm/9OXll3SvTFsC9Ezf7iWW5zHtlygJYesJmPrEsk32vTFkAwDIKABJTAJCYAoDEFAAkpgAgMQUAiSmAgF6eHnovobgZX9MMFAAkpgAgMQUAiSmAoGa6Z27xWh6fX6v/jBmF/z6A//7rP7/9EsY//vx39d/RXrKhZgrpDM7PWevz02uv7hF6Arh0QG/9+WxmKJieV//H59df/qlp1L0atgDuHbjoB5Y8Rt6rIQtg6QGLfGBLGXkKiLT2WmsZfa+GLACgDQUwgEhX0qVGXHNGCmAQIwVqpLVmpwAGMkKwRlgjPymAwUQOWOS1cZkC2KHXp88iBq3XmnwCcB8FcEXEkH0WaX2R1nJJ9PX1pAAGFmFjR1gD2ymAnXqPoD0D2Dv8vY/9DML/MhD3nYLYKhC9g085JoAblm70KFeil6eHquGs/f9fY+kxj7LeqBTAAqNtotJBjRT8pUZbby9uARZYcrV5fH4Nt+nO1zPLVXPp+eA+BXDHy9PD3c205O9EED3Yayw9L9zmFqCAEcI/G8e8DAWwwGgPA2c2y21MFCELYOn3qEX/vjXmN/peDVkAh8P9A9b6gJoC+ot69Y+2V9cIWwCHw/UD1/OAGi3j6nluIu7VJcK/CxDtAI76luDoRnjrL9peXSL0BBDNmlD33owzWXMsFe86CqAiJbCfY1iXAljJFSYu52Y9BbCRdwXqi/rUfyYKYINRPvqbhYeu2ymAjTwQrMuDvzbCvw0Y2ZpJ4PPfs2Ev21KUjuU+JoCdtmxAE8HvhL8PBVCAEthH+PtRAIWcNuTaZwOZi2Dt699yjLnNM4CCtr47cPpvsmzsraXnaX95CqCw1t/Qm4Xg1zHlLcDL08OiX8pY+vdaOZXGjOUR9bWNuldKOX799v1H70XU8vj8+te1f9fihJbY7KNf+UY5Br33Si9TF8DhcPnEtjyhpa54oxXBiK+7917pYfoCiOD08GqUq+EepV6jB35tKIBGSt/7RgvH7K9vVgqgsRoPwXqFZabXkpUC6KTkbcG5WiGqtVbjfj8KIIDWb41dC1uUddCODwIlFO29ePoxAQSSJZiu/HEogIBmLQLBj0cBBDZLEQh+XApgAKMWgeDHpwAGFa0UhH1MCmAS3sJjCwUwKR/NZQkFkMySYhD2PKb8QhBgGQUAiSkASEwBQGIKABJTAJCYAoDEFAAkpgAgMQUAiSkASEwBQGIKABJTAJCYAoDEFAAkpgAgMQUAiflKsM6ifbtvD76CrB8TACSmACAxBQCJKQBITAFAYgoAElMAkJgCgMR8EAgSMwFAYgoAEvvy8f527L0IoL2P97ejCQASUwCQmAKAxBQAJKYAILEvh8P/nwb2XgjQzinzJgBITAFAYn8XgNsAyOFz1k0AkJgCgMR+KQC3ATC384ybACAxBQCJ/VYAbgNgTpeybQKAxC4WgCkA5nIt0yYASOxqAZgCYA63smwCgMRuFoApAMZ2L8MmAEjsbgGYAmBMS7K7aAJQAjCWpZl1CwCJLS4AUwCMYU1WV00ASgBiW5vR1bcASgBi2pJNzwAgsU0FYAqAWLZmcvMEoAQghj1Z3HULoASgr70Z3P0MQAlAHyWyV+QhoBKAtkplrti7AEoA2iiZtaJvAyoBqKt0xop/DkAJQB01slU1rF+/ff9R8/8PGdS8qFb9JKBpAPapnaHqHwVWArBNi+w0DadbAriv5UWz6S8DmQbgttYZ6RZI0wD81Ovi2P2KrAjIrPdU3L0AThQBmfQO/kmIRXymCJhZlOCfhFrMOWXADKKF/rOwCzunDBhJ5NB/NsQizykDIhol9J8Nt+BbFAMtjBj0a/4HKCoNk9MNG4oAAAAASUVORK5CYII="

APP_NAME    = "Meeting Recorder"
APP_VERSION = "1.0.0"
PY_URL      = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
MIN_DISK_GB = 5.0

BG          = "#1a1c1e"
PANEL       = "#2d2f31"
INPUT       = "#3b3d3f"
ACCENT      = "#4fc3f7"
ACCENT2     = "#0277bd"
ACCENT_BG   = "#003a57"
TEXT        = "#e6e1e5"
MUTED       = "#938f99"
HINT        = "#6b6870"
BORDER      = "#4a4d50"
WARN_BG     = "#3a2a00"
WARN_FG     = "#ffcc80"
OK_BG       = "#1a3a1a"
OK_FG       = "#a5d6a7"
SUCCESS_DIM = "#a5d6a7"

FONT_TITLE = ("Segoe UI", 19, "bold")
FONT_BODY  = ("Segoe UI", 11)
FONT_SMALL = ("Segoe UI", 10)
FONT_MONO  = ("Cascadia Code", 10)

REQUIREMENTS_BASE = [
    "python-dotenv",
    "sounddevice",
    "soundfile",
    "scipy",
    "faster-whisper",
    "matplotlib",
    "huggingface_hub==0.23.0",
    "pyannote.audio==3.3.2",
    "pywin32",
    "anthropic",
]


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def get_free_disk_gb(path):
    try:
        total, used, free = shutil.disk_usage(path)
        return free / (1024 ** 3)
    except Exception:
        return 999.0


def has_nvidia_gpu():
    try:
        r = subprocess.run(["nvidia-smi"], capture_output=True, timeout=5)
        return r.returncode == 0
    except Exception:
        return False


def find_python311():
    candidates = ["python3.11", "python", "python3", "py"]
    for cmd in candidates:
        try:
            r = subprocess.run([cmd, "--version"], capture_output=True, text=True, timeout=5)
            if r.returncode == 0:
                ver = r.stdout.strip().split()[1]
                major, minor = int(ver.split(".")[0]), int(ver.split(".")[1])
                if major == 3 and minor == 11:
                    return cmd, ver
        except Exception:
            continue
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Python\PythonCore\3.11\InstallPath") as k:
            path = winreg.QueryValue(k, None)
            exe = Path(path) / "python.exe"
            if exe.exists():
                return str(exe), "3.11"
    except Exception:
        pass
    return None, None


def test_anthropic_key(key):
    try:
        import urllib.request
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/models",
            headers={"x-api-key": key, "anthropic-version": "2023-06-01"})
        with urllib.request.urlopen(req, timeout=10):
            return True, "API key is valid"
    except Exception as e:
        err = str(e)
        if "401" in err:
            return False, "Invalid API key — check you copied it correctly"
        if "403" in err:
            return False, "API key has no credits — add billing at console.anthropic.com"
        if "URLError" in err or "timeout" in err.lower():
            return False, "Cannot reach Anthropic — check internet connection"
        return False, f"Validation failed: {err[:100]}"


def test_hf_token(token):
    try:
        import urllib.request
        req = urllib.request.Request(
            "https://huggingface.co/api/whoami",
            headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read())
            name = data.get("name", "unknown")
            return True, f"Token valid — logged in as {name}"
    except Exception as e:
        err = str(e)
        if "401" in err:
            return False, "Invalid token — check you copied it correctly"
        return False, f"Validation failed: {err[:100]}"


def check_pyannote_access(token):
    import urllib.request
    for model_id in ["pyannote/speaker-diarization-3.1", "pyannote/segmentation-3.0"]:
        try:
            req = urllib.request.Request(
                f"https://huggingface.co/api/models/{model_id}",
                headers={"Authorization": f"Bearer {token}"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                pass
        except Exception as e:
            if "403" in str(e):
                return False, f"Terms not accepted for {model_id}"
    return True, "Model access confirmed"


def create_vbs(install_dir):
    vbs = (f'Set WshShell = CreateObject("WScript.Shell")\n'
           f'WshShell.Run "cmd /c cd {install_dir} && '
           f'call .venv\\Scripts\\activate && python main.py", 0, False\n')
    path = Path(install_dir) / "launch.vbs"
    path.write_text(vbs)
    return str(path)


def create_shortcut(install_dir, vbs_path):
    desktop = Path(os.path.expanduser("~")) / "Desktop"
    lnk = desktop / f"{APP_NAME}.lnk"
    ico = Path(install_dir) / "meeting_recorder.ico"
    ps = (f'$ws = New-Object -ComObject WScript.Shell\n'
          f'$s = $ws.CreateShortcut("{lnk}")\n'
          f'$s.TargetPath = "{vbs_path}"\n'
          f'$s.WorkingDirectory = "{install_dir}"\n'
          f'$s.Description = "{APP_NAME}"\n'
          f'$s.WindowStyle = 1\n')
    if ico.exists():
        ps += f'$s.IconLocation = "{ico}"\n'
    ps += "$s.Save()"
    subprocess.run(["powershell", "-Command", ps], capture_output=True)


class Installer(tk.Tk):

    STEPS = ["Welcome", "System Check", "Location", "Python 3.11",
             "Anthropic Key", "HuggingFace", "Model Terms", "Installing", "Done"]

    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} Installer v{APP_VERSION}")
        self.geometry("740x580")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._install_dir   = tk.StringVar(value=str(Path.home() / "meeting_recorder"))
        self._anthropic_key = tk.StringVar()
        self._hf_token      = tk.StringVar()
        self._python_cmd    = None
        self._python_ver    = None
        self._has_nvidia    = False
        self._step          = 0

        self._build_chrome()
        self._show_step(0)

    def _center(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - 740) // 2
        y = (self.winfo_screenheight() - 580) // 2
        self.geometry(f"740x580+{x}+{y}")

    def _build_chrome(self):
        sidebar = tk.Frame(self, bg=PANEL, width=190)
        sidebar.pack(side=tk.LEFT, fill=tk.Y)
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text="🎙", bg=PANEL, fg=ACCENT,
                 font=("Segoe UI", 26)).pack(pady=(24, 4))
        tk.Label(sidebar, text=APP_NAME, bg=PANEL, fg=TEXT,
                 font=("Segoe UI", 10, "bold"), wraplength=170).pack()
        tk.Label(sidebar, text=f"v{APP_VERSION}", bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 9)).pack(pady=(2, 20))
        self._step_labels = []
        for name in self.STEPS:
            frm = tk.Frame(sidebar, bg=PANEL)
            frm.pack(fill=tk.X, padx=10, pady=2)
            dot = tk.Label(frm, text="●", bg=PANEL, fg=HINT, font=("Segoe UI", 8))
            dot.pack(side=tk.LEFT, padx=(0, 6))
            lbl = tk.Label(frm, text=name, bg=PANEL, fg=HINT,
                           font=("Segoe UI", 9), anchor="w")
            lbl.pack(side=tk.LEFT)
            self._step_labels.append((dot, lbl))

        self._content = tk.Frame(self, bg=BG)
        self._content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._bottom = tk.Frame(self._content, bg=PANEL, height=58)
        self._bottom.pack(side=tk.BOTTOM, fill=tk.X)
        self._bottom.pack_propagate(False)
        self._back_btn = tk.Button(
            self._bottom, text="← Back", bg=PANEL, fg=MUTED,
            font=FONT_BODY, relief=tk.FLAT, padx=14, pady=8,
            cursor="hand2", command=self._back)
        self._back_btn.pack(side=tk.LEFT, padx=14, pady=10)
        self._next_btn = tk.Button(
            self._bottom, text="Next →", bg=ACCENT2, fg=TEXT,
            font=FONT_BODY, relief=tk.FLAT, padx=18, pady=8,
            cursor="hand2", command=self._next)
        self._next_btn.pack(side=tk.RIGHT, padx=14, pady=10)
        self._page = tk.Frame(self._content, bg=BG)
        self._page.pack(fill=tk.BOTH, expand=True, padx=24, pady=18)

    def _update_sidebar(self, idx):
        for i, (dot, lbl) in enumerate(self._step_labels):
            if i < idx:
                dot.config(fg=ACCENT2); lbl.config(fg=MUTED)
            elif i == idx:
                dot.config(fg=ACCENT); lbl.config(fg=TEXT)
            else:
                dot.config(fg=HINT); lbl.config(fg=HINT)

    def _clear_page(self):
        for w in self._page.winfo_children():
            w.destroy()

    def _show_step(self, step):
        self._step = step
        self._update_sidebar(step)
        self._clear_page()
        [self._pg_welcome, self._pg_syscheck, self._pg_location,
         self._pg_python, self._pg_anthropic, self._pg_huggingface,
         self._pg_pyannote, self._pg_install, self._pg_done][step]()

    def _next(self):
        s = self._step
        if s == 0:
            self._show_step(1)
        elif s == 1:
            self._show_step(2)
        elif s == 2:
            path = self._install_dir.get().strip()
            if not path:
                messagebox.showwarning("Required", "Please choose a location.")
                return
            free = get_free_disk_gb(Path(path).anchor or "C:\\")
            if free < MIN_DISK_GB:
                messagebox.showerror("Not Enough Space",
                    f"Need {MIN_DISK_GB:.0f} GB free, only {free:.1f} GB available.")
                return
            self._show_step(3)
        elif s == 3:
            cmd, ver = find_python311()
            if not cmd:
                messagebox.showwarning("Python 3.11 Required",
                    "Python 3.11 not found.\nInstall it using the link on this page, then click Next.")
                self._show_step(3)
                return
            self._python_cmd = cmd
            self._python_ver = ver
            self._show_step(4)
        elif s == 4:
            key = self._anthropic_key.get().strip()
            if not key:
                messagebox.showwarning("Required", "Anthropic API key is required.")
                return
            self._next_btn.config(state=tk.DISABLED, text="Validating...")
            def _v():
                ok, msg = test_anthropic_key(key)
                self.after(0, lambda: self._next_btn.config(state=tk.NORMAL, text="Validate & Continue →"))
                if ok:
                    self.after(0, lambda: self._show_step(5))
                else:
                    self.after(0, lambda: messagebox.showerror("API Key Error", msg))
            threading.Thread(target=_v, daemon=True).start()
        elif s == 5:
            token = self._hf_token.get().strip()
            if not token:
                messagebox.showwarning("Required", "HuggingFace token is required.")
                return
            self._next_btn.config(state=tk.DISABLED, text="Validating...")
            def _v():
                ok, msg = test_hf_token(token)
                self.after(0, lambda: self._next_btn.config(state=tk.NORMAL, text="Validate & Continue →"))
                if ok:
                    self.after(0, lambda: self._show_step(6))
                else:
                    self.after(0, lambda: messagebox.showerror("Token Error", msg))
            threading.Thread(target=_v, daemon=True).start()
        elif s == 6:
            token = self._hf_token.get().strip()
            self._next_btn.config(state=tk.DISABLED, text="Checking model access...")
            def _v():
                ok, msg = check_pyannote_access(token)
                self.after(0, lambda: self._next_btn.config(
                    state=tk.NORMAL, text="I've Accepted Both →"))
                if ok:
                    self.after(0, lambda: self._start_install())
                else:
                    self.after(0, lambda: messagebox.showerror("Model Terms Not Accepted",
                        f"{msg}\n\nPlease accept the terms for both models before continuing."))
            threading.Thread(target=_v, daemon=True).start()
        elif s == 8:
            self.destroy()

    def _start_install(self):
        self._show_step(7)
        threading.Thread(target=self._run_install, daemon=True).start()

    def _back(self):
        if 0 < self._step < 7:
            self._show_step(self._step - 1)

    # ── Page helpers ──────────────────────────────────────────────────

    def _heading(self, text):
        tk.Label(self._page, text=text, bg=BG, fg=TEXT,
                 font=FONT_TITLE, anchor="w").pack(anchor="w", pady=(0, 6))

    def _sub(self, text):
        tk.Label(self._page, text=text, bg=BG, fg=MUTED, font=FONT_BODY,
                 anchor="w", justify="left", wraplength=500
                 ).pack(anchor="w", pady=(0, 10))

    def _card(self, bg=PANEL, pady=10, padx=12):
        f = tk.Frame(self._page, bg=bg, pady=pady, padx=padx)
        f.pack(fill=tk.X, pady=(0, 8))
        return f

    def _guide_card(self, title, steps):
        c = self._card(bg=ACCENT_BG)
        tk.Label(c, text=title, bg=ACCENT_BG, fg=ACCENT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        for s in steps:
            tk.Label(c, text=s, bg=ACCENT_BG, fg=TEXT,
                     font=FONT_SMALL, anchor="w", justify="left"
                     ).pack(anchor="w", pady=1)

    def _link_btn(self, text, url, pady=(0, 10)):
        tk.Button(self._page, text=text, bg=ACCENT2, fg=TEXT,
                  font=FONT_BODY, relief=tk.FLAT, padx=14, pady=6,
                  cursor="hand2", command=lambda: webbrowser.open(url)
                  ).pack(anchor="w", pady=pady)

    def _key_field(self, var):
        row = tk.Frame(self._page, bg=BG)
        row.pack(fill=tk.X, pady=(0, 6))
        entry = tk.Entry(row, textvariable=var, bg=INPUT, fg=TEXT,
                         font=FONT_MONO, insertbackground=TEXT,
                         relief=tk.FLAT, highlightbackground=BORDER,
                         highlightthickness=1, show="•")
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(0, 8))
        tk.Button(row, text="Show/Hide", bg=PANEL, fg=MUTED,
                  font=FONT_SMALL, relief=tk.FLAT, cursor="hand2",
                  command=lambda e=entry: e.config(
                      show="" if e.cget("show") == "•" else "•")
                  ).pack(side=tk.LEFT)

    # ── Pages ─────────────────────────────────────────────────────────

    def _pg_welcome(self):
        self._back_btn.config(state=tk.DISABLED)
        self._next_btn.config(text="Get Started →", state=tk.NORMAL)
        self._heading(f"Welcome to {APP_NAME}")
        self._sub("This wizard sets up everything you need, step by step.")
        for icon, title, desc in [
            ("🔍", "System Check",   "Verifies your PC meets requirements"),
            ("🐍", "Python 3.11",    "Installs the right Python version"),
            ("📦", "AI Libraries",   "PyTorch, Whisper, pyannote, Anthropic SDK"),
            ("🔑", "API Keys",       "Guided setup with live key validation"),
            ("✅", "Model Terms",    "Step-by-step HuggingFace model acceptance"),
            ("🖥",  "Desktop Icon",  "Creates a shortcut to launch the app"),
        ]:
            row = tk.Frame(self._page, bg=PANEL, pady=7, padx=12)
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=icon, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 14)).pack(side=tk.LEFT, padx=(0, 10))
            col = tk.Frame(row, bg=PANEL)
            col.pack(side=tk.LEFT)
            tk.Label(col, text=title, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 10, "bold"), anchor="w").pack(anchor="w")
            tk.Label(col, text=desc, bg=PANEL, fg=MUTED,
                     font=("Segoe UI", 9), anchor="w").pack(anchor="w")
        tk.Label(self._page,
                 text="Total time: 15–25 minutes depending on internet speed.",
                 bg=BG, fg=HINT, font=FONT_SMALL).pack(anchor="w", pady=(10, 0))

    def _pg_syscheck(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="Next →", state=tk.NORMAL)
        self._heading("System Check")
        self._sub("Checking your PC before we begin...")

        self._has_nvidia = has_nvidia_gpu()
        free_gb = get_free_disk_gb(Path(self._install_dir.get()).anchor or "C:\\")
        disk_ok = free_gb >= MIN_DISK_GB
        admin   = is_admin()

        try:
            import urllib.request
            urllib.request.urlopen("https://pypi.org", timeout=5)
            net_ok = True
        except Exception:
            net_ok = False

        checks = [
            ("Administrator rights", admin,
             "Running as administrator ✓" if admin
             else "Not admin — if install fails, right-click and 'Run as administrator'"),
            ("Disk space", disk_ok,
             f"{free_gb:.1f} GB free ✓" if disk_ok
             else f"Only {free_gb:.1f} GB free — need {MIN_DISK_GB:.0f} GB. Free up space first."),
            ("GPU", True,
             "Nvidia GPU detected — fast GPU processing will be used ✓"
             if self._has_nvidia
             else "No Nvidia GPU — CPU processing will be used (slower but works fine)"),
            ("Internet", net_ok,
             "Internet connection available ✓" if net_ok
             else "Cannot reach pypi.org — check internet or contact IT about firewall"),
        ]

        for name, ok, msg in checks:
            bg = OK_BG if ok else WARN_BG
            fg = OK_FG if ok else WARN_FG
            row = tk.Frame(self._page, bg=bg, pady=8, padx=12)
            row.pack(fill=tk.X, pady=3)
            tk.Label(row, text="✓" if ok else "⚠", bg=bg, fg=fg,
                     font=("Segoe UI", 12, "bold"), width=2).pack(side=tk.LEFT)
            col = tk.Frame(row, bg=bg)
            col.pack(side=tk.LEFT, padx=(8, 0))
            tk.Label(col, text=name, bg=bg, fg=fg,
                     font=("Segoe UI", 10, "bold"), anchor="w").pack(anchor="w")
            tk.Label(col, text=msg, bg=bg, fg=fg,
                     font=("Segoe UI", 9), anchor="w",
                     wraplength=460, justify="left").pack(anchor="w")

        if not net_ok:
            self._next_btn.config(state=tk.DISABLED)
            tk.Label(self._page,
                     text="⚠ Internet required. Contact IT if on a corporate network.",
                     bg=BG, fg=WARN_FG, font=FONT_SMALL).pack(anchor="w", pady=(6, 0))

    def _pg_location(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="Next →", state=tk.NORMAL)
        self._heading("Install Location")
        self._sub("Choose where Meeting Recorder will be installed.")
        row = tk.Frame(self._page, bg=BG)
        row.pack(fill=tk.X, pady=(0, 10))
        entry = tk.Entry(row, textvariable=self._install_dir, bg=INPUT, fg=TEXT,
                         font=FONT_BODY, insertbackground=TEXT, relief=tk.FLAT,
                         highlightbackground=BORDER, highlightthickness=1)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(0, 8))
        tk.Button(row, text="Browse", bg=PANEL, fg=TEXT, font=FONT_BODY,
                  relief=tk.FLAT, padx=12, pady=6, cursor="hand2",
                  command=lambda: self._install_dir.set(
                      filedialog.askdirectory(title="Choose Install Folder").replace("/", "\\"))
                  ).pack(side=tk.LEFT)
        c = self._card(bg=ACCENT_BG)
        tk.Label(c, text="What gets installed here:", bg=ACCENT_BG, fg=ACCENT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        for item in ["• Python virtual environment", "• AI libraries (~3–4 GB)",
                     "• Application files", "• API key config (stays on your machine)"]:
            tk.Label(c, text=item, bg=ACCENT_BG, fg=TEXT,
                     font=FONT_SMALL).pack(anchor="w")
        tk.Label(self._page, text=f"Required: {MIN_DISK_GB:.0f} GB free",
                 bg=BG, fg=HINT, font=FONT_SMALL).pack(anchor="w", pady=(4, 0))

    def _pg_python(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="Check Again →")
        self._heading("Python 3.11")
        cmd, ver = find_python311()
        if cmd:
            self._python_cmd = cmd
            self._python_ver = ver
            self._next_btn.config(text="Next →")
            c = tk.Frame(self._page, bg=OK_BG, pady=12, padx=14)
            c.pack(fill=tk.X, pady=(0, 10))
            tk.Label(c, text=f"✓  Python {ver} found — ready!",
                     bg=OK_BG, fg=OK_FG, font=("Segoe UI", 11, "bold")).pack(anchor="w")
            tk.Label(c, text="Click Next to continue.",
                     bg=OK_BG, fg=OK_FG, font=FONT_SMALL).pack(anchor="w")
        else:
            c = tk.Frame(self._page, bg=WARN_BG, pady=10, padx=14)
            c.pack(fill=tk.X, pady=(0, 10))
            tk.Label(c, text="⚠  Python 3.11 not found",
                     bg=WARN_BG, fg=WARN_FG,
                     font=("Segoe UI", 11, "bold")).pack(anchor="w")
            tk.Label(c, text="Meeting Recorder requires Python 3.11 specifically.",
                     bg=WARN_BG, fg=WARN_FG, font=FONT_SMALL).pack(anchor="w")
            self._guide_card("Install Python 3.11 (follow these steps exactly):", [
                "1. Click 'Download Python 3.11.9' below",
                "2. Run the downloaded file",
                "3. ⚠ IMPORTANT: Check 'Add Python to PATH' at the bottom of the installer!",
                "4. Click 'Install Now' and wait",
                "5. Click 'Close' when done",
                "6. Come back here and click 'Check Again'",
            ])
            tk.Button(self._page, text="⬇  Download Python 3.11.9",
                      bg=ACCENT2, fg=TEXT, font=FONT_BODY, relief=tk.FLAT,
                      padx=14, pady=8, cursor="hand2",
                      command=lambda: webbrowser.open(PY_URL)
                      ).pack(anchor="w")

    def _pg_anthropic(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="Validate & Continue →", state=tk.NORMAL)
        self._heading("Anthropic API Key")
        self._sub("Powers AI meeting summarization.\nSaved only on your computer — never shared.")
        self._guide_card("Get your key in 5 minutes:", [
            "1. Click 'Open Anthropic Console' below",
            "2. Sign up for a free account (or log in)",
            "3. Click 'API Keys' in the left sidebar",
            "4. Click '+ Create Key', name it 'Meeting Recorder'",
            "5. Copy the key shown (starts with sk-ant-...)",
            "6. Click 'Billing' → 'Add Credits' — add at least $5",
            "7. Paste your key in the box below",
        ])
        self._link_btn("🌐  Open Anthropic Console",
                       "https://console.anthropic.com/settings/keys")
        tk.Label(self._page, text="Paste your Anthropic API key:",
                 bg=BG, fg=TEXT, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        self._key_field(self._anthropic_key)
        tk.Label(self._page,
                 text="~$0.04 per summary  |  $5 ≈ 125 summaries  |  Key will be validated before install",
                 bg=BG, fg=HINT, font=FONT_SMALL).pack(anchor="w")

    def _pg_huggingface(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="Validate & Continue →", state=tk.NORMAL)
        self._heading("HuggingFace Token")
        self._sub("Required for speaker detection (who said what).\n100% free — no payment needed.")
        self._guide_card("Get your free token in 3 minutes:", [
            "1. Click 'Open HuggingFace' below",
            "2. Sign up for a free account (or log in)",
            "3. Click your profile picture → Settings",
            "4. Click 'Access Tokens' in the left sidebar",
            "5. Click 'New token'",
            "6. Name it 'Meeting Recorder', Type = 'Read'",
            "7. Click 'Generate a token'",
            "8. Copy the token (starts with hf_...)",
            "9. Paste it in the box below",
        ])
        self._link_btn("🌐  Open HuggingFace", "https://huggingface.co/settings/tokens")
        tk.Label(self._page, text="Paste your HuggingFace token:",
                 bg=BG, fg=TEXT, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        self._key_field(self._hf_token)
        tk.Label(self._page, text="Token will be validated before install",
                 bg=BG, fg=HINT, font=FONT_SMALL).pack(anchor="w")

    def _pg_pyannote(self):
        self._back_btn.config(state=tk.NORMAL)
        self._next_btn.config(text="I've Accepted Both →", state=tk.NORMAL)
        self._heading("Accept AI Model Terms")
        self._sub("Two AI models require you to accept their terms on HuggingFace.\nThis is a one-time step.")

        warn = tk.Frame(self._page, bg=WARN_BG, pady=8, padx=12)
        warn.pack(fill=tk.X, pady=(0, 8))
        tk.Label(warn, text="⚠  Make sure you are logged in to HuggingFace first!",
                 bg=WARN_BG, fg=WARN_FG,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")

        for num, title, url, steps in [
            ("1", "Speaker Diarization Model",
             "https://huggingface.co/pyannote/speaker-diarization-3.1",
             ["1. Click 'Open Model 1' below",
              "2. Scroll down to the agreement box",
              "3. Fill in your info and click 'Agree and access repository'",
              "4. You'll see model files — this means it worked ✓"]),
            ("2", "Segmentation Model",
             "https://huggingface.co/pyannote/segmentation-3.0",
             ["1. Click 'Open Model 2' below",
              "2. Same process — agree and access repository",
              "3. You'll see model files — this means it worked ✓"]),
        ]:
            c = self._card(bg=PANEL)
            tk.Label(c, text=f"Step {num} of 2 — {title}",
                     bg=PANEL, fg=ACCENT,
                     font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
            for s in steps:
                tk.Label(c, text=s, bg=PANEL, fg=TEXT,
                         font=FONT_SMALL, anchor="w").pack(anchor="w", pady=1)
            tk.Button(c, text=f"🌐  Open Model {num}", bg=ACCENT2, fg=TEXT,
                      font=FONT_SMALL, relief=tk.FLAT, padx=12, pady=4,
                      cursor="hand2", command=lambda u=url: webbrowser.open(u)
                      ).pack(anchor="w", pady=(6, 0))

        tk.Label(self._page,
                 text="After accepting both, click 'I've Accepted Both' — the installer will verify automatically.",
                 bg=BG, fg=HINT, font=FONT_SMALL, wraplength=500).pack(anchor="w", pady=(6, 0))

    def _pg_install(self):
        self._back_btn.config(state=tk.DISABLED)
        self._next_btn.config(state=tk.DISABLED)
        self._heading("Installing...")
        self._prog_label = tk.Label(self._page, text="Preparing...",
                                     bg=BG, fg=MUTED, font=FONT_BODY, anchor="w")
        self._prog_label.pack(anchor="w", pady=(0, 8))
        style = ttk.Style()
        style.theme_use("default")
        style.configure("M.Horizontal.TProgressbar",
                         troughcolor=PANEL, background=ACCENT,
                         borderwidth=0, thickness=8)
        self._progress = ttk.Progressbar(
            self._page, style="M.Horizontal.TProgressbar",
            length=660, mode="determinate")
        self._progress.pack(anchor="w", pady=(0, 12))
        self._log = tk.Text(self._page, bg=PANEL, fg=MUTED,
                            font=("Segoe UI", 9), relief=tk.FLAT,
                            state=tk.DISABLED, height=16, wrap=tk.WORD,
                            highlightthickness=0)
        self._log.pack(fill=tk.BOTH, expand=True)

    def _pg_done(self):
        self._back_btn.config(state=tk.DISABLED)
        self._next_btn.config(text="Finish", state=tk.NORMAL)
        self._heading("All Done! 🎉")
        self._sub("Meeting Recorder is installed and ready to use.")
        for icon, label in [
            ("✓", "Virtual environment created"),
            ("✓", f"PyTorch installed ({'GPU accelerated' if self._has_nvidia else 'CPU mode'})"),
            ("✓", "All AI libraries installed"),
            ("✓", "API keys saved"),
            ("✓", "HuggingFace login complete"),
            ("✓", "Desktop shortcut created"),
        ]:
            row = tk.Frame(self._page, bg=BG)
            row.pack(anchor="w", pady=2)
            tk.Label(row, text=icon, bg=BG, fg=ACCENT,
                     font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=(0, 10))
            tk.Label(row, text=label, bg=BG, fg=TEXT, font=FONT_BODY).pack(side=tk.LEFT)
        note = tk.Frame(self._page, bg=ACCENT_BG, pady=10, padx=14)
        note.pack(fill=tk.X, pady=(14, 0))
        for line in [
            "ℹ  First launch downloads AI models (~700MB) — takes 2-3 min, one time only.",
            "ℹ  If Windows shows a security warning, click 'More info' then 'Run anyway'.",
            "ℹ  A README.txt with quick start instructions is in your install folder.",
        ]:
            tk.Label(note, text=line, bg=ACCENT_BG, fg=ACCENT,
                     font=("Segoe UI", 9), anchor="w", justify="left"
                     ).pack(anchor="w", pady=2)
        tk.Button(self._page, text="🚀  Launch Meeting Recorder Now",
                  bg=ACCENT2, fg=TEXT, font=("Segoe UI", 12, "bold"),
                  relief=tk.FLAT, padx=20, pady=10, cursor="hand2",
                  command=lambda: [
                      os.startfile(str(Path(self._install_dir.get()) / "launch.vbs"))
                      if (Path(self._install_dir.get()) / "launch.vbs").exists() else None,
                      self.destroy()
                  ]).pack(anchor="w", pady=(16, 0))

    # ── Install logic ──────────────────────────────────────────────────

    def _log_write(self, msg):
        def _do():
            self._log.config(state=tk.NORMAL)
            self._log.insert(tk.END, msg + "\n")
            self._log.see(tk.END)
            self._log.config(state=tk.DISABLED)
        self.after(0, _do)

    def _set_progress(self, value, label=""):
        def _do():
            self._progress["value"] = value
            if label:
                self._prog_label.config(text=label)
        self.after(0, _do)

    def _run_install(self):
        try:
            install_dir = Path(self._install_dir.get())
            install_dir.mkdir(parents=True, exist_ok=True)
            python = self._python_cmd or "python"

            self._set_progress(3, "Creating virtual environment...")
            self._log_write("→ Creating virtual environment...")
            r = subprocess.run([python, "-m", "venv", str(install_dir / ".venv")],
                               capture_output=True, text=True)
            if r.returncode != 0:
                raise RuntimeError(
                    "Failed to create virtual environment.\n"
                    "Try running the installer as Administrator.\n"
                    f"Error: {r.stderr[:200]}")
            self._log_write("  ✓ Virtual environment created")

            pip = str(install_dir / ".venv" / "Scripts" / "pip.exe")
            py  = str(install_dir / ".venv" / "Scripts" / "python.exe")

            self._set_progress(6, "Upgrading pip...")
            subprocess.run([py, "-m", "pip", "install", "--upgrade", "pip"],
                           capture_output=True)
            self._log_write("  ✓ pip upgraded")

            # PyTorch — GPU or CPU
            if self._has_nvidia:
                self._set_progress(8, "Installing PyTorch with GPU support (large download, please wait)...")
                self._log_write("→ Installing PyTorch CUDA (GPU)...")
                r = subprocess.run(
                    [pip, "install", "torch==2.6.0", "torchaudio==2.6.0",
                     "--index-url", "https://download.pytorch.org/whl/cu124"],
                    capture_output=True, text=True)
                if r.returncode != 0:
                    self._log_write("  ⚠ CUDA failed, falling back to CPU...")
                    r = subprocess.run(
                        [pip, "install", "torch==2.6.0", "torchaudio==2.6.0",
                         "--index-url", "https://download.pytorch.org/whl/cpu"],
                        capture_output=True, text=True)
                    if r.returncode != 0:
                        raise RuntimeError(
                            "PyTorch install failed.\nCheck internet connection.\n" + r.stderr[:200])
                    self._log_write("  ✓ PyTorch CPU installed")
                else:
                    self._log_write("  ✓ PyTorch CUDA installed")
            else:
                self._set_progress(8, "Installing PyTorch CPU (large download, please wait)...")
                self._log_write("→ Installing PyTorch CPU...")
                r = subprocess.run(
                    [pip, "install", "torch==2.6.0", "torchaudio==2.6.0",
                     "--index-url", "https://download.pytorch.org/whl/cpu"],
                    capture_output=True, text=True)
                if r.returncode != 0:
                    raise RuntimeError(
                        "PyTorch install failed.\nCheck internet connection.\n" + r.stderr[:200])
                self._log_write("  ✓ PyTorch CPU installed")

            self._set_progress(32)

            # PyAudio with fallback
            self._set_progress(33, "Installing PyAudio...")
            self._log_write("→ Installing PyAudio...")
            r = subprocess.run([pip, "install", "pyaudio"], capture_output=True, text=True)
            if r.returncode != 0:
                self._log_write("  ⚠ PyAudio failed — trying pipwin fallback...")
                subprocess.run([pip, "install", "pipwin"], capture_output=True)
                r2 = subprocess.run([py, "-m", "pipwin", "install", "pyaudio"],
                                    capture_output=True, text=True)
                if r2.returncode != 0:
                    self._log_write("  ⚠ PyAudio could not be installed automatically.")
                    self._log_write("    Audio recording may not work.")
                    self._log_write("    Fix: install Microsoft C++ Build Tools from microsoft.com")
                else:
                    self._log_write("  ✓ PyAudio installed via pipwin")
            else:
                self._log_write("  ✓ PyAudio installed")

            # Core requirements
            total = len(REQUIREMENTS_BASE)
            for i, pkg in enumerate(REQUIREMENTS_BASE):
                pct  = 36 + int((i / total) * 38)
                name = pkg.split("==")[0]
                self._set_progress(pct, f"Installing {name}...")
                self._log_write(f"→ Installing {name}...")
                r = subprocess.run([pip, "install", pkg], capture_output=True, text=True)
                if r.returncode != 0:
                    stderr_low = r.stderr.lower()
                    if any(w in stderr_low for w in ["firewall", "proxy", "ssl", "certificate"]):
                        raise RuntimeError(
                            f"Network blocked installing {name}.\n"
                            "Your corporate firewall may be blocking pip.\n"
                            "Ask IT to whitelist pypi.org and files.pythonhosted.org")
                    self._log_write(f"  ⚠ {name} failed: {r.stderr[:80]}")
                else:
                    self._log_write(f"  ✓ {name}")

            # Pin numpy
            self._set_progress(76, "Pinning numpy version...")
            self._log_write("→ Pinning numpy...")
            subprocess.run([pip, "install", "numpy==2.1.3"], capture_output=True)
            self._log_write("  ✓ numpy==2.1.3")

            # Write app files
            self._set_progress(80, "Writing application files...")
            self._log_write("→ Writing application files...")
            self._write_app_files(install_dir)
            self._log_write("  ✓ Application files written")

            # Write .env
            self._set_progress(85, "Saving configuration...")
            self._log_write("→ Saving configuration...")
            (install_dir / ".env").write_text(
                f"ANTHROPIC_API_KEY={self._anthropic_key.get().strip()}\n"
                f"HF_TOKEN={self._hf_token.get().strip()}\n"
                f"WHISPER_MODEL=base\n"
                f"MAX_SPEAKERS=8\n"
                f"RECORDINGS_DIR=recordings\n",
                encoding="utf-8")
            self._log_write("  ✓ Configuration saved")

            # HuggingFace login
            self._set_progress(88, "Logging in to HuggingFace...")
            self._log_write("→ HuggingFace login...")
            subprocess.run(
                [py, "-c",
                 f'from huggingface_hub import login; login(token="{self._hf_token.get().strip()}")'],
                capture_output=True)
            self._log_write("  ✓ HuggingFace login complete")

            # Patch main.py
            self._set_progress(91, "Applying compatibility patches...")
            self._log_write("→ Applying patches...")
            self._patch_main(install_dir)
            self._log_write("  ✓ Patches applied")

            # Shortcut
            self._set_progress(95, "Creating desktop shortcut...")
            self._log_write("→ Creating desktop shortcut...")
            vbs = create_vbs(str(install_dir))
            create_shortcut(str(install_dir), vbs)
            self._log_write("  ✓ Desktop shortcut created")

            # README
            self._set_progress(98, "Writing quick start guide...")
            self._write_readme(install_dir)
            self._log_write("  ✓ README.txt written")

            self._set_progress(100, "Done!")
            self._log_write("\n" + "=" * 48)
            self._log_write("  INSTALLATION COMPLETE!")
            self._log_write("=" * 48)
            self._log_write("Double-click the desktop shortcut to launch.")
            self.after(1500, lambda: self._show_step(8))

        except Exception as e:
            self._log_write(f"\n{'='*48}\n  INSTALLATION FAILED\n{'='*48}")
            self._log_write(f"Error: {e}")
            self._log_write("\nCommon fixes:")
            self._log_write("  • Right-click installer → Run as Administrator")
            self._log_write("  • Check internet connection")
            self._log_write("  • Ask IT to whitelist pypi.org if on corporate network")
            self._set_progress(self._progress["value"], "Failed — see log above")
            self.after(0, lambda: messagebox.showerror("Failed",
                f"{str(e)[:300]}\n\nSee the log for details."))
            self.after(0, lambda: self._next_btn.config(
                state=tk.NORMAL, text="Retry",
                command=lambda: threading.Thread(
                    target=self._run_install, daemon=True).start()))

    def _write_app_files(self, install_dir: Path):
        for rel_path, content in APP_FILES.items():
            dest = install_dir / rel_path
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_text(content, encoding="utf-8")
        if APP_ICON_B64:
            import base64
            (install_dir / "meeting_recorder.ico").write_bytes(
                base64.b64decode(APP_ICON_B64))

    def _patch_main(self, install_dir: Path):
        main_path = install_dir / "main.py"
        if not main_path.exists():
            return
        content = main_path.read_text(encoding="utf-8")
        if "np.NaN = np.nan" in content:
            return
        patch = '''import torch
import numpy as np
from torch.torch_version import TorchVersion

if not hasattr(np, "NaN"):
    np.NaN = np.nan
if not hasattr(np, "NAN"):
    np.NAN = np.nan
torch.serialization.add_safe_globals([TorchVersion])
_orig_load = torch.load
def _patched_load(f, *args, **kwargs):
    kwargs["weights_only"] = False
    return _orig_load(f, *args, **kwargs)
torch.load = _patched_load

'''
        main_path.write_text(patch + content, encoding="utf-8")

    def _write_readme(self, install_dir: Path):
        (install_dir / "README.txt").write_text(
            f"""{APP_NAME} — Quick Start Guide
{'=' * 44}

LAUNCHING
  Double-click the desktop shortcut.
  If Windows warns "unknown publisher": click More info → Run anyway.

FIRST LAUNCH
  Downloads AI models (~700MB). Takes 2-5 min. Happens once only.

HOW TO USE
  1. Select your microphone
  2. Click Start Recording before your meeting
  3. Click Stop Recording when done
  4. Click Process & Transcribe — wait for completion
  5. Rename speakers if needed
  6. Click Summarize for an AI summary
  7. Click Email Summary to send to your Outlook inbox

RECORDINGS
  Saved in: {install_dir}\\recordings\\
  Click Open Recordings in the app to browse.

TROUBLESHOOTING
  App won't start   → Run as Administrator
  No audio          → Check microphone in Windows Settings → Privacy → Microphone
  Processing fails  → Verify HuggingFace model terms were accepted:
                      huggingface.co/pyannote/speaker-diarization-3.1
                      huggingface.co/pyannote/segmentation-3.0
  Summary fails     → Check Anthropic billing at console.anthropic.com
  Email fails       → Make sure Outlook desktop is open

Install location: {install_dir}
""", encoding="utf-8")

    def _on_close(self):
        if self._step == 7:
            if not messagebox.askyesno("Cancel?", "Installation in progress. Cancel?"):
                return
        self.destroy()


if __name__ == "__main__":
    app = Installer()
    app.mainloop()
