"""
Orchestrates the full recording lifecycle.
"""

import asyncio
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional

import numpy as np
from math import gcd
from scipy.signal import resample_poly

from config.settings import Settings
from core.audio_capture import AudioCapture
from core.diarization import DiarizationEngine
from core.transcription import TranscriptionEngine
from models.segment import Segment
from models.session import Session
from utils.audio_utils import save_wav
from utils.logger import get_logger

logger = get_logger(__name__)

TARGET_SR = 16000
MAX_CHUNKS = int((90 * 60 * TARGET_SR) / 1024)


def _resample(audio: np.ndarray, orig_sr: int) -> np.ndarray:
    if orig_sr == TARGET_SR:
        return audio.astype(np.float32)
    divisor = gcd(TARGET_SR, orig_sr)
    up = TARGET_SR // divisor
    down = orig_sr // divisor
    resampled = resample_poly(audio.astype(np.float64), up, down)
    return np.clip(resampled, -1.0, 1.0).astype(np.float32)


class RecordingService:

    def __init__(
        self,
        settings: Settings,
        transcription_engine: TranscriptionEngine,
        diarization_engine: DiarizationEngine,
        on_status: Optional[Callable[[str], None]] = None,
    ):
        self._settings = settings
        self._transcription = transcription_engine
        self._diarization = diarization_engine
        self._on_status = on_status or (lambda _: None)
        self._session: Optional[Session] = None
        self._capture: Optional[AudioCapture] = None
        self._audio_chunks: List[np.ndarray] = []
        self._chunks_lock = threading.Lock()
        self._recording = False
        self._capture_sr = TARGET_SR

    @property
    def current_session(self) -> Optional[Session]:
        return self._session

    @property
    def is_recording(self) -> bool:
        return self._recording

    def set_session(self, session: Session) -> None:
        """Allow an externally created session (e.g. loaded file) to be processed."""
        self._session = session

    def start_recording(
        self,
        mic_device_index: Optional[int],
        output_device_index: Optional[int],
    ) -> Session:
        if self._recording:
            raise RuntimeError("A recording is already in progress.")

        session_id = uuid.uuid4().hex[:8].upper()
        self._session = Session(session_id=session_id)

        with self._chunks_lock:
            self._audio_chunks = []

        self._recording = True
        self._capture = AudioCapture(
            mic_device_index=mic_device_index,
            output_device_index=output_device_index,
            on_chunk=self._on_audio_chunk,
        )

        try:
            self._capture.start()
            if hasattr(self._capture, 'actual_sr'):
                self._capture_sr = self._capture.actual_sr
            else:
                self._capture_sr = TARGET_SR
        except Exception as e:
            self._recording = False
            self._capture = None
            raise RuntimeError(f"Failed to start audio capture: {e}") from e

        self._on_status(f"Recording started — Session {session_id}")
        logger.info(f"Session {session_id} recording started.")
        return self._session

    def stop_recording(self) -> Optional[Session]:
        if not self._recording or not self._capture:
            return self._session

        self._recording = False
        self._capture.stop()

        capture_sr = self._capture_sr
        self._capture = None

        with self._chunks_lock:
            chunks_snapshot = list(self._audio_chunks)
            self._audio_chunks = []

        if self._session and chunks_snapshot:
            try:
                audio = np.concatenate(chunks_snapshot, axis=0)
                if audio.ndim == 2:
                    audio = audio.mean(axis=1)
                audio = _resample(audio, capture_sr)
                path = self._build_audio_path(self._session.session_id)
                save_wav(path, audio, TARGET_SR)
                self._session.audio_path = path
                self._session.ended_at = datetime.now()
                self._on_status("Recording saved. Ready to process.")
                logger.info(f"Audio saved to {path} ({len(audio)/TARGET_SR:.1f}s)")
            except Exception as e:
                logger.error(f"Failed to save audio: {e}")
                self._on_status(f"Error saving audio: {e}")
        elif self._session and not chunks_snapshot:
            logger.warning("Recording stopped with no audio chunks captured.")
            self._on_status("No audio was captured. Try again.")

        return self._session

    async def process_session(self) -> Session:
        if not self._session or not self._session.audio_path:
            raise RuntimeError("No recorded session to process.")

        self._on_status("__stage:transcribe:active__")
        raw_segments = await self._transcription.transcribe(self._session.audio_path)

        if not raw_segments:
            self._on_status("Transcription produced no output. Check audio quality.")
            return self._session

        self._on_status("__stage:transcribe:done____stage:diarize:active__")
        diarization_turns = await self._diarization.diarize(self._session.audio_path)

        self._on_status("__stage:diarize:done____stage:speakers:active__")
        attributed = DiarizationEngine.assign_speakers(raw_segments, diarization_turns)

        for raw in attributed:
            speaker = self._session.get_or_create_speaker(raw["speaker_id"])
            segment = Segment(
                speaker_id=speaker.speaker_id,
                start=raw["start"],
                end=raw["end"],
                text=raw["text"],
            )
            self._session.segments.append(segment)

        self._on_status("Processing complete.")
        logger.info(f"Session {self._session.session_id} processing complete.")
        return self._session

    def _on_audio_chunk(self, chunk: np.ndarray) -> None:
        if not self._recording:
            return
        with self._chunks_lock:
            if len(self._audio_chunks) >= MAX_CHUNKS:
                self._audio_chunks.pop(0)
            self._audio_chunks.append(chunk)

    def _build_audio_path(self, session_id: str) -> str:
        recordings_dir = Path(self._settings.recordings_dir)
        recordings_dir.mkdir(parents=True, exist_ok=True)
        return str(recordings_dir / f"session_{session_id}.wav")
