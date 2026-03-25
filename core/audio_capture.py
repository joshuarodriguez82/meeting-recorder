import threading
from typing import Callable, List, Optional
import numpy as np
import sounddevice as sd
from utils.logger import get_logger

logger = get_logger(__name__)

SAMPLE_RATE = 16000
BLOCK_SIZE = 1024


def list_input_devices() -> List[dict]:
    devices = []
    for idx, dev in enumerate(sd.query_devices()):
        if dev["max_input_channels"] > 0:
            devices.append({
                "index": idx,
                "name": dev["name"],
                "max_input_channels": dev["max_input_channels"],
                "default_samplerate": dev["default_samplerate"],
            })
    return devices


def list_output_devices() -> List[dict]:
    devices = []
    for idx, dev in enumerate(sd.query_devices()):
        if dev["max_output_channels"] > 0:
            devices.append({
                "index": idx,
                "name": dev["name"],
                "max_output_channels": dev["max_output_channels"],
                "default_samplerate": dev["default_samplerate"],
            })
    return devices


class AudioCapture:

    def __init__(
        self,
        mic_device_index: Optional[int],
        output_device_index: Optional[int],
        on_chunk: Callable[[np.ndarray], None],
    ):
        self._mic_idx = mic_device_index
        self._out_idx = output_device_index
        self._on_chunk = on_chunk
        self._streams: List[sd.InputStream] = []
        self._lock = threading.Lock()
        self._mic_buffer: Optional[np.ndarray] = None
        self._out_buffer: Optional[np.ndarray] = None
        self._running = False
        self._chunk_count = 0
        self.actual_sr = SAMPLE_RATE

    def start(self) -> None:
        self._running = True
        self._chunk_count = 0
        logger.info(f"Starting capture: mic={self._mic_idx}, output={self._out_idx}")

        try:
            if self._mic_idx is not None:
                dev_info = sd.query_devices(self._mic_idx)
                native_sr = int(dev_info["default_samplerate"])
                channels = min(2, int(dev_info["max_input_channels"]))
                self.actual_sr = native_sr
                logger.info(f"Mic device: {dev_info['name']}, sr={native_sr}, ch={channels}")
                mic_stream = sd.InputStream(
                    device=self._mic_idx,
                    channels=channels,
                    samplerate=native_sr,
                    blocksize=BLOCK_SIZE,
                    dtype="float32",
                    callback=self._mic_callback,
                )
                mic_stream.start()
                self._streams.append(mic_stream)
                logger.info(f"Mic stream started at {native_sr}Hz, {channels}ch")

            if self._out_idx is not None:
                try:
                    dev_info = sd.query_devices(self._out_idx)
                    native_sr = int(dev_info["default_samplerate"])
                    channels = min(2, int(dev_info["max_output_channels"]))
                    logger.info(f"Output device: {dev_info['name']}, sr={native_sr}, ch={channels}")
                    out_stream = sd.InputStream(
                        device=self._out_idx,
                        channels=channels,
                        samplerate=native_sr,
                        blocksize=BLOCK_SIZE,
                        dtype="float32",
                        callback=self._out_callback,
                    )
                    out_stream.start()
                    self._streams.append(out_stream)
                    logger.info(f"Output stream started at {native_sr}Hz, {channels}ch")
                except Exception as e:
                    logger.warning(f"Loopback capture unavailable: {e}. Mic only.")
                    self._out_idx = None

        except Exception as e:
            self._close_all_streams()
            self._running = False
            raise

    def stop(self) -> None:
        self._running = False
        self._close_all_streams()
        logger.info(f"Audio capture stopped. Total chunks captured: {self._chunk_count}")

    def _close_all_streams(self) -> None:
        for stream in self._streams:
            try:
                stream.stop()
                stream.close()
            except Exception as e:
                logger.warning(f"Error closing stream: {e}")
        self._streams.clear()

    def _mic_callback(self, indata: np.ndarray, frames: int, time, status) -> None:
        try:
            if status:
                logger.warning(f"Mic stream status: {status}")
            chunk = indata.mean(axis=1) if indata.ndim > 1 else indata[:, 0].copy()
            self._chunk_count += 1
            if self._chunk_count % 100 == 0:
                logger.info(f"Mic chunks received: {self._chunk_count}")
            with self._lock:
                self._mic_buffer = chunk
                self._try_merge()
        except Exception as e:
            logger.error(f"Error in mic callback: {e}")

    def _out_callback(self, indata: np.ndarray, frames: int, time, status) -> None:
        try:
            if status:
                logger.warning(f"Output stream status: {status}")
            chunk = indata.mean(axis=1) if indata.ndim > 1 else indata[:, 0].copy()
            with self._lock:
                self._out_buffer = chunk
                self._try_merge()
        except Exception as e:
            logger.error(f"Error in output callback: {e}")

    def _try_merge(self) -> None:
        if self._mic_idx is not None and self._out_idx is not None:
            if self._mic_buffer is not None and self._out_buffer is not None:
                size = min(len(self._mic_buffer), len(self._out_buffer))
                merged = (self._mic_buffer[:size] + self._out_buffer[:size]) / 2.0
                self._mic_buffer = None
                self._out_buffer = None
                self._safe_invoke(merged)
        elif self._mic_buffer is not None:
            chunk = self._mic_buffer
            self._mic_buffer = None
            self._safe_invoke(chunk)
        elif self._out_buffer is not None:
            chunk = self._out_buffer
            self._out_buffer = None
            self._safe_invoke(chunk)

    def _safe_invoke(self, chunk: np.ndarray) -> None:
        try:
            self._on_chunk(chunk)
        except Exception as e:
            logger.error(f"Error in on_chunk callback: {e}")