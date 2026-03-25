"""
Persists and loads session data as JSON.
Uses atomic write (temp file + rename) to prevent corrupt JSON on crash.
"""

import json
import os
import tempfile
from pathlib import Path
from typing import Optional

from models.session import Session
from utils.logger import get_logger

logger = get_logger(__name__)


class SessionService:
    """Handles JSON serialization of Session objects."""

    def __init__(self, recordings_dir: str):
        self._recordings_dir = Path(recordings_dir)
        self._recordings_dir.mkdir(parents=True, exist_ok=True)

    def save(self, session: Session) -> str:
        """
        Serialize a session to a JSON file using an atomic write.

        Writes to a temporary file first, then renames it to the final path.
        This ensures the target file is never left in a half-written state
        if the process is interrupted mid-write.

        Args:
            session: The completed Session object.

        Returns:
            The path of the saved JSON file.

        Raises:
            OSError: If writing or renaming fails.
        """
        final_path = self._recordings_dir / f"session_{session.session_id}.json"
        data = json.dumps(session.to_dict(), indent=2, ensure_ascii=False)

        # FIX #10: write to temp file in same directory, then atomic rename
        try:
            fd, tmp_path = tempfile.mkstemp(
                dir=self._recordings_dir,
                suffix=".json.tmp",
            )
            try:
                with os.fdopen(fd, "w", encoding="utf-8") as f:
                    f.write(data)
                    f.flush()
                    os.fsync(f.fileno())  # Flush OS buffers before rename
                os.replace(tmp_path, final_path)  # Atomic on POSIX & Windows
            except Exception:
                # Clean up temp file if rename or write fails
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass
                raise
        except Exception as e:
            raise OSError(f"Failed to save session {session.session_id}: {e}") from e

        logger.info(f"Session atomically saved: {final_path}")
        return str(final_path)

    def load(self, session_id: str) -> Optional[dict]:
        """
        Load a session JSON by ID.

        Args:
            session_id: The session identifier.

        Returns:
            Parsed session dict, or None if not found.

        Raises:
            ValueError: If the file exists but contains invalid JSON.
        """
        path = self._recordings_dir / f"session_{session_id}.json"
        if not path.exists():
            logger.warning(f"Session file not found: {path}")
            return None
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(f"Corrupt session file {path}: {e}") from e
