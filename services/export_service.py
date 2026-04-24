"""
Exports transcripts and summaries to text files.
Uses meeting display name if available for clean filenames.

Each export method accepts an optional `export_dir` override so a session
tagged with a client that has a configured folder writes its artifacts
there instead of the default recordings directory.
"""

from pathlib import Path
from typing import Optional

from models.session import Session
from utils.logger import get_logger

logger = get_logger(__name__)


class ExportService:

    def __init__(self, recordings_dir: str):
        self._dir = Path(recordings_dir)
        self._dir.mkdir(parents=True, exist_ok=True)

    def _base_name(self, session: Session) -> str:
        if session.display_name:
            safe = "".join(
                c if c.isalnum() or c in " -_" else ""
                for c in session.display_name
            ).strip()
            return safe or session.session_id
        return f"session_{session.session_id}"

    def _target_dir(self, export_dir: Optional[str]) -> Path:
        if export_dir and str(export_dir).strip():
            target = Path(str(export_dir).strip())
            target.mkdir(parents=True, exist_ok=True)
            return target
        return self._dir

    def export_transcript(
        self, session: Session, export_dir: Optional[str] = None,
    ) -> str:
        target = self._target_dir(export_dir)
        name = self._base_name(session)
        path = target / f"transcript_{name}.txt"
        lines = []
        if session.display_name:
            lines.append(f"Meeting: {session.display_name}")
            lines.append("=" * 60)
            lines.append("")
        lines.append(session.full_transcript())
        path.write_text("\n".join(lines), encoding="utf-8")
        logger.info(f"Transcript exported: {path}")
        return str(path)

    def export_summary(
        self, session: Session, export_dir: Optional[str] = None,
    ) -> str:
        if not session.summary:
            raise ValueError("No summary to export.")
        target = self._target_dir(export_dir)
        name = self._base_name(session)
        path = target / f"summary_{name}.txt"
        lines = []
        if session.display_name:
            lines.append(f"Meeting: {session.display_name}")
            lines.append("=" * 60)
            lines.append("")
        lines.append(session.summary)
        path.write_text("\n".join(lines), encoding="utf-8")
        logger.info(f"Summary exported: {path}")
        return str(path)

    def export_action_items(
        self, session: Session, export_dir: Optional[str] = None,
    ) -> str:
        if not session.action_items:
            raise ValueError("No action items to export.")
        target = self._target_dir(export_dir)
        name = self._base_name(session)
        path = target / f"action_items_{name}.txt"
        lines = []
        if session.display_name:
            lines.append(f"Meeting: {session.display_name}")
            lines.append("=" * 60)
            lines.append("")
        lines.append(session.action_items)
        path.write_text("\n".join(lines), encoding="utf-8")
        logger.info(f"Action items exported: {path}")
        return str(path)

    def export_decisions(
        self, session: Session, export_dir: Optional[str] = None,
    ) -> str:
        if not session.decisions:
            raise ValueError("No decisions to export.")
        target = self._target_dir(export_dir)
        name = self._base_name(session)
        path = target / f"decisions_{name}.txt"
        lines = []
        if session.display_name:
            lines.append(f"Meeting: {session.display_name}")
            lines.append("=" * 60)
            lines.append("")
        lines.append(session.decisions)
        path.write_text("\n".join(lines), encoding="utf-8")
        logger.info(f"Decisions exported: {path}")
        return str(path)

    def export_requirements(
        self, session: Session, export_dir: Optional[str] = None,
    ) -> str:
        if not session.requirements:
            raise ValueError("No requirements to export.")
        target = self._target_dir(export_dir)
        name = self._base_name(session)
        path = target / f"requirements_{name}.txt"
        lines = []
        if session.display_name:
            lines.append(f"Meeting: {session.display_name}")
            lines.append("=" * 60)
            lines.append("")
        lines.append(session.requirements)
        path.write_text("\n".join(lines), encoding="utf-8")
        logger.info(f"Requirements exported: {path}")
        return str(path)
