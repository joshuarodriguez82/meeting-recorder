from __future__ import annotations
import datetime
from typing import Dict, List, Optional
from models.speaker import Speaker
from models.segment import Segment


class Session:

    def __init__(self, session_id: str):
        self.session_id: str = session_id
        self.display_name: str = ""
        self.started_at: datetime.datetime = datetime.datetime.now()
        self.ended_at: Optional[datetime.datetime] = None
        self.audio_path: Optional[str] = None
        self.speakers: Dict[str, Speaker] = {}
        self.segments: List[Segment] = []
        self.summary: Optional[str] = None
        self.action_items: Optional[str] = None
        self.requirements: Optional[str] = None
        self.template: str = "General"

    def get_or_create_speaker(self, speaker_id: str) -> Speaker:
        if speaker_id not in self.speakers:
            self.speakers[speaker_id] = Speaker(speaker_id=speaker_id)
        return self.speakers[speaker_id]

    def rename_speaker(self, speaker_id: str, name: str) -> None:
        if speaker_id in self.speakers:
            self.speakers[speaker_id].display_name = name

    def full_transcript(self) -> str:
        if not self.segments:
            return ""
        lines = []
        for seg in self.segments:
            speaker = self.speakers.get(seg.speaker_id)
            name = speaker.display_name if speaker else seg.speaker_id
            start = _fmt_time(seg.start)
            end = _fmt_time(seg.end)
            lines.append(f"[{start} → {end}] {name}: {seg.text}")
        return "\n".join(lines)

    def to_dict(self) -> dict:
        return {
            "session_id": self.session_id,
            "display_name": self.display_name,
            "started_at": self.started_at.isoformat() if self.started_at else None,
            "ended_at": self.ended_at.isoformat() if self.ended_at else None,
            "audio_path": self.audio_path,
            "speakers": {k: v.to_dict() for k, v in self.speakers.items()},
            "segments": [s.to_dict() for s in self.segments],
            "summary": self.summary,
            "action_items": self.action_items,
            "requirements": self.requirements,
            "template": self.template,
        }


def _fmt_time(seconds: float) -> str:
    m, s = divmod(int(seconds), 60)
    return f"{m:02d}:{s:02d}"
