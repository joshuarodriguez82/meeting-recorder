"""
Application configuration loaded from environment variables.
All secrets are sourced from .env — never hardcoded.
"""

import os
from pathlib import Path
from dataclasses import dataclass
from dotenv import load_dotenv

load_dotenv()

ENV_PATH = Path(__file__).resolve().parent.parent / ".env"


@dataclass(frozen=True)
class Settings:
    """Immutable application settings resolved at startup."""

    anthropic_api_key: str
    hf_token: str
    whisper_model: str
    max_speakers: int
    recordings_dir: str
    email_to: str

    @classmethod
    def from_env(cls) -> "Settings":
        """
        Load settings from environment variables.
        Missing API keys default to empty strings — the user can set them
        via the in-app Settings dialog without blocking app startup.
        """
        return cls(
            anthropic_api_key=os.getenv("ANTHROPIC_API_KEY", ""),
            hf_token=os.getenv("HF_TOKEN", ""),
            whisper_model=os.getenv("WHISPER_MODEL", "base"),
            max_speakers=int(os.getenv("MAX_SPEAKERS", "10")),
            recordings_dir=os.getenv("RECORDINGS_DIR", "recordings"),
            email_to=os.getenv("EMAIL_TO", ""),
        )

    @property
    def is_configured(self) -> bool:
        """True if both required API keys are set."""
        return bool(self.anthropic_api_key) and bool(self.hf_token)

    @staticmethod
    def save_to_env(
        anthropic_api_key: str,
        hf_token: str,
        whisper_model: str,
        max_speakers: int,
        recordings_dir: str,
        email_to: str = "",
    ) -> None:
        """Write settings back to the .env file."""
        content = (
            f"ANTHROPIC_API_KEY={anthropic_api_key}\n"
            f"HF_TOKEN={hf_token}\n"
            f"WHISPER_MODEL={whisper_model}\n"
            f"MAX_SPEAKERS={max_speakers}\n"
            f"RECORDINGS_DIR={recordings_dir}\n"
            f"EMAIL_TO={email_to}\n"
        )
        ENV_PATH.write_text(content, encoding="utf-8")