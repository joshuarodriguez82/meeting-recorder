"""
Application configuration loaded from environment variables.
All secrets are sourced from .env — never hardcoded.
"""

import os
from dataclasses import dataclass
from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class Settings:
    """Immutable application settings resolved at startup."""

    anthropic_api_key: str
    hf_token: str
    whisper_model: str
    max_speakers: int
    recordings_dir: str

    @classmethod
    def from_env(cls) -> "Settings":
        """
        Load settings from environment variables.

        Raises:
            EnvironmentError: If a required variable is missing.
        """
        required = {
            "ANTHROPIC_API_KEY": os.getenv("ANTHROPIC_API_KEY"),
            "HF_TOKEN": os.getenv("HF_TOKEN"),
        }

        missing = [k for k, v in required.items() if not v]
        if missing:
            raise EnvironmentError(
                f"Missing required environment variables: {', '.join(missing)}\n"
                "Copy .env.example to .env and fill in your keys."
            )

        return cls(
            anthropic_api_key=required["ANTHROPIC_API_KEY"],
            hf_token=required["HF_TOKEN"],
            whisper_model=os.getenv("WHISPER_MODEL", "base"),
            max_speakers=int(os.getenv("MAX_SPEAKERS", "10")),
            recordings_dir=os.getenv("RECORDINGS_DIR", "recordings"),
        )