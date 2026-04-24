"""
Client configuration persistence.

Stores a `clients.json` in the recordings folder mapping client name to a
designated output folder. When a session is tagged with a client that has a
folder configured, the app routes its exports (transcript, summary,
action items, decisions, requirements) to that folder instead of the
default recordings directory.

Schema (clients.json):
    [
      {"name": "Acme Corp", "folder": "C:\\customers\\acme"},
      {"name": "TechCo",    "folder": ""}
    ]

An empty folder string means "use the default recordings directory".
"""

import json
import os
import tempfile
from pathlib import Path
from typing import Dict, List, Optional

from utils.logger import get_logger

logger = get_logger(__name__)

CLIENTS_FILENAME = "clients.json"


class ClientService:
    """Load/save per-client output folder mappings."""

    def __init__(self, recordings_dir: str):
        self._recordings_dir = Path(recordings_dir)
        self._recordings_dir.mkdir(parents=True, exist_ok=True)
        self._path = self._recordings_dir / CLIENTS_FILENAME

    def load(self) -> List[Dict[str, str]]:
        """Return the current client list. Missing file → empty list."""
        if not self._path.exists():
            return []
        try:
            with open(self._path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, list):
                logger.warning(f"{self._path.name} is not a list; ignoring.")
                return []
            cleaned: List[Dict[str, str]] = []
            for entry in data:
                if not isinstance(entry, dict):
                    continue
                name = str(entry.get("name", "") or "").strip()
                folder = str(entry.get("folder", "") or "").strip()
                if name:
                    cleaned.append({"name": name, "folder": folder})
            return cleaned
        except (json.JSONDecodeError, OSError) as e:
            logger.warning(f"Could not read {self._path}: {e}")
            return []

    def save(self, clients: List[Dict[str, str]]) -> None:
        """Atomically persist the client list."""
        payload = []
        seen_names = set()
        for entry in clients:
            name = str(entry.get("name", "") or "").strip()
            folder = str(entry.get("folder", "") or "").strip()
            if not name or name.lower() in seen_names:
                continue
            seen_names.add(name.lower())
            payload.append({"name": name, "folder": folder})

        data = json.dumps(payload, indent=2, ensure_ascii=False)
        fd, tmp_path = tempfile.mkstemp(
            dir=self._recordings_dir, suffix=".json.tmp")
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(data)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_path, self._path)
        except Exception:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            raise
        logger.info(f"Saved {len(payload)} client(s) to {self._path}")

    def folder_for(self, client_name: str) -> str:
        """Return the configured folder for a client, or '' if none.

        Empty string = use the default recordings directory.
        """
        if not client_name or not client_name.strip():
            return ""
        target = client_name.strip().lower()
        for entry in self.load():
            if entry["name"].strip().lower() == target:
                return entry["folder"]
        return ""

    def upsert(self, name: str, folder: str) -> None:
        """Add or update a client entry."""
        name = name.strip()
        if not name:
            return
        clients = self.load()
        target = name.lower()
        for entry in clients:
            if entry["name"].strip().lower() == target:
                entry["folder"] = folder.strip()
                self.save(clients)
                return
        clients.append({"name": name, "folder": folder.strip()})
        self.save(clients)

    def remove(self, name: str) -> None:
        """Remove a client entry by name (case-insensitive)."""
        target = name.strip().lower()
        clients = [c for c in self.load()
                   if c["name"].strip().lower() != target]
        self.save(clients)


def resolve_export_dir(
    client_folder: str,
    default_dir: str,
) -> str:
    """
    Given a client's configured folder and the default recordings dir,
    return the folder where exports should be written. Creates the
    client folder on demand. Falls back to default_dir on any error.
    """
    if not client_folder or not client_folder.strip():
        return default_dir
    target = client_folder.strip()
    try:
        Path(target).mkdir(parents=True, exist_ok=True)
        return target
    except OSError as e:
        logger.warning(
            f"Could not use client folder '{target}' ({e}); "
            f"falling back to default")
        return default_dir
