"""Shared, file-backed store for stage-2 per-person weekly updates.

Keyed by (team, project, start_date, end_date) so anyone who opens the same
team/project for the same week sees everyone else's already-saved (locked)
entries. Entries are only ever appended once a person ticks/saves their own
update - there is no partial/draft state on the server."""

import json
import os
import threading
from datetime import datetime, timezone

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(os.path.dirname(_THIS_DIR))
DATA_DIR = os.path.join(BASE_DIR, "data")
STORE_PATH = os.path.join(DATA_DIR, "weekly_updates.json")

_lock = threading.Lock()


def _key(team: str, project: str, start_date: str, end_date: str) -> str:
    return "||".join([team, project, start_date, end_date])


def _load() -> dict:
    if not os.path.exists(STORE_PATH):
        return {}
    with open(STORE_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _save(data: dict) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp_path = STORE_PATH + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp_path, STORE_PATH)


def get_entries(team: str, project: str, start_date: str, end_date: str) -> list:
    with _lock:
        data = _load()
        return data.get(_key(team, project, start_date, end_date), [])


def add_entry(team: str, project: str, start_date: str, end_date: str, name: str, update: str) -> list:
    with _lock:
        data = _load()
        key = _key(team, project, start_date, end_date)
        entries = data.setdefault(key, [])
        entries.append({
            "name": name,
            "update": update,
            "locked": True,
            "saved_at": datetime.now(timezone.utc).isoformat(),
        })
        _save(data)
        return entries


def remove_entry(team: str, project: str, start_date: str, end_date: str, index: int) -> list:
    with _lock:
        data = _load()
        key = _key(team, project, start_date, end_date)
        entries = data.get(key, [])
        if 0 <= index < len(entries):
            entries.pop(index)
            data[key] = entries
            _save(data)
        return entries
