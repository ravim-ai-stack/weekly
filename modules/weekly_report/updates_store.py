"""Shared store for stage-2 per-person weekly updates, persisted in Vercel
Blob (see blob_store.py) rather than local disk - Vercel's deployment
filesystem is read-only, and even /tmp is wiped between deploys and isn't
shared across serverless instances, so anything written there could vanish
or be invisible to the very next request.

Keyed by (team, project, start_date, end_date) so anyone who opens the same
team/project for the same week sees everyone else's already-saved (locked)
entries. Entries are only ever appended once a person ticks/saves their own
update - there is no partial/draft state on the server."""

import json
import threading
from datetime import datetime, timezone

from . import blob_store

STORE_PATHNAME = "data/weekly_updates.json"

_lock = threading.Lock()


def _key(team: str, project: str, start_date: str, end_date: str) -> str:
    return "||".join([team, project, start_date, end_date])


def _load() -> dict:
    raw = blob_store.get(STORE_PATHNAME)
    return json.loads(raw) if raw else {}


def _save(data: dict) -> None:
    blob_store.put(
        STORE_PATHNAME,
        json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8"),
        content_type="application/json",
    )


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
