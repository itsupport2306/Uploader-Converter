"""Logging for every DOCX file the pipeline touches.

Three sinks, all optional but all on by default:
  * JSONL manifest  - one line per record, used for resume/retry on rerun
  * CSV upload log  - one row per upload attempt, easy to open in Excel
  * run log file    - the full console transcript of the run
"""
from __future__ import annotations

import csv
import json
import sys
import threading
from datetime import datetime, timezone
from pathlib import Path

CSV_COLUMNS = [
    "logged_at", "status", "profile_id", "db_action", "first_name", "last_name",
    "full_name", "city", "state_code", "specialty", "years_experience",
    "years_experience_source", "docx_sha256", "docx_sections", "cloudflare_key",
    "cloudflare_uploaded", "resume_url", "source_file", "extraction_mode", "error",
]

_lock = threading.Lock()


def utc_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def optional_path(value: str | None, default: Path | None) -> Path | None:
    """Resolve a CLI path argument; 'none' disables the sink."""
    if value is None:
        return default
    normalized = value.strip()
    if not normalized or normalized.lower() in {"none", "null", "off", "false", "-"}:
        return None
    return Path(normalized).expanduser()


def append_manifest(path: Path, record: dict) -> None:
    with _lock:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(record, sort_keys=True, default=str) + "\n")


def append_csv(path: Path, record: dict) -> None:
    with _lock:
        path.parent.mkdir(parents=True, exist_ok=True)
        write_header = not path.exists() or path.stat().st_size == 0
        with path.open("a", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=CSV_COLUMNS, extrasaction="ignore")
            if write_header:
                writer.writeheader()
            writer.writerow({column: record.get(column) for column in CSV_COLUMNS})


def latest_status(path: Path | None) -> dict[str, dict]:
    """Last manifest entry per record_id, for skip-on-rerun."""
    latest: dict[str, dict] = {}
    if path is None or not path.exists():
        return latest
    with path.open("r", encoding="utf-8") as fh:
        for line in fh:
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                continue
            record_id = record.get("record_id")
            if record_id:
                latest[record_id] = record
    return latest


class Tee:
    """Mirror stdout/stderr into a run log file."""

    def __init__(self, stream, handle):
        self._stream = stream
        self._handle = handle

    def write(self, data):
        self._stream.write(data)
        try:
            self._handle.write(data)
            self._handle.flush()
        except Exception:
            pass
        return len(data)

    def flush(self):
        self._stream.flush()
        try:
            self._handle.flush()
        except Exception:
            pass

    def isatty(self):
        return getattr(self._stream, "isatty", lambda: False)()


class RunLog:
    """Context manager that tees console output to a timestamped log file."""

    def __init__(self, path: Path | None):
        self.path = path
        self._handle = None
        self._stdout = None
        self._stderr = None

    def __enter__(self) -> "RunLog":
        if self.path is None:
            return self
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._handle = self.path.open("a", encoding="utf-8")
        self._handle.write(f"\n===== run started {utc_iso()} =====\n")
        self._stdout, self._stderr = sys.stdout, sys.stderr
        sys.stdout = Tee(self._stdout, self._handle)
        sys.stderr = Tee(self._stderr, self._handle)
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        if self._handle is None:
            return False
        sys.stdout, sys.stderr = self._stdout, self._stderr
        self._handle.write(f"===== run finished {utc_iso()} =====\n")
        self._handle.close()
        self._handle = None
        return False
