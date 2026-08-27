"""Environment loading and runtime settings for the PDF profile pipeline.

Everything configurable lives in a .env file; nothing is hardcoded here.
"""
from __future__ import annotations

import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_ENV_FILE = ".env"

# (key, human label) pairs that must be present before any non-dry-run work.
REQUIRED_DB_ENV = [
    ("DATABASE_URL", "Neon Postgres connection string"),
]
REQUIRED_STORAGE_ENV = [
    ("S3_ENDPOINT_URL", "Cloudflare R2 endpoint URL"),
    ("S3_BUCKET", "Cloudflare R2 bucket"),
    ("S3_ACCESS_KEY", "Cloudflare R2 access key"),
    ("S3_SECRET_KEY", "Cloudflare R2 secret key"),
]
REQUIRED_LLM_ENV = [
    ("LLM_BASE_URL", "OpenAI-compatible base URL for the Qwen2.5 model"),
    ("LLM_MODEL", "Qwen2.5 model name"),
]

DEFAULT_ENV_VALUES = {
    "S3_REGION": "auto",
    "S3_ACL": "",
    "STORAGE_ENABLED": "true",
    "LLM_ENABLED": "true",
    "LLM_API_KEY": "ollama",
    "LLM_TIMEOUT_SECONDS": "180",
    "LLM_MAX_INPUT_CHARS": "18000",
    "LLM_MAX_TOKENS": "4096",
    "LLM_TEMPERATURE": "0",
    "R2_KEY_PREFIX": "resumes",
    "TARGET_TABLE": "profiles",
    "ONEDRIVE_TENANT": "common",
}


def env_path(path: str | os.PathLike[str]) -> Path:
    candidate = Path(path).expanduser()
    return candidate if candidate.is_absolute() else (ROOT / candidate)


def load_env(path: str | os.PathLike[str] = DEFAULT_ENV_FILE) -> Path:
    """Load KEY=VALUE lines from an env file without overriding real env vars."""
    resolved = env_path(path)
    if resolved.exists():
        with open(resolved, "r", encoding="utf-8") as fh:
            for raw_line in fh:
                line = raw_line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key:
                    os.environ.setdefault(key, value)
    for key, value in DEFAULT_ENV_VALUES.items():
        os.environ.setdefault(key, value)
    return resolved


def env_bool(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None or value.strip() == "":
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def env_int(name: str, default: int) -> int:
    try:
        return int(str(os.environ.get(name, "")).strip())
    except (TypeError, ValueError):
        return default


def env_float(name: str, default: float) -> float:
    try:
        return float(str(os.environ.get(name, "")).strip())
    except (TypeError, ValueError):
        return default


def normalize_db_url(url: str) -> str:
    url = (url or "").strip()
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    if url.startswith("postgresql://") and "+psycopg" not in url:
        url = url.replace("postgresql://", "postgresql+psycopg://", 1)
    return url


def safe_db_label(url: str) -> str:
    """Host/database only, so connection strings never reach the logs."""
    from urllib.parse import urlparse

    try:
        parsed = urlparse(url)
        return f"{parsed.hostname or '?'}{parsed.path or ''}"
    except ValueError:
        return "(unparseable DATABASE_URL)"


def validate_env(*, dry_run: bool, use_llm: bool) -> None:
    if dry_run and not use_llm:
        return

    missing: list[str] = []
    if use_llm:
        missing += [key for key, _label in REQUIRED_LLM_ENV if not os.environ.get(key)]
    if not dry_run:
        missing += [key for key, _label in REQUIRED_DB_ENV if not os.environ.get(key)]
        if not env_bool("STORAGE_ENABLED", True):
            raise SystemExit("ERROR: STORAGE_ENABLED must be true to upload PDFs to Cloudflare R2.")
        missing += [key for key, _label in REQUIRED_STORAGE_ENV if not os.environ.get(key)]
    if missing:
        raise SystemExit("ERROR: missing required .env setting(s): " + ", ".join(sorted(set(missing))))

    if not dry_run:
        db_url = normalize_db_url(os.environ.get("DATABASE_URL", ""))
        if not db_url.startswith("postgresql"):
            raise SystemExit("ERROR: DATABASE_URL must point to Neon Postgres.")
