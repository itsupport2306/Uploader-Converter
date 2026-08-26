"""Neon Postgres writes for extracted PDF profiles.

The profile_id is the identifier: it is derived from the PDF's SHA-256 so the
same source PDF always maps to the same row, which makes reruns idempotent.
"""
from __future__ import annotations

import json
import os
import re
import uuid
from datetime import datetime, timezone

from . import config

PROFILE_NAMESPACE = uuid.UUID("6f9b1a6e-6f0e-5f6b-9a0e-2f1c3d4e5a6b")

REQUIRED_COLUMNS = {
    "profile_id", "first_name", "last_name", "bio", "specialty",
    "years_experience", "city", "state_code", "resume_url", "resume_sections",
    "completion_score", "search_text", "source", "created_at", "updated_at",
}


def _import_sqlalchemy():
    try:
        import sqlalchemy  # noqa: PLC0415
    except ImportError as exc:  # pragma: no cover - dependency guard
        raise SystemExit("ERROR: SQLAlchemy is required. Run: pip install -r requirements.txt") from exc
    return sqlalchemy


def text_sql(sql: str):
    return _import_sqlalchemy().text(sql)


def create_engine():
    sqlalchemy = _import_sqlalchemy()
    db_url = config.normalize_db_url(os.environ.get("DATABASE_URL", ""))
    if not db_url:
        raise SystemExit("ERROR: DATABASE_URL is missing.")
    if not db_url.startswith("postgresql"):
        raise SystemExit("ERROR: DATABASE_URL must point to Neon Postgres.")
    return sqlalchemy.create_engine(db_url, future=True, pool_pre_ping=True)


def quoted_table(name: str) -> str:
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
        raise SystemExit(f"ERROR: unsafe table name: {name!r}")
    return f'"{name}"'


def utcnow() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def ensure_table(engine, table_name: str) -> None:
    with engine.connect() as conn:
        exists = conn.execute(
            text_sql("SELECT to_regclass(:qualified)"),
            {"qualified": f"public.{table_name}"},
        ).scalar()
        if exists is None:
            raise SystemExit(f"ERROR: Neon database is missing public.{table_name}.")
        columns = set(conn.execute(text_sql("""
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = :table_name
        """), {"table_name": table_name}).scalars().all())
    missing = sorted(REQUIRED_COLUMNS - columns)
    if missing:
        raise SystemExit(f"ERROR: {table_name} is missing column(s): " + ", ".join(missing))


def profile_id_for(digest: str) -> str:
    """Deterministic profile ID derived from the PDF content hash."""
    return str(uuid.uuid5(PROFILE_NAMESPACE, f"pdf-profile:{digest}"))


def find_existing_profile_id(conn, table_name: str, *, profile_id: str, digest: str, fields: dict) -> str | None:
    """Match on the generated ID, then the PDF hash, then name + location."""
    table = quoted_table(table_name)
    row = conn.execute(text_sql(f"""
        SELECT profile_id
        FROM {table}
        WHERE profile_id = CAST(:profile_id AS TEXT)
           OR resume_sections::text LIKE :digest_like
           OR (
                CAST(:first_name AS TEXT) IS NOT NULL
                AND CAST(:last_name AS TEXT) IS NOT NULL
                AND lower(first_name) = lower(CAST(:first_name AS TEXT))
                AND lower(last_name) = lower(CAST(:last_name AS TEXT))
                AND (
                     CAST(:city AS TEXT) IS NULL
                     OR lower(coalesce(city, '')) = lower(CAST(:city AS TEXT))
                )
           )
        LIMIT 1
    """), {
        "profile_id": profile_id,
        "digest_like": f"%{digest}%",
        "first_name": fields.get("first_name"),
        "last_name": fields.get("last_name"),
        "city": fields.get("city"),
    }).first()
    return str(row[0]) if row else None


def completion_score(fields: dict) -> int:
    score = 20  # a resume PDF is always attached
    if fields.get("bio"):
        score += 20
    if fields.get("specialty"):
        score += 20
    if fields.get("years_experience"):
        score += 15
    if fields.get("city") and fields.get("state_code"):
        score += 15
    if fields.get("first_name") and fields.get("last_name"):
        score += 10
    return min(score, 100)


def search_text(fields: dict) -> str:
    parts = [
        fields.get("first_name"), fields.get("last_name"), fields.get("full_name"),
        fields.get("specialty"), fields.get("city"), fields.get("state_code"),
        fields.get("bio"),
    ]
    return " ".join(str(part) for part in parts if part).lower()[:8000]


def upsert_profile(
    conn,
    *,
    table_name: str,
    profile_id: str,
    existing_id: str | None,
    fields: dict,
    resume_sections: dict,
    resume_url: str,
) -> tuple[str, str]:
    """Insert or update one profile row. Returns (profile_id, 'inserted'|'updated')."""
    table = quoted_table(table_name)
    now = utcnow()
    row_id = existing_id or profile_id
    params = {
        "profile_id": row_id,
        "first_name": fields.get("first_name"),
        "last_name": fields.get("last_name"),
        "headline": fields.get("specialty"),
        "bio": fields.get("bio"),
        "specialty": fields.get("specialty"),
        "years_experience": int(fields.get("years_experience") or 0),
        "city": fields.get("city"),
        "state_code": fields.get("state_code"),
        "resume_url": resume_url,
        "resume_sections": json.dumps(resume_sections, default=str),
        "completion_score": completion_score(fields),
        "search_text": search_text(fields),
        "source": "resume_parse",
        "capture_source": "pdf_profile_pipeline",
        "created_at": now,
        "updated_at": now,
        "captured_at": datetime.now(timezone.utc),
    }

    if existing_id:
        conn.execute(text_sql(f"""
            UPDATE {table}
            SET first_name = :first_name,
                last_name = :last_name,
                headline = :headline,
                bio = :bio,
                specialty = :specialty,
                years_experience = :years_experience,
                city = :city,
                state_code = :state_code,
                resume_url = :resume_url,
                resume_sections = CAST(:resume_sections AS JSON),
                completion_score = :completion_score,
                search_text = :search_text,
                capture_source = :capture_source,
                captured_at = :captured_at,
                updated_at = :updated_at
            WHERE profile_id = :profile_id
        """), params)
        return row_id, "updated"

    inserted = conn.execute(text_sql(f"""
        INSERT INTO {table} (
            profile_id, first_name, last_name, headline, bio, specialty,
            years_experience, city, state_code, open_to_work, resume_url,
            resume_sections, completion_score, search_text, source,
            capture_source, captured_at, created_at, updated_at
        )
        VALUES (
            :profile_id, :first_name, :last_name, :headline, :bio, :specialty,
            :years_experience, :city, :state_code, TRUE, :resume_url,
            CAST(:resume_sections AS JSON), :completion_score, :search_text,
            CAST(:source AS profilesource),
            :capture_source, :captured_at, :created_at, :updated_at
        )
        RETURNING profile_id
    """), params).first()
    return str(inserted[0]), "inserted"
