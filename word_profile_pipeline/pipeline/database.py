"""Neon Postgres writes for extracted DOCX profiles.

The profile_id is the identifier: it is derived from the DOCX SHA-256 so the
same source DOCX always maps to the same row, which makes reruns idempotent.
"""
from __future__ import annotations

import json
import os
import re
import uuid
from datetime import datetime, timezone

from . import config
from .schema import PROFILE_LIMITS, WORK_HISTORY_LIMITS, truncate
#git issue

PROFILE_NAMESPACE = uuid.UUID("6f9b1a6e-6f0e-5f6b-9a0e-2f1c3d4e5a6b")

REQUIRED_COLUMNS = {
    "profile_id", "first_name", "last_name", "headline", "bio", "specialty",
    "profession_type", "years_experience", "email", "phone", "city", "state_code",
    "zip_code", "work_authorization", "education", "resume_url", "resume_sections",
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
    # One pooled connection per worker, plus headroom. Each worker opens its own
    # `engine.begin()` transaction, so a pool smaller than the worker count would
    # serialise them behind a checkout wait instead of running concurrently.
    workers = max(1, config.env_int("PIPELINE_WORKERS", 2))
    return sqlalchemy.create_engine(
        db_url,
        future=True,
        pool_pre_ping=True,
        pool_size=workers + 1,
        max_overflow=workers,
        pool_recycle=300,
    )


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
    """Deterministic profile ID derived from the DOCX content hash."""
    return str(uuid.uuid5(PROFILE_NAMESPACE, f"word-profile:{digest}"))


def find_existing_profile_id(conn, table_name: str, *, profile_id: str, digest: str, fields: dict) -> str | None:
    """Find the row this resume belongs to, so a rerun updates instead of duplicating.

    Checked in order of confidence, each against an index:
      1. the deterministic profile_id, which is itself derived from the DOCX hash,
         so this already covers "the same DOCX was seen before";
      2. email, which has a unique index and would otherwise raise on insert;
      3. first + last name with the same city.
    """
    table = quoted_table(table_name)

    row = conn.execute(
        text_sql(f"SELECT profile_id FROM {table} WHERE profile_id = :profile_id LIMIT 1"),
        {"profile_id": profile_id},
    ).first()
    if row:
        return str(row[0])

    email = (fields.get("email") or "").strip().lower()
    if email:
        row = conn.execute(
            text_sql(f"SELECT profile_id FROM {table} WHERE lower(email) = :email LIMIT 1"),
            {"email": email},
        ).first()
        if row:
            return str(row[0])

    first, last, city = fields.get("first_name"), fields.get("last_name"), fields.get("city")
    if first and last and city:
        row = conn.execute(text_sql(f"""
            SELECT profile_id FROM {table}
            WHERE lower(first_name) = lower(:first_name)
              AND lower(last_name) = lower(:last_name)
              AND lower(coalesce(city, '')) = lower(:city)
            LIMIT 1
        """), {"first_name": first, "last_name": last, "city": city}).first()
        if row:
            return str(row[0])
    return None


def completion_score(fields: dict) -> int:
    score = 20  # a resume DOCX is always attached
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
        fields.get("headline"), fields.get("specialty"), fields.get("profession_type"),
        fields.get("city"), fields.get("state_code"), fields.get("bio"),
    ]
    parts += [position.get("employer") for position in (fields.get("positions") or [])]
    parts += [position.get("title") for position in (fields.get("positions") or [])]
    parts += list(fields.get("certifications") or [])
    parts += list(fields.get("skills") or [])
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
    """Insert or update one profile row. Returns (profile_id, 'inserted'|'updated').

    An update never destroys data: every optional column is written with
    COALESCE(NULLIF(new, ''), existing), so a weaker rerun -- a failed model
    call, a worse OCR pass -- can only add to a row, never blank it out.
    """
    table = quoted_table(table_name)
    now = utcnow()
    row_id = existing_id or profile_id
    # Every varchar column is clamped here as well as in schema.validate_profile.
    # An over-long headline used to abort the whole INSERT, which surfaced as
    # "extraction failed" rather than as the one-column problem it was.
    fields = {
        key: (truncate(value, PROFILE_LIMITS[key]) if key in PROFILE_LIMITS and
              isinstance(value, str) else value)
        for key, value in fields.items()
    }
    params = {
        "profile_id": row_id,
        "first_name": fields.get("first_name"),
        "last_name": fields.get("last_name"),
        "headline": truncate(
            fields.get("headline") or fields.get("specialty"), PROFILE_LIMITS["headline"]),
        "bio": fields.get("bio"),
        "specialty": fields.get("specialty"),
        "profession_type": fields.get("profession_type"),
        # smallint column: a nonsense value from a bad parse must not abort the row.
        "years_experience": max(0, min(int(fields.get("years_experience") or 0), 99)),
        "email": fields.get("email"),
        "phone": fields.get("phone"),
        "city": fields.get("city"),
        "state_code": fields.get("state_code"),
        "zip_code": fields.get("zip_code"),
        "work_authorization": fields.get("work_authorization"),
        "education": json.dumps(fields.get("education") or [], default=str),
        "resume_url": resume_url,
        "resume_sections": json.dumps(resume_sections, default=str),
        "completion_score": completion_score(fields),
        "search_text": search_text(fields),
        "source": "resume_parse",
        "capture_source": "word_profile_pipeline",
        "created_at": now,
        "updated_at": now,
        "captured_at": datetime.now(timezone.utc),
    }

    if existing_id:
        conn.execute(text_sql(f"""
            UPDATE {table}
            SET first_name = coalesce(nullif(:first_name, ''), first_name),
                last_name = coalesce(nullif(:last_name, ''), last_name),
                headline = coalesce(nullif(:headline, ''), headline),
                -- The longer summary wins, so a truncated rerun cannot shrink it.
                bio = CASE
                        WHEN coalesce(CAST(:bio AS TEXT), '') = '' THEN bio
                        WHEN bio IS NULL OR length(bio) <= length(CAST(:bio AS TEXT))
                            THEN CAST(:bio AS TEXT)
                        ELSE bio
                      END,
                specialty = coalesce(nullif(:specialty, ''), specialty),
                profession_type = coalesce(nullif(:profession_type, ''), profession_type),
                -- 0 means "not found this time", not "zero years".
                years_experience = CASE
                        WHEN CAST(:years_experience AS SMALLINT) > 0
                            THEN CAST(:years_experience AS SMALLINT)
                        ELSE years_experience
                      END,
                email = coalesce(nullif(:email, ''), email),
                phone = coalesce(nullif(:phone, ''), phone),
                city = coalesce(nullif(:city, ''), city),
                state_code = coalesce(nullif(:state_code, ''), state_code),
                zip_code = coalesce(nullif(:zip_code, ''), zip_code),
                work_authorization = coalesce(nullif(:work_authorization, ''), work_authorization),
                education = CASE
                        WHEN CAST(:education AS TEXT) IN ('[]', 'null') THEN education
                        ELSE CAST(:education AS JSON)
                      END,
                resume_url = coalesce(nullif(:resume_url, ''), resume_url),
                resume_sections = CAST(:resume_sections AS JSON),
                completion_score = greatest(
                    coalesce(completion_score, 0), CAST(:completion_score AS SMALLINT)),
                search_text = coalesce(nullif(:search_text, ''), search_text),
                capture_source = :capture_source,
                captured_at = :captured_at,
                updated_at = :updated_at
            WHERE profile_id = :profile_id
        """), params)
        return row_id, "updated"

    inserted = conn.execute(text_sql(f"""
        INSERT INTO {table} (
            profile_id, first_name, last_name, headline, bio, specialty,
            profession_type, years_experience, email, phone, city, state_code,
            zip_code, work_authorization, education, open_to_work, job_type_prefs,
            resume_url,
            resume_sections, completion_score, search_text, source,
            capture_source, captured_at, created_at, updated_at
        )
        VALUES (
            :profile_id, :first_name, :last_name, :headline, :bio, :specialty,
            :profession_type, :years_experience, nullif(:email, ''), :phone,
            :city, :state_code, :zip_code, :work_authorization,
            -- job_type_prefs is NOT NULL with no default, and a resume never
            -- states one, so it starts empty for the candidate to fill in.
            CAST(:education AS JSON), TRUE, CAST('[]' AS JSON), :resume_url,
            CAST(:resume_sections AS JSON), :completion_score, :search_text,
            CAST(:source AS profilesource),
            :capture_source, :captured_at, :created_at, :updated_at
        )
        RETURNING profile_id
    """), params).first()
    return str(inserted[0]), "inserted"


# ---------------------------------------------------------------------------
# work_history

WORK_HISTORY_NAMESPACE = uuid.UUID("2c8f4b1d-7a3e-5c92-b0d6-1e4f8a7c3b52")

_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}


def parse_month_year(raw) -> "date | None":
    """'August 2025', '2025-08', '2013' -> a date. None when undatable.

    A year-only value becomes 1 January of that year; the resume simply does not
    say more than that, and inventing a month would be worse than a flat date.
    """
    from datetime import date  # noqa: PLC0415

    text = str(raw or "").strip()
    if not text or text.lower() in {"present", "current", "now", "none", "null"}:
        return None
    year_match = re.search(r"(19|20)\d{2}", text)
    if not year_match:
        return None
    year = int(year_match.group(0))

    month = 1
    iso = re.match(r"^\s*(19|20)\d{2}[-/](\d{1,2})", text)
    if iso:
        month = max(1, min(12, int(iso.group(2))))
    else:
        name = re.search(r"[A-Za-z]{3,}", text)
        if name:
            month = _MONTHS.get(name.group(0)[:4].lower().rstrip("t"), _MONTHS.get(
                name.group(0)[:3].lower(), 1))
    try:
        return date(year, month, 1)
    except ValueError:
        return None


def work_id_for(profile_id: str, position: dict) -> str:
    """Stable ID per (profile, employer, title, start), so reruns do not duplicate."""
    key = "|".join(
        re.sub(r"[^a-z0-9]+", "", str(position.get(field) or "").lower())
        for field in ("employer", "title", "start_date")
    )
    return str(uuid.uuid5(WORK_HISTORY_NAMESPACE, f"{profile_id}:{key}"))


def sync_work_history(conn, *, profile_id: str, positions: list[dict]) -> int:
    """Insert this profile's positions, skipping ones already stored.

    Existing rows are left alone rather than rewritten: a recruiter may have
    corrected them by hand, and a rerun of the parser should not undo that.
    """
    if not positions:
        return 0
    if conn.execute(text_sql("SELECT to_regclass('public.work_history')")).scalar() is None:
        return 0

    existing = set(conn.execute(
        text_sql("SELECT work_id FROM work_history WHERE profile_id = :profile_id"),
        {"profile_id": profile_id},
    ).scalars().all())

    now = utcnow()
    added = 0
    seen: set[str] = set()
    for position in positions:
        employer = (position.get("employer") or "").strip()
        title = (position.get("title") or "").strip()
        # Both columns are NOT NULL, so a half-parsed entry cannot be stored.
        if not employer or not title:
            continue
        work_id = work_id_for(profile_id, position)
        if work_id in existing or work_id in seen:
            continue
        seen.add(work_id)
        conn.execute(text_sql("""
            INSERT INTO work_history (
                work_id, profile_id, employer_name, job_title, specialty,
                start_date, end_date, city, state_code, description, created_at
            )
            VALUES (
                :work_id, :profile_id, :employer_name, :job_title, :specialty,
                :start_date, :end_date, :city, :state_code, :description, :created_at
            )
            ON CONFLICT (work_id) DO NOTHING
        """), {
            "work_id": work_id,
            "profile_id": profile_id,
            # These columns are varchar(200)/(120)/(100)/(2) - not 255.
            "employer_name": truncate(employer, WORK_HISTORY_LIMITS["employer"]),
            "job_title": truncate(title, WORK_HISTORY_LIMITS["title"]),
            "specialty": truncate(
                position.get("specialty") or None, WORK_HISTORY_LIMITS["specialty"]),
            "start_date": parse_month_year(position.get("start_date")),
            "end_date": parse_month_year(position.get("end_date")),
            "city": truncate(position.get("city") or None, WORK_HISTORY_LIMITS["city"]),
            "state_code": truncate(
                position.get("state_code") or None, WORK_HISTORY_LIMITS["state_code"]),
            "description": (position.get("description") or None),
            "created_at": now,
        })
        added += 1
    return added
