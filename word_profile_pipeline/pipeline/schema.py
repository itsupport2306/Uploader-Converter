"""Validate, ground and clamp extracted fields before they reach Postgres.

Three jobs, in this order:

1. **Coerce** - a small model happily returns a list where a string belongs, a
   dict where a list belongs, or the string "null". Every value is forced into
   the shape the rest of the pipeline expects, or dropped.
2. **Ground** - a value that does not appear in the resume text was invented by
   the model, and an invented specialty or employer is worse than a NULL. Text
   fields are checked against the source and nulled when they are not supported.
3. **Clamp** - the profiles/work_history columns are varchar(N). An over-long
   headline used to abort the whole INSERT, which read as "extraction failed".

Nothing here ever fills a value in; it only removes what cannot be trusted.
"""
from __future__ import annotations

import re

# varchar(N) limits, read from the live Neon schema. A value longer than its
# column is truncated on a word boundary rather than failing the insert.
PROFILE_LIMITS = {
    "first_name": 100,
    "last_name": 100,
    "headline": 255,
    "specialty": 100,
    "profession_type": 50,
    "phone": 30,
    "email": 255,
    "city": 120,
    "state_code": 2,
    "zip_code": 10,
    "work_authorization": 80,
}

WORK_HISTORY_LIMITS = {
    "employer": 200,
    "title": 200,
    "specialty": 100,
    "city": 120,
    "state_code": 2,
}

# Fields that must be a plain string when present.
PROFILE_STRING_FIELDS = (
    "first_name", "last_name", "full_name", "email", "phone", "city",
    "state_code", "zip_code", "headline", "specialty", "current_employer",
    "profession_type", "work_authorization", "bio",
)

POSITION_STRING_FIELDS = (
    "title", "employer", "city", "state_code", "start_date", "end_date", "description",
)

_NULLISH = {
    "", "null", "none", "n/a", "na", "nan", "unknown", "not specified",
    "not available", "not provided", "not mentioned", "-", "--", "tbd",
    "no information", "not stated", "undefined", "string", "[]", "{}",
}


def clean_scalar(value) -> str | None:
    """Flatten whatever the model returned into a trimmed string, or None."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, dict):
        # {"city": {"name": "Mount Pleasant"}} - take the first usable leaf.
        for item in value.values():
            found = clean_scalar(item)
            if found:
                return found
        return None
    if isinstance(value, list):
        parts = [clean_scalar(item) for item in value]
        joined = ", ".join(part for part in parts if part)
        return joined or None
    text = re.sub(r"\s+", " ", str(value)).strip().strip(",;|")
    if text.lower() in _NULLISH:
        return None
    return text or None


def truncate(value: str | None, limit: int) -> str | None:
    """Cut to the column width, preferring a word boundary."""
    if not value or len(value) <= limit:
        return value
    cut = value[:limit]
    space = cut.rfind(" ")
    # Only back off to a word boundary when that keeps most of the value.
    if space > limit * 0.6:
        cut = cut[:space]
    return cut.rstrip(" ,;-") or value[:limit]


# ---------------------------------------------------------------------------
# Grounding: does the resume actually say this?

_TOKEN = re.compile(r"[a-z0-9]+")
# Words that carry no evidence, so they are ignored when scoring a value.
_STOPWORDS = {
    "and", "or", "of", "the", "a", "an", "in", "at", "on", "for", "to", "with",
    "de", "la", "el", "inc", "llc", "ltd", "co", "company", "corp", "group",
}


def _tokens(text: str | None) -> list[str]:
    return [token for token in _TOKEN.findall((text or "").lower())
            if len(token) > 1 and token not in _STOPWORDS]


def build_evidence(text: str) -> set[str]:
    """The bag of words the resume actually contains."""
    return set(_tokens(text))


def is_grounded(value: str | None, evidence: set[str], *, threshold: float = 0.6) -> bool:
    """True when enough of the value's words appear in the resume.

    A strict substring test is too brittle: OCR spacing and the model's own
    re-casing both break it while the value is still faithful. Word overlap
    tolerates that but still catches an employer or specialty the model made up.
    """
    tokens = _tokens(value)
    if not tokens:
        return True  # nothing to disprove
    hits = sum(1 for token in tokens if token in evidence)
    return hits / len(tokens) >= threshold


# ---------------------------------------------------------------------------
# Coercion


def coerce_positions(value) -> list[dict]:
    """Force the model's "positions" into a list of well-shaped dicts."""
    if isinstance(value, dict):
        # Some replies wrap the list: {"positions": [...]} or {"1": {...}}.
        for key in ("positions", "work_experience", "experience", "jobs"):
            if isinstance(value.get(key), list):
                value = value[key]
                break
        else:
            value = [item for item in value.values() if isinstance(item, dict)] or [value]
    if not isinstance(value, list):
        return []

    cleaned: list[dict] = []
    for entry in value:
        if isinstance(entry, str):
            # "Infant room teacher - Creme de la Creme" arrived as a bare string.
            title, _, employer = entry.partition(" - ")
            entry = {"title": title, "employer": employer}
        if not isinstance(entry, dict):
            continue
        item = {field: clean_scalar(entry.get(field)) for field in POSITION_STRING_FIELDS}
        # Common key aliases from a model that drifted off the schema.
        item["employer"] = item["employer"] or clean_scalar(
            entry.get("company") or entry.get("employer_name") or entry.get("organization"))
        item["title"] = item["title"] or clean_scalar(
            entry.get("job_title") or entry.get("position") or entry.get("role"))
        item["start_date"] = item["start_date"] or clean_scalar(
            entry.get("start") or entry.get("from") or entry.get("start_year"))
        item["end_date"] = item["end_date"] or clean_scalar(
            entry.get("end") or entry.get("to") or entry.get("end_year"))
        item["description"] = item["description"] or clean_scalar(
            entry.get("duties") or entry.get("responsibilities") or entry.get("summary"))
        if item["title"] or item["employer"]:
            cleaned.append(item)
    return cleaned


def coerce_string_list(value) -> list[str]:
    """Force a list-valued field into a list of short strings."""
    if value is None:
        return []
    if isinstance(value, str):
        value = re.split(r"\s*[,;|\n]\s*", value)
    if isinstance(value, dict):
        value = list(value.values())
    if not isinstance(value, list):
        return []
    items: list[str] = []
    for entry in value:
        text = clean_scalar(entry)
        if text:
            items.append(text)
    return items


def coerce_number(value) -> float | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"\d+(?:\.\d+)?", str(value))
    return float(match.group(0)) if match else None


def coerce_raw_profile(raw) -> dict:
    """Normalise one model reply into the canonical key set and value types."""
    if not isinstance(raw, dict):
        return {}
    # A reply nested one level deep ({"profile": {...}}) is still usable.
    if not any(key in raw for key in PROFILE_STRING_FIELDS + ("positions",)):
        for value in raw.values():
            if isinstance(value, dict) and any(key in value for key in PROFILE_STRING_FIELDS):
                raw = value
                break

    result: dict = {field: clean_scalar(raw.get(field)) for field in PROFILE_STRING_FIELDS}
    result["full_name"] = result["full_name"] or clean_scalar(raw.get("name"))
    result["bio"] = result["bio"] or clean_scalar(
        raw.get("summary") or raw.get("professional_summary") or raw.get("profile_summary"))
    result["headline"] = result["headline"] or clean_scalar(
        raw.get("title") or raw.get("professional_headline"))
    result["specialty"] = result["specialty"] or clean_scalar(
        raw.get("speciality") or raw.get("field"))
    result["total_years_experience"] = coerce_number(
        raw.get("total_years_experience") or raw.get("years_experience"))
    result["positions"] = coerce_positions(
        raw.get("positions") or raw.get("work_experience") or raw.get("experience"))
    for field in ("certifications", "skills", "languages"):
        result[field] = coerce_string_list(raw.get(field))
    education = raw.get("education")
    result["education"] = education if isinstance(education, (list, dict, str)) else []
    return result


# ---------------------------------------------------------------------------
# The public entry point


# Grounding is only meaningful for values the resume should contain verbatim.
# first/last name are exempt: they may legitimately come from the file name.
_GROUNDED_PROFILE_FIELDS = (
    "city", "headline", "specialty", "current_employer", "profession_type",
    "work_authorization", "bio",
)
_GROUNDED_POSITION_FIELDS = ("title", "employer", "city", "description")


def validate_profile(fields: dict, source_text: str) -> tuple[dict, list[str]]:
    """Drop ungrounded values and clamp everything to its column width.

    Returns (fields, notes). The notes list names each value that was removed or
    shortened, so a run log shows why a column came out NULL.
    """
    notes: list[str] = []
    evidence = build_evidence(source_text)

    for field in _GROUNDED_PROFILE_FIELDS:
        value = fields.get(field)
        if not value:
            continue
        # A derived specialty is built from resume words by us, not the model,
        # and is allowed through even when it reads as a category name.
        if field == "specialty" and fields.get("_specialty_source") == "derived":
            continue
        if not is_grounded(value, evidence):
            notes.append(f"dropped ungrounded {field}: {str(value)[:60]!r}")
            fields[field] = None

    positions: list[dict] = []
    for position in fields.get("positions") or []:
        for field in _GROUNDED_POSITION_FIELDS:
            value = position.get(field)
            if value and not is_grounded(value, evidence):
                notes.append(f"dropped ungrounded position.{field}: {str(value)[:60]!r}")
                position[field] = None
        # A position with neither an employer nor a title is not a job.
        if position.get("title") or position.get("employer"):
            positions.append(position)
    fields["positions"] = positions

    for field, limit in PROFILE_LIMITS.items():
        value = fields.get(field)
        clamped = truncate(value, limit)
        if clamped != value:
            notes.append(f"truncated {field} to {limit} chars")
        fields[field] = clamped

    for position in fields["positions"]:
        for field, limit in WORK_HISTORY_LIMITS.items():
            position[field] = truncate(position.get(field), limit)

    return fields, notes
