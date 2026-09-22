"""Qwen-based detection of profiles whose stored name (or other identifying
fields) looks wrong.

reprocess.py used to select candidates with a fixed SQL substring match
(first_name/last_name ILIKE '%university%' / '%assistant%' / '%certified%').
That only catches bad names that happen to contain one of those words, and
misses junk that still looks like a plausible two-word name --
"Class Of" / "Samaritan Vie" / "Santa Hospital", or a first/last name pair
that simply does not match the person named in the underlying resume.

This module replaces that fixed list with a Qwen2.5 pass: for every profile
it shows the model the stored first_name/last_name/headline/specialty plus
an excerpt of the OCR'd resume text (resume_sections.raw_extracted_text,
already sitting in the row -- no re-OCR needed), and asks it to flag the
ones whose name looks incorrect or corrupted. The reprocessing logic itself
(reprocess.py's _reprocess_one) is untouched; this only changes how the
candidate profile_ids are chosen.

Talks to the same OpenAI-compatible Qwen endpoint already configured for
word_profile_pipeline (LLM_BASE_URL / LLM_MODEL / LLM_API_KEY /
LLM_TIMEOUT_SECONDS / LLM_MAX_CONCURRENCY in the shared .env), independently
of that pipeline's code.
"""
from __future__ import annotations

import json
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen

DEFAULT_BATCH_SIZE = 8
DEFAULT_EXCERPT_CHARS = 800
DEFAULT_MAX_TOKENS = 512


class QwenError(RuntimeError):
    pass


@dataclass(frozen=True)
class ScanCandidate:
    profile_id: str
    first_name: str | None
    last_name: str | None
    headline: str | None
    specialty: str | None
    excerpt: str


# ---------------------------------------------------------------------------
# Fetching candidates

SELECT_ALL_SQL = """
    SELECT profile_id, first_name, last_name, headline, specialty, resume_sections::text
    FROM {table}
    {where_clause}
    ORDER BY updated_at NULLS LAST, profile_id
"""
# Lets several servers each scan a disjoint created_at slice (START_DATE /
# END_DATE in .env, or --start-date / --end-date) instead of the whole table.
DATE_WHERE_SQL = "WHERE created_at BETWEEN :start_date AND :end_date"


def fetch_candidates(engine, table: str, uploader, cp, *, limit: int | None,
                      excerpt_chars: int, start_date: str | None = None,
                      end_date: str | None = None) -> list[ScanCandidate]:
    params: dict = {}
    where_clause = ""
    if start_date and end_date:
        where_clause = DATE_WHERE_SQL
        params = {"start_date": start_date, "end_date": end_date}
    sql = SELECT_ALL_SQL.format(table=cp._quoted_table(table), where_clause=where_clause)
    if limit:
        sql += " LIMIT :limit"
        params["limit"] = limit
    with engine.connect() as conn:
        rows = conn.execute(uploader._text(sql), params).all()
    candidates: list[ScanCandidate] = []
    for profile_id, first, last, headline, specialty, sections_text in rows:
        try:
            sections = json.loads(sections_text) if sections_text else {}
        except json.JSONDecodeError:
            sections = {}
        if not isinstance(sections, dict):
            sections = {}
        text = str(sections.get("raw_extracted_text") or "").strip()
        candidates.append(ScanCandidate(
            profile_id=str(profile_id), first_name=first, last_name=last,
            headline=headline, specialty=specialty, excerpt=text[:excerpt_chars],
        ))
    return candidates


# ---------------------------------------------------------------------------
# Prompting

SYSTEM_PROMPT = (
    "You review candidate profiles that were extracted by OCR from resume "
    "screenshots. For each profile you are given the name the pipeline "
    "stored (first_name, last_name), the headline/specialty it stored, and "
    "an excerpt of the text actually OCR'd from that resume.\n"
    "A stored name is WRONG when any of these is true:\n"
    "- it is not a plausible human first/last name at all -- an "
    "organization, a school or hospital, a job title, a certification, a "
    "generic resume word, a sentence fragment (e.g. 'Class Of'), or OCR "
    "noise\n"
    "- the resume excerpt contains a different person's name and the "
    "stored name does not match it\n"
    "- the excerpt has no name in it at all, so the stored name could not "
    "plausibly have come from this resume\n"
    "A stored name is CORRECT when it is an ordinary human name and either "
    "matches the name in the excerpt, or the excerpt is simply too short to "
    "contain a name (e.g. only a certifications list) while the stored name "
    "itself still looks like a genuine person's name.\n"
    "Only flag a profile when you are confident its name is wrong. Reply "
    "with ONLY a JSON object, no prose, no code fence."
)

USER_PROMPT_TEMPLATE = """Review these {count} profiles and return the ids of the ones whose stored name looks incorrect or corrupted.

PROFILES:
{items}

Return exactly this JSON shape:
{{"bad_ids": [string, ...]}}
Only include an id if its stored name is wrong. An empty list is a valid answer."""

RESULT_SCHEMA = {
    "type": "object",
    "properties": {"bad_ids": {"type": "array", "items": {"type": "string"}}},
    "required": ["bad_ids"],
    "additionalProperties": False,
}


def _format_item(local_id: str, c: ScanCandidate) -> str:
    headline = (c.headline or "").strip() or "(none)"
    specialty = (c.specialty or "").strip() or "(none)"
    excerpt = c.excerpt or "(no OCR text stored)"
    return (
        f"[{local_id}] first_name={c.first_name!r} last_name={c.last_name!r} "
        f"headline={headline!r} specialty={specialty!r}\n"
        f"resume excerpt:\n{excerpt}\n---"
    )


# ---------------------------------------------------------------------------
# Talking to the model (same LLM_* env vars as word_profile_pipeline)

def _endpoint() -> str:
    base = (os.environ.get("LLM_BASE_URL") or "").strip().rstrip("/")
    if not base:
        raise QwenError("LLM_BASE_URL is not set.")
    if base.endswith("/chat/completions"):
        return base
    return urljoin(base + "/", "chat/completions")


def require_llm_config() -> None:
    if os.environ.get("LLM_ENABLED", "true").strip().lower() in {"0", "false", "no", "off", ""}:
        raise QwenError("LLM_ENABLED is false; the Qwen detection pass needs it enabled.")
    if not (os.environ.get("LLM_BASE_URL") or "").strip():
        raise QwenError("LLM_BASE_URL is not set (needed for the Qwen detection pass).")
    if not (os.environ.get("LLM_MODEL") or "").strip():
        raise QwenError("LLM_MODEL is not set (needed for the Qwen detection pass).")


def _post(payload: dict, timeout: int) -> dict:
    data = json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    api_key = os.environ.get("LLM_API_KEY")
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    request = Request(_endpoint(), data=data, headers=headers, method="POST")
    with urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8", errors="replace") or "{}")


# Set once the server refuses a json_schema response_format, so the rest of
# the scan does not each pay for the same rejected attempt.
_schema_unsupported = False
_schema_lock = threading.Lock()


def _call_model(messages: list[dict], *, timeout: int, temperature: float, max_tokens: int) -> str:
    global _schema_unsupported

    def build(use_schema: bool) -> dict:
        payload = {
            "model": os.environ.get("LLM_MODEL", ""),
            "messages": messages,
            "temperature": temperature,
            "stream": False,
            "max_tokens": max_tokens,
        }
        if use_schema:
            payload["response_format"] = {
                "type": "json_schema",
                "json_schema": {"name": "bad_profiles", "strict": True, "schema": RESULT_SCHEMA},
            }
        else:
            payload["response_format"] = {"type": "json_object"}
        return payload

    with _schema_lock:
        use_schema = not _schema_unsupported

    last_error: Exception | None = None
    for attempt in range(1, 4):
        try:
            body = _post(build(use_schema), timeout)
            choices = body.get("choices") or []
            if not choices:
                raise QwenError(f"Model returned no choices: {str(body)[:200]}")
            return (choices[0].get("message") or {}).get("content") or ""
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:300]
            if use_schema and exc.code in {400, 404, 422, 500}:
                with _schema_lock:
                    _schema_unsupported = True
                use_schema = False
                last_error = QwenError(f"HTTP {exc.code}: {detail}")
                continue
            last_error = QwenError(f"HTTP {exc.code} from model endpoint: {detail}")
            if exc.code not in {408, 429, 500, 502, 503, 504} or attempt == 3:
                raise last_error from exc
        except (URLError, TimeoutError) as exc:
            last_error = QwenError(f"Could not reach the model at {_endpoint()}: {exc}")
            if attempt == 3:
                raise last_error from exc
        time.sleep(min(20, 2 ** attempt))
    raise last_error or QwenError("Model call failed.")


_FENCE = re.compile(r"```(?:json|JSON)?\s*(.+?)(?:```|$)", re.DOTALL)


def _parse_bad_ids(raw: str, valid_ids: set[str]) -> set[str]:
    text = (raw or "").strip()
    fenced = _FENCE.search(text)
    if fenced:
        text = fenced.group(1).strip()
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end > start:
        text = text[start:end + 1]
    try:
        obj = json.loads(text)
    except json.JSONDecodeError as exc:
        raise QwenError(f"Model returned unusable JSON: {text[:200]}") from exc
    ids = obj.get("bad_ids") if isinstance(obj, dict) else None
    if not isinstance(ids, list):
        raise QwenError(f"Model reply missing a 'bad_ids' list: {text[:200]}")
    # Guard against the model inventing or mangling an id: only ids we
    # actually offered it (the local "1", "2", ... labels) are honored.
    return {str(i) for i in ids if str(i) in valid_ids}


def classify_batch(candidates: list[ScanCandidate], *, timeout: int, max_tokens: int) -> dict[str, str]:
    """One Qwen call over a batch of profiles. Returns {profile_id: reason} for flagged ones."""
    if not candidates:
        return {}
    local_ids = [str(i) for i in range(1, len(candidates) + 1)]
    by_local = dict(zip(local_ids, candidates))
    items = "\n".join(_format_item(lid, c) for lid, c in zip(local_ids, candidates))
    prompt = USER_PROMPT_TEMPLATE.format(count=len(candidates), items=items)
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": prompt},
    ]
    content = _call_model(messages, timeout=timeout, temperature=0.0, max_tokens=max_tokens)
    flagged_local = _parse_bad_ids(content, set(local_ids))
    return {
        by_local[lid].profile_id: "qwen: stored name looks incorrect or corrupted"
        for lid in flagged_local
    }


# ---------------------------------------------------------------------------
# Scan manifest (skip already-classified profiles on the next run, like
# reprocess.py's own manifest skips already-reprocessed ones)

def _scanned_ids(path: Path) -> set[str]:
    scanned: set[str] = set()
    if not path.exists():
        return scanned
    with path.open("r", encoding="utf-8") as fh:
        for line in fh:
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            if rec.get("profile_id"):
                scanned.add(rec["profile_id"])
    return scanned


def _bad_ids_from_manifest(path: Path) -> set[str]:
    bad: set[str] = set()
    if not path.exists():
        return bad
    with path.open("r", encoding="utf-8") as fh:
        for line in fh:
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            if rec.get("verdict") == "bad" and rec.get("profile_id"):
                bad.add(rec["profile_id"])
    return bad


def _append_jsonl(path: Path, record: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(record, sort_keys=True, default=str) + "\n")


def scan_for_bad_profiles(
    engine, table: str, uploader, cp, *,
    manifest: Path,
    scan_limit: int | None = None,
    batch_size: int = DEFAULT_BATCH_SIZE,
    excerpt_chars: int = DEFAULT_EXCERPT_CHARS,
    concurrency: int = 1,
    timeout: int = 180,
    max_tokens: int = DEFAULT_MAX_TOKENS,
    rescan: bool = False,
    start_date: str | None = None,
    end_date: str | None = None,
) -> set[str]:
    """Classify every profile (or the first `scan_limit`) with Qwen and return
    the accumulated set of profile_ids whose name looks wrong.

    When `start_date`/`end_date` are given, only profiles with a created_at
    in that range are fetched at all -- the partition a server running this
    range is responsible for.

    Profiles already recorded in `manifest` are skipped unless `rescan` is
    set, so re-running only pays for profiles added since the last scan.
    """
    require_llm_config()
    candidates = fetch_candidates(
        engine, table, uploader, cp, limit=scan_limit, excerpt_chars=excerpt_chars,
        start_date=start_date, end_date=end_date,
    )
    already = set() if rescan else _scanned_ids(manifest)
    todo = [c for c in candidates if c.profile_id not in already]
    print(f"Qwen scan: {len(candidates)} profile(s) in range, {len(todo)} to classify "
          f"({len(candidates) - len(todo)} already scanned in {manifest.name}).")

    batches = [todo[i:i + batch_size] for i in range(0, len(todo), batch_size)]
    lock = threading.Lock()

    def _run_batch(batch: list[ScanCandidate]) -> None:
        try:
            flagged = classify_batch(batch, timeout=timeout, max_tokens=max_tokens)
        except QwenError as exc:
            print(f"Qwen scan: batch of {len(batch)} failed, leaving them unscanned: {exc}")
            return
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        with lock:
            for c in batch:
                verdict = "bad" if c.profile_id in flagged else "ok"
                _append_jsonl(manifest, {
                    "profile_id": c.profile_id, "verdict": verdict,
                    "reason": flagged.get(c.profile_id), "scanned_at": now,
                })
        if flagged:
            for pid in flagged:
                print(f"  flagged bad: {pid}")

    if batches:
        with ThreadPoolExecutor(max_workers=max(1, concurrency), thread_name_prefix="qwen-scan") as pool:
            list(pool.map(_run_batch, batches))

    bad_ids = _bad_ids_from_manifest(manifest)
    print(f"Qwen scan: {len(bad_ids)} profile(s) flagged bad in total.")
    return bad_ids
