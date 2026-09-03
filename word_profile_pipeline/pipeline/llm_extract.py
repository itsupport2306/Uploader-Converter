"""Structured profile extraction from PDF text using Qwen2.5.

The model is reached over an OpenAI-compatible /chat/completions endpoint
(Ollama, vLLM, LM Studio, or a hosted gateway). Base URL, API key, and model
name all come from the environment -- nothing is hardcoded.

Extraction runs as three small passes -- header, work history, list sections --
rather than one request for the whole profile. A 1.5B model given the entire
schema at once drops fields and runs out of output tokens mid-array; given one
short schema and one slice of the resume it stays complete. Where the server
supports it, a JSON schema is compiled into a sampling grammar so the reply is
valid JSON by construction.

If the model is unavailable or returns unusable output, a regex fallback fills
in what it can so a run never dies on one bad PDF.
"""
from __future__ import annotations

import json
import os
import re
import threading
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen

from . import config
from . import schema as field_schema

FIELDS = (
    "first_name", "last_name", "full_name", "email", "phone",
    "city", "state_code", "zip_code", "headline", "specialty",
    "profession_type", "work_authorization", "years_experience", "bio",
)

# List-valued fields that end up in resume_sections / education.
LIST_FIELDS = ("certifications", "skills", "languages", "education", "positions")

# ---------------------------------------------------------------------------
# Prompts
#
# The model on the other end is Qwen2.5-1.5B-Instruct. A 1.5B model asked for
# the entire profile -- header, ten jobs with duties, skills, certifications --
# in one JSON object reliably drifts: it drops "specialty", truncates the
# headline, and runs out of output tokens mid-array, which is where the invalid
# JSON came from. So the work is split into three small passes, each with a
# short schema and a short slice of the resume. Every pass is independently
# recoverable: a failure costs that section, not the whole profile.

SYSTEM_PROMPT = (
    "You are a precise resume parser. You read the text of one resume, which "
    "may come from OCR and may contain noise, and return structured data.\n"
    "Absolute rules:\n"
    "1. Output ONE JSON object and nothing else. No prose, no explanation, no "
    "code fence, no trailing text after the closing brace.\n"
    "2. Copy values from the resume text. Never invent, guess, infer or "
    "complete a value that is not written in the text.\n"
    "3. Use null for a missing string or number and [] for a missing list. "
    "Never write \"N/A\", \"unknown\", \"Not specified\", \"none\" or \"\".\n"
    "4. Never abbreviate or shorten a value. Copy it in full.\n"
    "5. Ignore OCR debris: stray single letters, the U+FFFD character, and "
    "short runs of unrelated capitals."
)

# --- pass 1: the header block -------------------------------------------------
#
# An Indeed/job-board profile PDF always opens the same way:
#     Donna Omara
#     Infant room teacher - Creme de la Creme     <- the headline
#     Mount Pleasant, SC                          <- the candidate's location
#     Professional summary
#     ...
# Feeding only that block keeps the pass short and stops the model from
# picking a city or title out of a job listed further down the page.

CORE_PROMPT_TEMPLATE = """Read the top of this resume and return the candidate's identity.

Return exactly these keys:
{{
  "full_name": string|null,          // the candidate's own name, no credentials
  "email": string|null,
  "phone": string|null,
  "city": string|null,               // the CANDIDATE's city, from the header
  "state_code": string|null,         // 2-letter US state code, uppercase
  "zip_code": string|null,
  "headline": string|null,
  "specialty": string|null,
  "current_employer": string|null,
  "profession_type": string|null,
  "work_authorization": string|null,
  "bio": string|null,
  "total_years_experience": number|null
}}

headline: the professional title line that sits directly under the name. Copy
the WHOLE line including the employer if it is written as "Title - Employer".
Do not shorten it and do not stop at the dash.

specialty: the candidate's own job title, copied from that same headline line.
The line is written "Specialty - Current organization", so the specialty is the
part BEFORE the dash and the organization is the part after it:
  "Infant room teacher - Creme de la Creme"
  -> specialty "Infant room teacher", current_employer "Creme de la Creme"
Copy the specialty exactly as written. Do NOT replace it with a broader
category such as "Childcare", "Education" or "Nursing", and do not include the
employer in it. If the resume states a specialty explicitly somewhere else
("Specialty: Pediatrics"), prefer that wording instead.

current_employer: the organization on that headline line, after the dash.

profession_type: the occupation itself, one or two words -- "Teacher", "Nurse",
"Driver", "Caregiver", "Manager".

city / state_code: the candidate's own location from the header block, not an
employer's address. "Mount Pleasant, SC" -> city "Mount Pleasant",
state_code "SC".

bio: the professional summary / about / additional information text, copied
verbatim. If the resume has no such paragraph, use null -- do not write one.
A work-authorization sentence on its own is not a bio; put it in
work_authorization instead.

total_years_experience: only a total career length that the resume states or
that its earliest job start clearly implies. Never one job's length.

RESUME HEADER:
---
{text}
---
JSON:"""

# --- pass 2: the experience section ------------------------------------------
#
# Positions are the field that used to get cut off, because they are the bulk
# of the output. They now get their own call (or several, for a long history),
# with nothing else competing for the token budget.

POSITIONS_PROMPT_TEMPLATE = """List EVERY job in this section of a resume's work history.

Return exactly this shape:
{{
  "positions": [
    {{
      "title": string|null,
      "employer": string|null,
      "city": string|null,
      "state_code": string|null,
      "start_date": string|null,
      "end_date": string|null,
      "description": string|null
    }}
  ]
}}

Rules:
- One entry per job heading. A heading is written "Title - Employer".
- Include short entries that have only a heading and no dates or duties.
- Copy the dates as written ("August 2025", "2013"). Use "Present" for a
  current job. Never calculate or invent a date. A span such as "11 mo" or
  "3 yr 2 mo" is a duration, not a date -- ignore it.
- description: the duty text under that heading, copied and tidied. null if the
  heading has none.
- Do not merge two jobs and do not repeat one.
- Return the jobs in the order they appear.

WORK HISTORY:
---
{text}
---
JSON:"""

# --- pass 3: the list sections ------------------------------------------------

LISTS_PROMPT_TEMPLATE = """Extract the list sections of this resume.

Return exactly this shape:
{{
  "certifications": [string],
  "skills": [string],
  "languages": [string],
  "education": [
    {{"degree": string|null, "field": string|null, "school": string|null, "year": string|null}}
  ]
}}

Rules:
- Each item is a short phrase written once. Never a sentence, never a duty.
- Deduplicate: the same certification listed twice appears once.
- Use [] for a section the text does not contain.

SECTIONS:
---
{text}
---
JSON:"""


US_STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR", "california": "CA",
    "colorado": "CO", "connecticut": "CT", "delaware": "DE", "district of columbia": "DC",
    "florida": "FL", "georgia": "GA", "hawaii": "HI", "idaho": "ID", "illinois": "IL",
    "indiana": "IN", "iowa": "IA", "kansas": "KS", "kentucky": "KY", "louisiana": "LA",
    "maine": "ME", "maryland": "MD", "massachusetts": "MA", "michigan": "MI", "minnesota": "MN",
    "mississippi": "MS", "missouri": "MO", "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM", "new york": "NY",
    "north carolina": "NC", "north dakota": "ND", "ohio": "OH", "oklahoma": "OK", "oregon": "OR",
    "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC", "south dakota": "SD",
    "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT", "virginia": "VA",
    "washington": "WA", "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
}
STATE_CODES = set(US_STATES.values())


class LlmError(RuntimeError):
    pass


def llm_enabled() -> bool:
    return config.env_bool("LLM_ENABLED", True)


def model_name() -> str:
    return os.environ.get("LLM_MODEL", "")

def _endpoint() -> str:
    base = (os.environ.get("LLM_BASE_URL") or "").strip().rstrip("/")
    if not base:
        raise LlmError("LLM_BASE_URL is not set.")
    if base.endswith("/chat/completions"):
        return base
    return urljoin(base + "/", "chat/completions")


# llama.cpp serves one request per slot. Two PDF workers firing at the same
# server would queue inside it anyway, and on a single-slot build the second
# request is rejected outright, so the calls are serialised here instead.
# Built lazily: the env file is loaded after this module is imported.
_llm_gate = None
_gate_lock = threading.Lock()


def llm_gate() -> threading.BoundedSemaphore:
    global _llm_gate
    with _gate_lock:
        if _llm_gate is None:
            _llm_gate = threading.BoundedSemaphore(
                max(1, config.env_int("LLM_MAX_CONCURRENCY", 1)))
        return _llm_gate

# Set once, the first time the server refuses a json_schema response_format, so
# the remaining calls in the run do not each pay for the same rejected attempt.
_schema_unsupported = False
_schema_lock = threading.Lock()


# --- JSON schemas ------------------------------------------------------------
#
# llama.cpp compiles a json_schema response_format into a GBNF grammar and
# constrains sampling with it, which is what makes "valid JSON every time" a
# property of the decoder rather than a hope about the prompt. Servers that do
# not support it fall back to json_object, and then to the repair path below.

_NULLABLE_STRING = {"type": ["string", "null"]}

CORE_SCHEMA = {
    "type": "object",
    "properties": {
        "full_name": _NULLABLE_STRING,
        "email": _NULLABLE_STRING,
        "phone": _NULLABLE_STRING,
        "city": _NULLABLE_STRING,
        "state_code": _NULLABLE_STRING,
        "zip_code": _NULLABLE_STRING,
        "headline": _NULLABLE_STRING,
        "specialty": _NULLABLE_STRING,
        "current_employer": _NULLABLE_STRING,
        "profession_type": _NULLABLE_STRING,
        "work_authorization": _NULLABLE_STRING,
        "bio": _NULLABLE_STRING,
        "total_years_experience": {"type": ["number", "null"]},
    },
    "required": [
        "full_name", "city", "state_code", "headline", "specialty",
        "profession_type", "bio",
    ],
    "additionalProperties": False,
}

POSITIONS_SCHEMA = {
    "type": "object",
    "properties": {
        "positions": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "title": _NULLABLE_STRING,
                    "employer": _NULLABLE_STRING,
                    "city": _NULLABLE_STRING,
                    "state_code": _NULLABLE_STRING,
                    "start_date": _NULLABLE_STRING,
                    "end_date": _NULLABLE_STRING,
                    "description": _NULLABLE_STRING,
                },
                "required": ["title", "employer"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["positions"],
    "additionalProperties": False,
}

LISTS_SCHEMA = {
    "type": "object",
    "properties": {
        "certifications": {"type": "array", "items": {"type": "string"}},
        "skills": {"type": "array", "items": {"type": "string"}},
        "languages": {"type": "array", "items": {"type": "string"}},
        "education": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "degree": _NULLABLE_STRING,
                    "field": _NULLABLE_STRING,
                    "school": _NULLABLE_STRING,
                    "year": _NULLABLE_STRING,
                },
                "additionalProperties": False,
            },
        },
    },
    "required": ["certifications", "skills", "languages", "education"],
    "additionalProperties": False,
}


def _post(payload: dict, *, timeout: int) -> dict:
    data = json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    api_key = os.environ.get("LLM_API_KEY")
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    request = Request(_endpoint(), data=data, headers=headers, method="POST")
    with urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8", errors="replace") or "{}")


def _call_model(
    messages: list[dict],
    *,
    timeout: int,
    temperature: float,
    max_tokens: int,
    schema: dict | None = None,
) -> str:
    """One chat completion. Returns the assistant's raw content string."""
    global _schema_unsupported

    def build(use_schema: bool) -> dict:
        payload = {
            "model": model_name(),
            "messages": messages,
            "temperature": temperature,
            "stream": False,
            "max_tokens": max_tokens,
        }
        if use_schema and schema is not None:
            payload["response_format"] = {
                "type": "json_schema",
                "json_schema": {"name": "profile", "strict": True, "schema": schema},
            }
        else:
            payload["response_format"] = {"type": "json_object"}
        return payload

    with _schema_lock:
        use_schema = schema is not None and not _schema_unsupported

    last_error: Exception | None = None
    for attempt in range(1, 4):
        try:
            body = _post(build(use_schema), timeout=timeout)
            choices = body.get("choices") or []
            if not choices:
                raise LlmError(f"Model returned no choices: {str(body)[:200]}")
            return (choices[0].get("message") or {}).get("content") or ""
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:300]
            # A server that does not understand json_schema rejects the request
            # outright. Drop to json_object for this call and the rest of the run
            # rather than reporting the whole extraction as failed.
            if use_schema and exc.code in {400, 404, 422, 500}:
                with _schema_lock:
                    _schema_unsupported = True
                use_schema = False
                last_error = LlmError(f"HTTP {exc.code}: {detail}")
                continue
            last_error = LlmError(f"HTTP {exc.code} from model endpoint: {detail}")
            if exc.code not in {408, 429, 500, 502, 503, 504} or attempt == 3:
                raise last_error from exc
        except (URLError, TimeoutError) as exc:
            last_error = LlmError(f"Could not reach the model at {_endpoint()}: {exc}")
            if attempt == 3:
                raise last_error from exc
        time.sleep(min(20, 2 ** attempt))

    raise last_error or LlmError("Model call failed.")


# ---------------------------------------------------------------------------
# JSON recovery
#
# With a grammar-constrained server none of this fires. It exists for the
# servers and builds that ignore response_format, where a small model produces
# JSON with a code fence around it, a trailing comma, a // comment, Python
# literals, or an object that simply ran out of output tokens.

_FENCE = re.compile(r"```(?:json|JSON)?\s*(.+?)(?:```|$)", re.DOTALL)
_TRAILING_COMMA = re.compile(r",(\s*[}\]])")
_LINE_COMMENT = re.compile(r"(?<![:\"'\\])//[^\n\"]*$", re.MULTILINE)
_BLOCK_COMMENT = re.compile(r"/\*.*?\*/", re.DOTALL)
_PY_LITERAL = re.compile(r"(?<![\w\"'])(True|False|None)(?![\w\"'])")
_PY_LITERALS = {"True": "true", "False": "false", "None": "null"}


def _tidy_json_text(text: str) -> str:
    """Fix the syntax errors a small model actually makes, and nothing else."""
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = _BLOCK_COMMENT.sub("", text)
    text = _LINE_COMMENT.sub("", text)
    text = _PY_LITERAL.sub(lambda m: _PY_LITERALS[m.group(1)], text)
    text = _TRAILING_COMMA.sub(r"\1", text)
    return text.strip()


def _repair_truncated_json(fragment: str) -> dict | None:
    """Recover a reply that ran out of output tokens mid-object.

    A resume with a dozen jobs produces a long object, and losing the whole
    extraction over the last few characters would drop us to regex for a reply
    that was 95% complete. The trailing partial value is discarded, then the
    open brackets are closed.
    """
    # Walk the text tracking string state, and remember the last position where
    # the structure was safe to cut: just after a completed element.
    depth: list[str] = []
    in_string = False
    escaped = False
    cut = None
    for index, char in enumerate(fragment):
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
                cut = index + 1
            continue
        if char == '"':
            in_string = True
        elif char in "{[":
            depth.append("}" if char == "{" else "]")
        elif char in "}]":
            if depth:
                depth.pop()
            cut = index + 1
        elif char == ",":
            cut = index  # drop the dangling comma with the partial element
        elif char.isdigit() or char in "eE.+-" or char.isalpha():
            cut = index + 1

    if cut is None:
        return None
    candidate = fragment[:cut].rstrip().rstrip(",")
    # Re-derive the still-open brackets for the trimmed text.
    depth = []
    in_string = False
    escaped = False
    for char in candidate:
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char in "{[":
            depth.append("}" if char == "{" else "]")
        elif char in "}]" and depth:
            depth.pop()
    if in_string:
        candidate += '"'
    # A cut that landed on a key with no value would close as {"key"} - drop it.
    candidate = re.sub(r",?\s*\"[^\"]*\"\s*:\s*$", "", candidate)
    candidate += "".join(reversed(depth))

    try:
        parsed = json.loads(candidate)
    except json.JSONDecodeError:
        return None
    return parsed if isinstance(parsed, dict) else None


def _parse_json_object(raw: str) -> dict:
    """Turn a model reply into a dict, repairing what can be repaired."""
    text = (raw or "").strip()
    fenced = _FENCE.search(text)
    if fenced:
        text = fenced.group(1).strip()

    for candidate in _json_candidates(text):
        try:
            parsed = json.loads(candidate)
        except json.JSONDecodeError:
            parsed = _repair_truncated_json(candidate)
        if isinstance(parsed, dict):
            return parsed
    raise LlmError(f"Model returned unusable JSON: {text[:200]}")


def _json_candidates(text: str):
    """Progressively more aggressive readings of the reply."""
    yield text
    tidied = _tidy_json_text(text)
    if tidied != text:
        yield tidied
    for source in (tidied, text):
        start, end = source.find("{"), source.rfind("}")
        if start != -1 and end > start:
            yield source[start:end + 1]
        if start != -1:
            # No usable closing brace: the reply hit the token limit.
            yield source[start:]


# Appended to the second attempt when the first reply could not be parsed.
NUDGE = (
    "\n\nReply with ONLY the JSON object. Start your reply with { and end it "
    "with }. Do not write anything before or after it."
)


def _ask_json(
    prompt: str,
    schema: dict,
    *,
    max_tokens: int,
    label: str,
) -> dict:
    """One extraction pass. Retries once with a blunter instruction, then gives up.

    Raising is per-pass, so a failed positions pass still leaves a good header
    pass in place instead of dropping the whole profile to regex.
    """
    timeout = config.env_int("LLM_TIMEOUT_SECONDS", 180)
    temperature = config.env_float("LLM_TEMPERATURE", 0.0)
    attempts = [
        [{"role": "system", "content": SYSTEM_PROMPT},
         {"role": "user", "content": prompt}],
        [{"role": "system", "content": SYSTEM_PROMPT},
         {"role": "user", "content":
          prompt + NUDGE}],
    ]
    last_error: Exception | None = None
    for messages in attempts:
        with llm_gate():
            try:
                content = _call_model(
                    messages,
                    timeout=timeout,
                    temperature=temperature,
                    max_tokens=max_tokens,
                    schema=schema,
                )
            except LlmError:
                # A transport failure will not be fixed by rewording the prompt.
                raise
        try:
            return _parse_json_object(content)
        except LlmError as exc:
            last_error = exc
    raise LlmError(f"{label}: {last_error}")


# ---------------------------------------------------------------------------
# Slicing the resume for each pass
#
# Each pass sees only the part of the resume it needs. That is what keeps a
# 1.5B model's output short enough to stay complete and well-formed, and it
# stops the header pass from picking a city out of a job halfway down the page.

_EXPERIENCE_HEADINGS = (
    "experience", "work experience", "professional experience",
    "employment history", "work history", "employment",
)
_LIST_HEADINGS = (
    "certifications & licenses", "certifications and licenses", "certifications",
    "licenses & certifications", "licenses and certifications", "licenses",
    "licensure", "credentials", "skills", "technical skills", "core competencies",
    "competencies", "languages", "language proficiency", "education",
    "education & training", "academic background",
)


def _heading_index(lines: list[str], headings) -> int | None:
    for index, line in enumerate(lines):
        if line.lower().strip(": ") in headings:
            return index
    return None


def split_sections(text: str) -> dict:
    """Cut the resume into the header, the experience block and the list blocks.

    Falls back to a character split when a resume has no recognisable headings,
    so a badly OCR'd page still produces three non-empty slices.
    """
    lines = (text or "").split("\n")
    experience_at = _heading_index(lines, _EXPERIENCE_HEADINGS)
    lists_at = None
    for index, line in enumerate(lines):
        if line.lower().strip(": ") in _LIST_HEADINGS and (
            experience_at is None or index > experience_at
        ):
            lists_at = index
            break

    if experience_at is None:
        head_end = min(len(lines), 40)
        return {
            "header": "\n".join(lines[:head_end]),
            "experience": "\n".join(lines[head_end:]),
            "lists": "\n".join(lines[head_end:]),
        }

    header = "\n".join(lines[:experience_at])
    experience_end = lists_at if lists_at is not None else len(lines)
    return {
        "header": header,
        "experience": "\n".join(lines[experience_at + 1:experience_end]),
        "lists": "\n".join(lines[experience_end:]) if lists_at is not None else "",
    }


def _clip(text: str, limit: int) -> str:
    text = (text or "").strip()
    return text if len(text) <= limit else text[:limit].rsplit("\n", 1)[0] + "\n[...truncated...]"


def _chunk_experience(text: str, limit: int) -> list[str]:
    """Split a long work history on job-heading boundaries, never mid-job."""
    text = (text or "").strip()
    if not text:
        return []
    if len(text) <= limit:
        return [text]
    chunks: list[str] = []
    current: list[str] = []
    size = 0
    for block in re.split(r"\n\s*\n", text):
        block_size = len(block) + 2
        if current and size + block_size > limit:
            chunks.append("\n\n".join(current))
            current, size = [], 0
        current.append(block)
        size += block_size
    if current:
        chunks.append("\n\n".join(current))
    return chunks


def extract_profile(text: str) -> dict:
    """Ask the model for the structured profile in three focused passes.

    Returns the merged raw dict. Raises LlmError only when nothing at all could
    be extracted; a single failed pass is reported through "_pass_errors" so the
    regex fallback can fill that section and the rest is still used.
    """
    cleaned = scrub_ocr_noise(text)
    sections = split_sections(cleaned)
    header_chars = config.env_int("LLM_HEADER_CHARS", 2500)
    chunk_chars = config.env_int("LLM_CHUNK_CHARS", 5000)
    max_tokens = config.env_int("LLM_MAX_TOKENS", 4096)

    raw: dict = {}
    errors: list[str] = []

    # Pass 1 -- header. The bio may live under "Additional information" at the
    # very bottom, so the tail is appended when the header block has no summary.
    header = _clip(sections["header"], header_chars)
    extra = _extra_information(cleaned)
    if extra:
        header = f"{header}\n\nAdditional information\n{_clip(extra, 1500)}"
    try:
        raw.update(_ask_json(
            CORE_PROMPT_TEMPLATE.format(text=header),
            CORE_SCHEMA, max_tokens=min(max_tokens, 900), label="header pass",
        ))
    except LlmError as exc:
        errors.append(str(exc))

    # Pass 2 -- work history, chunked so a long list never hits the token cap.
    positions: list[dict] = []
    for chunk in _chunk_experience(sections["experience"], chunk_chars):
        try:
            reply = _ask_json(
                POSITIONS_PROMPT_TEMPLATE.format(text=chunk),
                POSITIONS_SCHEMA, max_tokens=max_tokens, label="positions pass",
            )
        except LlmError as exc:
            errors.append(str(exc))
            continue
        positions.extend(field_schema.coerce_positions(reply.get("positions")))
    if positions:
        raw["positions"] = positions

    # Pass 3 -- the list sections. Optional: the regex pass handles these well,
    # so a failure here is recorded and otherwise ignored.
    lists_text = sections["lists"] or sections["experience"]
    if lists_text.strip():
        try:
            raw.update(_ask_json(
                LISTS_PROMPT_TEMPLATE.format(text=_clip(lists_text, chunk_chars)),
                LISTS_SCHEMA, max_tokens=max_tokens, label="lists pass",
            ))
        except LlmError as exc:
            errors.append(str(exc))

    if not raw:
        raise LlmError("; ".join(errors) or "the model returned nothing usable")
    if errors:
        raw["_pass_errors"] = errors
    return raw


_EXTRA_HEADINGS = ("additional information", "about", "about me", "professional summary",
                   "summary", "profile", "objective")


def _extra_information(text: str) -> str:
    """The free-text paragraph a job-board profile puts at the very bottom.

    On an Indeed profile the "Professional summary" heading is often empty and
    the candidate's own words sit under "Additional information" instead, so the
    header pass is shown both.
    """
    lines = [line.strip() for line in (text or "").split("\n")]
    for index, line in enumerate(lines):
        if line.lower().strip(": ") == "additional information":
            return "\n".join(lines[index + 1:]).strip()
    return ""


# ---------------------------------------------------------------------------
# OCR clean-up
#
# Tesseract turns the icon glyphs a resume uses for bullets into replacement
# characters and short runs of unrelated capitals ("Dn Hb HB BD YH DW DW").
# Left in, the model copies them into names and skills.

_OCR_ICON_CHARS = "�■●▪•"

# Symbol characters an icon glyph decodes to. Letters and ordinary punctuation
# are left alone so real content is never touched.
_ICON_SYMBOLS = re.compile(r"[^\w\s,.:;'\"()/&+#%$@!?\[\]{}=*<>|~`^\\-]")

# A line that is only short letter tokens is an icon run, not content
# ("Dn Hb HB BD YH DW DW" is what a column of bullet icons OCRs to).
_ICON_RUN = re.compile(r"^(?:[A-Za-z]{1,3}[ ,.]+){2,}[A-Za-z]{1,3}[.]?$")

# Leading punctuation left where an icon used to be ("= five towns college").
_LEADING_PUNCT = re.compile(r"^[\W_]+")

# A resume's bullet icons OCR as a stray 1-2 letter token before the real text
# ("Fy Infant room teacher"). Genuine short leading words must survive.
_REAL_SHORT_WORDS = {
    "i", "a", "an", "at", "as", "in", "on", "of", "to", "up", "by", "or", "is", "it",
    "dr", "mr", "ms", "mx", "st", "us", "rn", "md", "do", "np", "pa", "bs", "ba", "ms",
}
_LEADING_ICON_TOKEN = re.compile(r"^([A-Za-z]{1,2})\s+(?=[A-Za-z].*\s)")


def _strip_leading_icon(line: str) -> str:
    match = _LEADING_ICON_TOKEN.match(line)
    if match and match.group(1).lower() not in _REAL_SHORT_WORDS:
        return line[match.end():]
    return line


def scrub_ocr_noise(text: str) -> str:
    """Drop icon glyphs and icon-run lines so the model sees only real content."""
    lines: list[str] = []
    for raw in (text or "").split("\n"):
        line = raw
        for char in _OCR_ICON_CHARS:
            line = line.replace(char, "")
        line = line.replace("’", "'").replace("‘", "'")
        line = line.replace("“", '"').replace("”", '"')
        # Tesseract routinely reads a sentence-initial "I" as "|" or "l".
        line = re.sub(r"^[|]\s+(?=[a-z])", "I ", line)
        # Runs of spaces are kept (squashed to exactly two) because they are the
        # only surviving trace of a column break in an OCR'd skills grid.
        line = _ICON_SYMBOLS.sub(" ", line)
        line = re.sub(r"\t", " ", line)
        line = re.sub(r" {2,}", "  ", line).strip()
        if not line:
            lines.append("")
            continue
        if _ICON_RUN.match(line):
            continue
        line = _strip_leading_icon(_LEADING_PUNCT.sub("", line)).strip()
        if line:
            lines.append(line)
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


# ---------------------------------------------------------------------------
# Normalisation and regex fallback


def _clean(value) -> str | None:
    if value is None:
        return None
    text = re.sub(r"\s+", " ", str(value)).strip().strip(",;|")
    if not text or text.lower() in {"null", "none", "n/a", "na", "unknown", "-"}:
        return None
    return text


def normalize_state(value) -> str | None:
    text = _clean(value)
    if not text:
        return None
    upper = text.upper()
    if upper in STATE_CODES:
        return upper
    return US_STATES.get(text.lower())


def _title_case(value: str | None) -> str | None:
    if not value:
        return None
    return " ".join(part.capitalize() if part.islower() or part.isupper() else part
                    for part in value.split())


def split_full_name(full_name: str | None) -> tuple[str | None, str | None]:
    text = _clean(full_name)
    if not text:
        return None, None
    text = re.split(r"[,(]", text)[0].strip()
    parts = [p.strip(".,") for p in text.split() if p.strip(".,")]
    if not parts:
        return None, None
    if len(parts) == 1:
        return _title_case(parts[0]), None
    return _title_case(parts[0]), _title_case(parts[-1])


_NAME_LINE = re.compile(r"^[A-Z][A-Za-z'\-]+(?: [A-Z][A-Za-z'\-.]+){1,3}$")
_CITY_STATE = re.compile(r"\b([A-Z][A-Za-z .'\-]{2,30}),\s*([A-Z]{2})\b")

# Post-nominals that follow a name. Several collide with state codes (MD, PA,
# OR, IN...), so a "City, ST" match has to be checked against them.
CREDENTIALS = {
    "RN", "BSN", "MSN", "DNP", "APRN", "NP", "FNP", "PNP", "AGNP", "CRNA", "CNM",
    "LPN", "LVN", "CNA", "MD", "DO", "PA", "PAC", "PT", "OT", "RT", "MBA", "MPH",
    "PHD", "MS", "BS", "BA", "CCRN", "ACLS", "BLS", "FACS", "FACP", "ESQ", "JR", "SR",
}
_SECTION_HEADINGS = {
    "professional summary", "summary", "profile", "objective", "experience",
    "work experience", "professional experience", "education", "skills",
    "certifications", "licenses", "employment history", "contact", "references",
}


def _strip_credentials(line: str) -> str:
    """'Maria Gonzalez, RN, BSN' -> 'Maria Gonzalez'."""
    head = line.split(",")[0].strip()
    tokens = [t for t in head.split() if t.strip(".,").upper() not in CREDENTIALS]
    return " ".join(tokens).strip()


def regex_profile(text: str) -> dict:
    """Best-effort structured fields without the model."""
    profile: dict = {}
    # Same clean-up the model sees, so both paths agree on section boundaries.
    text = scrub_ocr_noise(text)
    lines = [line.strip() for line in (text or "").split("\n")]

    name_line_index = None
    for index, stripped in enumerate(lines[:12]):
        if stripped.lower() in _SECTION_HEADINGS:
            continue
        # An all-caps line is a heading, not a name.
        if stripped.isupper():
            continue
        candidate = _strip_credentials(stripped)
        if 4 <= len(candidate) <= 60 and _NAME_LINE.match(candidate):
            profile["full_name"] = candidate
            name_line_index = index
            break

    # The headline is the title line directly under the name. On a job-board
    # profile it reads "Infant room teacher - Creme de la Creme", and the whole
    # line is the headline -- taking only the part before the dash was what
    # truncated it.
    if name_line_index is not None:
        profile["headline"] = _headline_after_name(lines, name_line_index)

    # The candidate's own location sits in the header block. Searching the whole
    # document instead would find the first employer's address.
    header_end = _experience_line_index(lines)
    profile.update(_regex_location(lines, name_line_index, header_end))

    # A resume that names its specialty outright is the most reliable source.
    stated = re.search(
        r"(?:specialty|specialt(?:y|ies)|department|practice area)\s*[:\-]\s*([^\n]{2,60})",
        text or "", re.IGNORECASE,
    )
    if stated:
        profile["specialty"] = _clean(stated.group(1))
    else:
        # Otherwise it is the headline line, which reads
        # "Specialty - Current organization".
        title, employer = split_headline(profile.get("headline"))
        profile["specialty"] = title
        profile["current_employer"] = employer

    profile["bio"] = _regex_bio(text)

    profile.update(_regex_contact(text))
    profile.update(_regex_sections(text))
    profile["positions"] = _regex_positions(text)
    if not profile.get("headline") and profile.get("positions"):
        profile["headline"] = _position_headline(profile["positions"][0])
    return profile


# A summary section that only states work authorisation is not a bio; the
# candidate's own words are then under "Additional information" instead.
_SUMMARY_HEADINGS = (
    "professional summary", "summary", "profile", "objective", "about",
    "about me", "career summary", "overview",
)


def _regex_bio(text: str) -> str | None:
    """The candidate's own summary paragraph, wherever the profile put it."""
    for headings in (_SUMMARY_HEADINGS, ("additional information",)):
        summary = _clean(" ".join(_section_lines(text, headings)))
        if not summary or len(summary) <= 20:
            continue
        # "Authorized to work in the US for any employer" is a status line that
        # a job board prints under the summary heading, not a summary.
        if _WORK_AUTH.match(summary) and len(summary) < 120:
            continue
        return summary[:4000]
    return None


# The headline line of a job-board profile is "Specialty - Current organization".
# Only a spaced dash separates them: a hyphenated job title ("Full-time nanny")
# and an employer that contains one ("Creme de la Creme - Mt Pleasant") both
# survive, because neither has spaces around the hyphen.
_HEADLINE_SPLIT = re.compile(r"\s+[-–—]\s+")
# "2 yrs", "11 mo", "3 yr 2 mo", "Present" -- a tenure, not an organization.
_DURATION_ONLY = re.compile(
    r"(?:\d+\s*(?:\+)?\s*(?:yr|yrs|year|years|mo|mos|month|months)\s*)+|present|current",
    re.IGNORECASE,
)


def split_headline(headline: str | None) -> tuple[str | None, str | None]:
    """'Infant room teacher - Creme de la Creme' -> ('Infant room teacher', ...).

    Returns (specialty, current_employer). A headline with no dash is all
    specialty and names no employer.
    """
    text = _clean(headline)
    if not text:
        return None, None
    parts = _HEADLINE_SPLIT.split(text, maxsplit=1)
    if len(parts) == 1:
        return _clean(parts[0]), None
    specialty, employer = _clean(parts[0]), _clean(parts[1])
    # Some profiles put a tenure after the dash instead of an employer
    # ("Home Health CNA - 2 yrs"). That is not an organization.
    if employer and _DURATION_ONLY.fullmatch(employer):
        return specialty, None
    return specialty, employer


def _position_headline(position: dict) -> str | None:
    """Rebuild "Title - Employer" from a parsed position."""
    title, employer = position.get("title"), position.get("employer")
    if title and employer:
        return f"{title} - {employer}"
    return title or employer


def _experience_line_index(lines: list[str]) -> int:
    """Where the header block ends and the work history begins."""
    for index, line in enumerate(lines):
        if line.lower().strip(": ") in _EXPERIENCE_HEADINGS:
            return index
    return min(len(lines), 15)


# A headline line is a title, not a heading, a contact detail or a location.
_CONTACT_LINE = re.compile(r"@|\d{3}[\s.-]?\d{4}|^https?://|linkedin\.com", re.IGNORECASE)


def _headline_after_name(lines: list[str], name_index: int) -> str | None:
    """The first real title line under the name, copied whole."""
    for line in lines[name_index + 1:name_index + 5]:
        line = line.strip()
        if not line or line.lower().strip(": ") in _SECTION_HEADINGS:
            continue
        if _CONTACT_LINE.search(line):
            continue
        # "Mount Pleasant, SC" on its own is the location line, not a headline.
        match = _CITY_STATE.fullmatch(line)
        if match and match.group(2) in STATE_CODES:
            continue
        if 3 <= len(line) <= 255:
            return _clean(line)
    return None


def _regex_location(lines: list[str], name_index: int | None, header_end: int) -> dict:
    """City and state from the header block, falling back to the whole document."""
    for window in (lines[:header_end], lines):
        for index, line in enumerate(window):
            if window is lines and index == name_index:
                continue
            match = _CITY_STATE.search(line)
            if not match:
                continue
            city, state = match.group(1).strip(), match.group(2)
            if state not in STATE_CODES or len(city.split()) > 3:
                continue
            # "Gonzalez, RN, BSN" is a credential list, not a city and state.
            if state in CREDENTIALS and any(
                token.strip(".,").upper() in CREDENTIALS
                for token in line.split(",")[2:3] + line.split()[-2:]
            ):
                continue
            return {"city": _title_case(city), "state_code": state}
    return {}


# --- regex extraction of the remaining schema fields -----------------------

_EMAIL = re.compile(r"\b[\w.+-]+@[\w-]+\.[\w.-]{2,}\b")
# US phone: optional +1, then 3-3-4 with any of space, dot, dash or parens.
_PHONE = re.compile(r"(?<!\d)(?:\+?1[\s.-]?)?\(?([2-9]\d{2})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})(?!\d)")
_ZIP = re.compile(r"\b(\d{5})(?:-\d{4})?\b")
_WORK_AUTH = re.compile(
    r"^.{0,40}\b(authorized to work[^\n.]{0,60}|us citizen|green card[^\n.]{0,30}"
    r"|permanent resident|h-?1b[^\n.]{0,20}|opt\b[^\n.]{0,20}|ead\b[^\n.]{0,20})",
    re.IGNORECASE | re.MULTILINE,
)


def _regex_contact(text: str) -> dict:
    found: dict = {}
    email = _EMAIL.search(text or "")
    if email:
        found["email"] = email.group(0).lower()
    phone = _PHONE.search(text or "")
    if phone:
        found["phone"] = f"({phone.group(1)}) {phone.group(2)}-{phone.group(3)}"
    # Only trust a ZIP that sits right after a state code, so years and counts
    # elsewhere in the resume are never mistaken for one.
    zip_match = re.search(r"\b[A-Z]{2}\s+" + _ZIP.pattern, text or "")
    if zip_match:
        found["zip_code"] = zip_match.group(1)
    auth = _WORK_AUTH.search(text or "")
    if auth:
        found["work_authorization"] = _clean(auth.group(1))
    return found


# Section headings as they appear in a resume, mapped to our field names.
_SECTION_MAP = (
    ("certifications", (
        "certifications & licenses", "certifications and licenses", "certifications",
        "licenses & certifications", "licenses and certifications", "licenses",
        "licensure", "credentials",
    )),
    ("skills", ("skills", "technical skills", "core competencies", "competencies")),
    ("languages", ("languages", "language proficiency")),
    ("education", ("education", "education & training", "academic background")),
)
_ALL_HEADINGS = tuple(
    heading for _field, headings in _SECTION_MAP for heading in headings
) + (
    "experience", "work experience", "professional experience", "employment history",
    "work history", "professional summary", "summary", "additional information",
    "objective", "profile", "references", "awards", "publications", "contact",
)

# Boilerplate the source job board prints under each section heading.
_BOILERPLATE = re.compile(
    r"^related .{0,60} (?:come|comes) from the candidate.{0,30}$", re.IGNORECASE
)


def _section_lines(text: str, headings: tuple[str, ...], *, keep_blanks: bool = False) -> list[str]:
    """Lines between the given heading and the next heading for a *different* field.

    A resume often nests a heading of the same family inside its own section
    ("Certifications & licenses" ... "Certifications"). Stopping at those would
    truncate the section to its first entry, so only a foreign heading ends it.
    """
    lines = [line.strip() for line in (text or "").split("\n")]
    start = None
    for index, line in enumerate(lines):
        if line.lower().strip(":") in headings:
            start = index + 1
            break
    if start is None:
        return []
    collected: list[str] = []
    for line in lines[start:]:
        key = line.lower().strip(":")
        if key in headings:
            continue  # a nested heading of the same family
        if key in _ALL_HEADINGS:
            break
        if _BOILERPLATE.match(line):
            continue
        if line or keep_blanks:
            collected.append(line)
    return collected


def _dedupe(items) -> list[str]:
    """Case-insensitive de-duplication that keeps the first spelling seen."""
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        value = _clean(item)
        if not value:
            continue
        key = re.sub(r"[^a-z0-9]+", "", value.lower())
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result


# A resume's skills grid OCRs as one line per row, with the column break shown
# only by a word-boundary: "Toddler Care Infant Care English Preschool experience".
# Splitting where a lowercase word is followed by a capitalised one recovers the
# items. It is a heuristic for the no-model fallback; the model does this better.
_SKILL_COLUMN_BREAK = re.compile(r"(?<=[a-z])\s+(?=[A-Z])")


def _split_skill_line(line: str) -> list[str]:
    """A skills line is comma-, bullet-, wide-space- or column-separated."""
    parts: list[str] = []
    for chunk in re.split(r"\s*[,;|•·]\s*|\s{2,}|\.\s+", line):
        chunk = chunk.strip(" .")
        if not chunk:
            continue
        # Only worth splitting a long fused row; short entries stay intact.
        parts.extend(_SKILL_COLUMN_BREAK.split(chunk) if len(chunk) > 30 else [chunk])
    return [part.strip(" .") for part in parts if part.strip(" .")]


def _regex_sections(text: str) -> dict:
    found: dict = {}

    certifications = _dedupe(
        line for line in _section_lines(text, _SECTION_MAP[0][1])
        if 2 < len(line) <= 80 and line.lower() != "certifications"
    )
    if certifications:
        found["certifications"] = certifications

    skills: list[str] = []
    for line in _section_lines(text, _SECTION_MAP[1][1]):
        # A long sentence in the skills section is a duty, not a skill.
        if len(line) > 200:
            continue
        skills.extend(part for part in _split_skill_line(line) if 1 < len(part) <= 60)
    skills = _dedupe(skills)
    if skills:
        found["skills"] = skills

    languages = _dedupe(
        part for line in _section_lines(text, _SECTION_MAP[2][1])
        for part in _split_skill_line(line) if 1 < len(part) <= 40
    )
    if languages:
        found["languages"] = languages

    education_lines = _dedupe(
        line for line in _section_lines(text, _SECTION_MAP[3][1]) if 2 < len(line) <= 120
    )
    if education_lines:
        found["education"] = [_education_entry(line) for line in education_lines]

    return found


_DEGREE_WORDS = (
    "associate", "bachelor", "master", "doctor", "phd", "ph.d", "md", "bsn", "msn",
    "dnp", "adn", "rn", "diploma", "certificate", "ged", "high school", "some college",
    "b.s", "b.a", "m.s", "m.a", "mba",
)


def _education_entry(line: str) -> dict:
    """Turn one education line into the {degree, field, school, year} shape."""
    year = re.search(r"\b(19|20)\d{2}\b", line)
    lowered = line.lower()
    is_degree = any(word in lowered for word in _DEGREE_WORDS)
    return {
        "degree": _clean(line) if is_degree else None,
        "field": None,
        "school": None if is_degree else _clean(line),
        "year": year.group(0) if year else None,
    }


# "Title - Employer" heading, then an optional date line, then optional "City, ST".
# Commas are excluded from the title: a wrapped duty sentence that happens to
# contain " - " ("...spectrum , non - verbal...") otherwise parses as a new job.
_POSITION_HEADING = re.compile(r"^(?P<title>[^\n,;.]{2,60}?)\s+[-–—]\s+(?P<employer>[^\n;]{2,70})$")
_DATE_LINE = re.compile(
    r"^(?:(?P<start>(?:[A-Z][a-z]+\s+)?(?:19|20)\d{2}|Through\s+(?:19|20)\d{2})"
    r"\s*[-–—]+\s*(?P<end>Present|(?:[A-Z][a-z]+\s+)?(?:19|20)\d{2})"
    r"|(?P<only>Through\s+(?:19|20)\d{2}|(?:19|20)\d{2}))\b",
    re.IGNORECASE,
)


def _regex_positions(text: str) -> list[dict]:
    """Parse the experience section into position entries without the model."""
    lines = _section_lines(text, (
        "experience", "work experience", "professional experience",
        "employment history", "work history",
    ), keep_blanks=True)
    positions: list[dict] = []
    current: dict | None = None
    at_block_start = True
    for line in lines:
        if not line:
            at_block_start = True
            continue
        starts_block, at_block_start = at_block_start, False
        heading = _POSITION_HEADING.match(line)
        # A date line can also contain a dash, so check dates first.
        date_match = _DATE_LINE.match(line)
        # A job heading always opens a block; mid-paragraph it is a wrapped line.
        if heading and not date_match and starts_block:
            current = {
                "title": _clean(heading.group("title")),
                "employer": _clean(heading.group("employer")),
                "city": None, "state_code": None,
                "start_date": None, "end_date": None, "description": None,
            }
            positions.append(current)
            continue
        if current is None:
            continue
        if date_match:
            current["start_date"] = _clean(date_match.group("start") or date_match.group("only"))
            current["end_date"] = _clean(date_match.group("end"))
            continue
        city_state = _CITY_STATE.match(line)
        if city_state and city_state.group(2) in STATE_CODES and not current["city"]:
            current["city"] = _title_case(city_state.group(1).strip())
            current["state_code"] = city_state.group(2)
            continue
        current["description"] = f"{current['description']} {line}".strip() if current["description"] else line
    return _dedupe_positions(positions)


def _position_key(position: dict) -> str:
    return re.sub(
        r"[^a-z0-9]+", "",
        f"{position.get('title') or ''}|{position.get('employer') or ''}".lower(),
    )


def _dedupe_positions(positions) -> list[dict]:
    """One entry per title+employer; the richer duplicate wins."""
    merged: dict[str, dict] = {}
    for position in positions:
        if not isinstance(position, dict):
            continue
        entry = {
            "title": _clean(position.get("title")),
            "employer": _clean(position.get("employer")),
            "city": _title_case(_clean(position.get("city"))),
            "state_code": normalize_state(position.get("state_code")),
            "start_date": _clean(position.get("start_date") or position.get("start_year")),
            "end_date": _clean(position.get("end_date") or position.get("end_year")),
            "description": _clean(position.get("description")),
        }
        if not entry["title"] and not entry["employer"]:
            continue
        key = _position_key(entry)
        existing = merged.get(key)
        if existing is None:
            merged[key] = entry
            continue
        for field, value in entry.items():
            if value and not existing.get(field):
                existing[field] = value
    return list(merged.values())


def _merge_lists(model_value, fallback_value, *, limit: int) -> list[str]:
    """Model list first, regex list second, deduplicated and capped."""
    model_items = model_value if isinstance(model_value, list) else []
    fallback_items = fallback_value if isinstance(fallback_value, list) else []
    return _dedupe(list(model_items) + list(fallback_items))[:limit]


def _normalize_education(model_value, fallback_value) -> list[dict]:
    entries = model_value if isinstance(model_value, list) else []
    cleaned: list[dict] = []
    seen: set[str] = set()
    for entry in entries:
        if isinstance(entry, str):
            entry = _education_entry(entry)
        if not isinstance(entry, dict):
            continue
        item = {
            "degree": _clean(entry.get("degree")),
            "field": _clean(entry.get("field")),
            "school": _clean(entry.get("school")),
            "year": _clean(entry.get("year")),
        }
        if not any(item.values()):
            continue
        key = re.sub(r"[^a-z0-9]+", "", "".join(str(v or "") for v in item.values()).lower())
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(item)
    if cleaned:
        return cleaned
    return fallback_value if isinstance(fallback_value, list) else []


# ---------------------------------------------------------------------------
# Specialty
#
# Most resumes never write the word "specialty", so the model is asked for the
# candidate's field and this fills the gap when it declines. Each entry maps
# evidence words to the label used when those words are in the resume -- the
# label is a category name, not an invention about the candidate, and it is
# only applied when the resume supplies the evidence.

_SPECIALTY_RULES = (
    ("Early Childhood Education", (
        "infant room", "preschool", "toddler", "early childhood", "daycare",
        "day care", "childcare", "child care", "kinder", "pre-k", "nursery")),
    ("Special Education", ("special education", "special needs", "autism",
                           "developmental disabilities", "iep")),
    ("Childcare", ("nanny", "nannying", "babysitting", "au pair")),
    ("Nursing", ("registered nurse", " rn ", "bsn", "msn", "lpn", "cna",
                 "nursing", "patient care", "med surg", "icu", "telemetry")),
    ("Emergency Medicine", ("emergency department", "emergency room", " er ",
                            "trauma bay", "paramedic", "emt")),
    ("Allied Health", ("physical therapy", "occupational therapy",
                       "respiratory therapy", "radiology", "phlebotomy")),
    ("Teaching", ("teacher", "classroom", "curriculum", "lesson plan",
                  "instructor", "tutoring")),
    ("Commercial Driving", ("cdl", "class a driver", "tractor trailer",
                            "long haul", "bus driver", "delivery driver")),
    ("Retail Management", ("shift manager", "store manager", "merchandising",
                           "planograms", "cash register", "retail")),
    ("Food Service", ("cook", "chef", "kitchen", "cafeteria", "food service",
                      "barista", "server")),
    ("Construction & Trades", ("construction", "electrician", "plumbing",
                               "welding", "hvac", "pipeline", "meter reader")),
    ("Administrative Support", ("administrative assistant", "office manager",
                                "receptionist", "data entry", "scheduling")),
    ("Customer Service", ("customer service", "call center", "customer support",
                          "help desk")),
    ("Information Technology", ("software", "developer", "engineer", "devops",
                                "sql", "python", "network administrator")),
)


def derive_specialty(text: str, headline: str | None, positions: list[dict]) -> str | None:
    """Last-resort specialty for a resume with no headline line and none stated.

    Normally the specialty is copied straight off the headline line ("Infant
    room teacher - Creme de la Creme" -> "Infant room teacher"); this only runs
    when there is no such line to copy.

    Pick the specialty label whose evidence the resume actually contains.

    The most recent job title and the headline are weighted above the rest of
    the document, so a career changer is filed under what they do now rather
    than under whichever section happens to be longest.
    """
    body = (text or "").lower()
    recent = " ".join(str(part or "").lower() for part in (
        headline,
        positions[0].get("title") if positions else None,
        positions[0].get("description") if positions else None,
    ))
    best_label, best_score = None, 0
    for label, keywords in _SPECIALTY_RULES:
        score = 0
        for keyword in keywords:
            padded = f" {keyword.strip()} "
            if padded in f" {recent} ":
                score += 3
            elif keyword.strip() in body:
                score += 1
        if score > best_score:
            best_label, best_score = label, score
    # One passing mention somewhere in a long skills list is not a specialty.
    return best_label if best_score >= 2 else None


def normalize_profile(raw: dict, text: str) -> dict:
    """Merge model output with regex fallbacks into the canonical field set.

    The model wins wherever it produced a value; the regex pass fills the gaps
    and tops up individual blank fields inside a position. Everything then goes
    through schema validation, which drops values the resume does not support
    and clamps the rest to their column widths -- so a weak pass can only leave
    a field NULL, never write something wrong into it.
    """
    model = field_schema.coerce_raw_profile(raw)
    pass_errors = (raw or {}).get("_pass_errors") or []
    fallback = regex_profile(text)

    full_name = model.get("full_name") or fallback.get("full_name")
    first, last = split_full_name(full_name)

    # Positions: the model's list wins, the regex list fills in what it missed.
    positions = _dedupe_positions(model.get("positions") or [])
    fallback_positions = fallback.get("positions") or []
    if not positions:
        positions = fallback_positions
    else:
        by_key = {_position_key(position): position for position in positions}
        for candidate in fallback_positions:
            existing = by_key.get(_position_key(candidate))
            if existing is None:
                positions.append(candidate)
                continue
            for field, value in candidate.items():
                if value and not existing.get(field):
                    existing[field] = value

    # The full headline line, not just the job title. The regex pass reads it
    # straight off the line under the name, which is the most faithful source.
    headline = (
        model.get("headline")
        or fallback.get("headline")
        or (_position_headline(positions[0]) if positions else None)
    )

    bio = model.get("bio") or fallback.get("bio")
    if bio and len(bio) > 4000:
        bio = bio[:4000].rsplit(" ", 1)[0]

    email = model.get("email") or fallback.get("email")
    if email and not _EMAIL.fullmatch(email):
        email = fallback.get("email")

    # Specialty is the candidate's own job title off the headline line, which
    # reads "Specialty - Current organization" -- not a broader category. In
    # priority order: what the model read, what the regex pass read off that
    # same line, the line split here as a backstop, and only then a category
    # derived from the resume's keywords.
    headline_specialty, headline_employer = split_headline(headline)
    specialty = model.get("specialty") or fallback.get("specialty") or headline_specialty
    specialty_source = (
        "model" if model.get("specialty")
        else "regex" if fallback.get("specialty")
        else "headline" if headline_specialty
        else None
    )
    if not specialty:
        specialty = derive_specialty(text, headline, positions)
        specialty_source = "derived" if specialty else None

    current_employer = (
        model.get("current_employer")
        or fallback.get("current_employer")
        or headline_employer
        or (positions[0].get("employer") if positions else None)
    )

    fields = {
        "first_name": _title_case(first),
        "last_name": _title_case(last),
        "full_name": _title_case(full_name),
        "email": (email or "").lower() or None,
        "phone": model.get("phone") or fallback.get("phone"),
        "city": _title_case(model.get("city") or fallback.get("city")),
        "state_code": normalize_state(model.get("state_code")) or fallback.get("state_code"),
        "zip_code": model.get("zip_code") or fallback.get("zip_code"),
        "headline": headline,
        "specialty": specialty,
        "current_employer": current_employer,
        "profession_type": model.get("profession_type"),
        "work_authorization": model.get("work_authorization") or fallback.get("work_authorization"),
        "bio": bio,
        "certifications": _merge_lists(
            model.get("certifications"), fallback.get("certifications"), limit=100),
        "skills": _merge_lists(model.get("skills"), fallback.get("skills"), limit=200),
        "languages": _merge_lists(model.get("languages"), fallback.get("languages"), limit=30),
        "education": _normalize_education(model.get("education"), fallback.get("education")),
        "positions": positions,
        "_specialty_source": specialty_source,
        "_model_total_years": model.get("total_years_experience"),
        "_model_positions": positions,
        "_pass_errors": pass_errors,
    }

    fields, notes = field_schema.validate_profile(fields, text)

    # Anything validation removed was the model's invention. The regex pass
    # reads only from the resume, so its value for that field is grounded by
    # construction and is the right thing to fall back to -- leaving the column
    # NULL would lose a value the resume plainly contains.
    for field in ("city", "headline", "work_authorization", "bio"):
        if not fields.get(field) and fallback.get(field):
            fields[field] = field_schema.truncate(
                fallback[field], field_schema.PROFILE_LIMITS.get(field, 10_000))
            notes.append(f"restored {field} from the regex pass")
    if fields.get("city") and not fields.get("state_code"):
        fields["state_code"] = fallback.get("state_code")
    if not fields.get("specialty"):
        from_headline, _employer = split_headline(fields.get("headline"))
        if fallback.get("specialty"):
            fields["specialty"] = fallback["specialty"]
            fields["_specialty_source"] = "regex"
        elif from_headline:
            fields["specialty"] = field_schema.truncate(
                from_headline, field_schema.PROFILE_LIMITS["specialty"])
            fields["_specialty_source"] = "headline"

    # Validation may still have removed a headline or specialty the resume did
    # not support. Rebuild those from evidence rather than leaving a hole.
    if not fields.get("headline") and fields.get("positions"):
        fields["headline"] = field_schema.truncate(
            _position_headline(fields["positions"][0]), field_schema.PROFILE_LIMITS["headline"])
    if not fields.get("specialty"):
        # Last resort only: a resume with no headline line and no stated
        # specialty still gets filed under a field its own words support.
        derived = derive_specialty(text, fields.get("headline"), fields.get("positions") or [])
        if derived:
            fields["specialty"] = derived
            fields["_specialty_source"] = "derived"
    fields["_schema_notes"] = notes
    return fields
