from __future__ import annotations

import argparse
import base64
import csv
import importlib.util
import json
import os
import re
import sys
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from time import monotonic
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlparse
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parent
CONVERTER_DIR = ROOT / "Converter" / "Converter" / "Converter"
UPLOAD_SCRIPT = ROOT / "data_upload 5.py"
DEFAULT_OUTPUT_FOLDER_NAME = "generated_docx"
DEFAULT_OUTPUT_PARENT_DIR = Path(r"C:\virinchi")
DEFAULT_PROCESS_MANIFEST = ROOT / "complete_process_profiles_manifest.jsonl"
DEFAULT_UPLOAD_LOG = ROOT / "complete_process_upload_log.csv"
DEFAULT_TARGET_TABLE = "profiles"
DEFAULT_ENV_FILE = "env 1" if (ROOT / "env 1").exists() else ".env"
DEFAULT_ONEDRIVE_TENANT = os.environ.get("ONEDRIVE_TENANT", "common")
DEFAULT_ONEDRIVE_CLIENT_SECRET = os.environ.get("ONEDRIVE_CLIENT_SECRET")
DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
GRAPH_APP_SCOPE = "https://graph.microsoft.com/.default"

PROFILE_COLUMNS = {
    "profile_id", "user_id", "first_name", "last_name", "headline", "bio",
    "specialty", "profession_type", "years_experience", "city", "state_code",
    "lat", "lng", "open_to_work", "job_type_prefs", "pay_min_hourly",
    "available_date", "npi_number", "profile_photo_url", "resume_url",
    "completion_score", "source", "search_text", "created_at", "updated_at",
    "phone", "email", "provider_category", "american_board",
    "contact_updated_by_user_id", "contact_updated_by_email", "contact_updated_at",
    "is_listable", "zip_code", "resume_sections", "work_authorization",
    "education", "capture_source", "captured_by_user_id", "captured_by_email",
    "captured_at", "screen_reason", "screen_score", "screened_at",
    "merged_into", "merged_at",
}


@dataclass(frozen=True)
class ScreenshotSource:
    record_id: str
    source_file: str
    display_name: str
    docx_name: str
    local_path: Path | None = None
    remote_drive_id: str | None = None
    remote_item_id: str | None = None

    @property
    def is_remote(self) -> bool:
        return self.remote_drive_id is not None and self.remote_item_id is not None


class ConfidentialOneDriveClient:
    def __init__(self, *, client_id: str, tenant: str, client_secret: str):
        self.client_id = client_id
        self.tenant = tenant.strip().strip("/") or "common"
        self.client_secret = client_secret
        self.access_token = ""
        self.refresh_token = ""
        self.expires_at = 0.0

    @property
    def oauth_base_url(self) -> str:
        tenant = quote(self.tenant, safe="")
        return f"{converter.MICROSOFT_LOGIN_BASE_URL}/{tenant}/oauth2/v2.0"

    def authenticate(self) -> None:
        token = converter._post_form(
            f"{self.oauth_base_url}/token",
            {
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": GRAPH_APP_SCOPE,
            },
        )
        self._set_token(token)
        print("OneDrive app login complete.")

    def _set_token(self, token: dict) -> None:
        self.access_token = token["access_token"]
        self.expires_at = time.time() + int(token.get("expires_in", 3600))

    def _ensure_token(self) -> None:
        if self.access_token and time.time() < self.expires_at - 120:
            return
        self.authenticate()


def _load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load module from {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


sys.path.insert(0, str(CONVERTER_DIR))
converter = _load_module("screenshot_to_word", CONVERTER_DIR / "screenshot_to_word.py")
uploader = _load_module("data_upload_5", UPLOAD_SCRIPT)


IMAGE_EXTS = converter.IMAGE_EXTS


def _utc_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _append_manifest(path: Path, record: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(record, sort_keys=True) + "\n")


def _append_upload_log(path: Path, record: dict) -> None:
    columns = [
        "uploaded_at", "status", "profile_id", "first_name", "last_name",
        "email", "phone", "npi_number", "specialty", "profession_type",
        "provider_category", "city", "state_code", "resume_url",
        "cloudflare_key", "docx_sha256", "source_file", "docx_file", "error",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    write_header = not path.exists() or path.stat().st_size == 0
    with path.open("a", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
        if write_header:
            writer.writeheader()
        writer.writerow({column: record.get(column) for column in columns})


def _latest_manifest_status(path: Path | None) -> dict[str, dict]:
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


def _record_id(path: Path) -> str:
    return str(path.resolve()).lower()


def _docx_path_for(image_path: Path, output_dir: Path) -> Path:
    return output_dir / f"{image_path.stem}.docx"


def _default_output_dir_for(_: Path) -> Path:
    return DEFAULT_OUTPUT_PARENT_DIR / DEFAULT_OUTPUT_FOLDER_NAME


def _quoted_table(name: str) -> str:
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
        raise SystemExit(f"ERROR: unsafe table name: {name!r}")
    return f'"{name}"'


def _json_value(value) -> str:
    return json.dumps(value or [])


def _safe_cloudflare_segment(value: str | None) -> str:
    text = re.sub(r"[^A-Za-z0-9]+", "_", value or "Unknown")
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:80] or "Unknown"


def _cloudflare_key_for(path: Path, digest: str, *segments: str | None) -> str:
    folder = _safe_cloudflare_segment(next((value for value in segments if value), None))
    return f"resumes/{folder}/{digest}{path.suffix.lower()}"


def _is_remote_input(value: str) -> bool:
    parsed = urlparse(value)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def _optional_local_path(value: str | None, default: Path | None) -> Path | None:
    if value is None:
        return default
    normalized = value.strip()
    if not normalized or normalized.lower() in {"none", "null", "off", "false", "-"}:
        return None
    return Path(normalized).expanduser()


def _label_from_source(source: ScreenshotSource) -> str:
    return Path(source.display_name).name or source.display_name


def _virtual_docx_reference(source: ScreenshotSource) -> str:
    docx_like = Path(source.display_name).with_suffix(".docx")
    return f"memory://{docx_like.as_posix()}"


def _iter_local_screenshots(input_path: Path, limit: int | None) -> list[ScreenshotSource]:
    if input_path.is_file():
        if input_path.suffix.lower() not in IMAGE_EXTS:
            supported = ", ".join(sorted(IMAGE_EXTS))
            raise SystemExit(f"ERROR: unsupported input type: {input_path} (supported: {supported})")
        files = [input_path]
    elif input_path.is_dir():
        files = sorted(
            p for p in input_path.rglob("*")
            if p.is_file() and p.suffix.lower() in IMAGE_EXTS
        )
    else:
        raise SystemExit(f"ERROR: input path not found: {input_path}")
    if not files:
        raise SystemExit(f"ERROR: no supported screenshots found in {input_path}")
    selected = files[:limit] if limit else files
    return [
        ScreenshotSource(
            record_id=_record_id(path),
            source_file=str(path),
            display_name=str(path),
            docx_name=f"{path.stem}.docx",
            local_path=path,
        )
        for path in selected
    ]


def _encode_share_url(shared_url: str) -> str:
    encoded = base64.b64encode(shared_url.encode("utf-8")).decode("ascii")
    return "u!" + encoded.rstrip("=").replace("/", "_").replace("+", "-")


def _onedrive_request(
    client,
    method: str,
    path_or_url: str,
    *,
    body: dict | bytes | None = None,
    ok_statuses: set[int] | None = None,
    headers: dict[str, str] | None = None,
    content_type: str = "application/json",
    expect_json: bool = True,
):
    data = None
    if isinstance(body, dict):
        data = json.dumps(body).encode("utf-8")
    elif isinstance(body, bytes):
        data = body

    ok_statuses = ok_statuses or ({200, 201} if expect_json else {200})
    retry_statuses = {429, 500, 502, 503, 504}
    url = path_or_url if path_or_url.startswith(("http://", "https://")) else f"{converter.GRAPH_BASE_URL}{path_or_url}"
    token_lock = getattr(client, "_token_lock", None)
    if token_lock is None:
        token_lock = threading.Lock()
        setattr(client, "_token_lock", token_lock)

    for attempt in range(1, 9):
        with token_lock:
            client._ensure_token()
            access_token = client.access_token
        request_headers = {"Authorization": f"Bearer {access_token}"}
        if headers:
            request_headers.update(headers)
        if data is not None:
            request_headers.setdefault("Content-Type", content_type)
        req = Request(url, data=data, headers=request_headers, method=method)
        try:
            with urlopen(req, timeout=60) as resp:
                payload = resp.read()
                status = getattr(resp, "status", resp.getcode())
                if status not in ok_statuses:
                    raise converter.OneDriveError(f"HTTP {status}", status=status)
                if not expect_json:
                    return payload
                return json.loads(payload.decode("utf-8", errors="replace")) if payload else {}
        except HTTPError as exc:
            parsed = converter._read_http_error(exc)
            retry_after = None
            header_value = exc.headers.get("Retry-After") if exc.headers else None
            if header_value and header_value.isdigit():
                retry_after = int(header_value)
            error = converter.OneDriveError(
                parsed["message"],
                status=exc.code,
                payload=parsed["payload"],
                retry_after=retry_after,
            )
            if exc.code == 401 and getattr(client, "refresh_token", ""):
                with token_lock:
                    client.access_token = ""
                continue
            if exc.code in retry_statuses and attempt < 8:
                wait = retry_after or min(120, 2 ** attempt)
                print(
                    f"OneDrive is busy/throttling (HTTP {exc.code}); "
                    f"waiting {wait}s before retry {attempt + 1}/8..."
                )
                time.sleep(wait)
                continue
            raise error from exc
        except URLError as exc:
            raise converter.OneDriveError(f"Could not reach Microsoft endpoint: {exc}") from exc

    raise converter.OneDriveError("Microsoft Graph request failed after retries")


def _item_drive_id(item: dict) -> str:
    parent = item.get("parentReference") or {}
    drive_id = parent.get("driveId")
    if drive_id:
        return drive_id
    remote_parent = (item.get("remoteItem") or {}).get("parentReference") or {}
    drive_id = remote_parent.get("driveId")
    if drive_id:
        return drive_id
    raise converter.OneDriveError(f"Missing driveId for OneDrive item: {item.get('name') or item.get('id')}")


def _resolve_shared_item(client, shared_url: str) -> dict:
    print("Resolving shared OneDrive/SharePoint link...")
    token = _encode_share_url(shared_url)
    item = _onedrive_request(
        client,
        "GET",
        f"/shares/{token}/driveItem",
        headers={"Prefer": "redeemSharingLink"},
    )
    item_name = item.get("name") or item.get("id") or "shared item"
    item_kind = "folder" if "folder" in item else "file"
    print(f"Resolved shared item: {item_name} ({item_kind})")
    return item


def _iter_drive_children(client, drive_id: str, item_id: str):
    next_url = (
        f"{converter.GRAPH_BASE_URL}/drives/{quote(drive_id, safe='')}"
        f"/items/{quote(item_id, safe='')}/children?$top=200"
    )
    while next_url:
        payload = _onedrive_request(client, "GET", next_url)
        for item in payload.get("value", []):
            yield item
        next_url = payload.get("@odata.nextLink")


def _walk_remote_images(
    client,
    item: dict,
    *,
    path_parts: list[str],
    results: list[ScreenshotSource],
    limit: int | None,
) -> None:
    if limit is not None and len(results) >= limit:
        return

    name = item.get("name") or item.get("id") or "item"
    current_parts = path_parts + [name]

    if "file" in item:
        if Path(name).suffix.lower() in IMAGE_EXTS:
            display_name = "/".join(current_parts)
            results.append(
                ScreenshotSource(
                    record_id=f"onedrive://{_item_drive_id(item).lower()}/{str(item['id']).lower()}",
                    source_file=item.get("webUrl") or display_name,
                    display_name=display_name,
                    docx_name=Path(display_name).with_suffix(".docx").name,
                    remote_drive_id=_item_drive_id(item),
                    remote_item_id=str(item["id"]),
                )
            )
            if len(results) <= 5 or len(results) % 25 == 0:
                print(f"Discovered {len(results)} image(s)... latest: {display_name}")
        return

    if "folder" not in item:
        return

    children = sorted(
        _iter_drive_children(client, _item_drive_id(item), str(item["id"])),
        key=lambda child: (child.get("name") or "").lower(),
    )
    for child in children:
        _walk_remote_images(client, child, path_parts=current_parts, results=results, limit=limit)
        if limit is not None and len(results) >= limit:
            return


def _iter_remote_screenshots(client, shared_url: str, limit: int | None) -> list[ScreenshotSource]:
    root_item = _resolve_shared_item(client, shared_url)
    print("Listing images from shared location...")
    results: list[ScreenshotSource] = []
    _walk_remote_images(client, root_item, path_parts=[], results=results, limit=limit)
    if not results:
        raise SystemExit(f"ERROR: no supported screenshots found in shared location: {shared_url}")
    print(f"Finished listing shared location. Found {len(results)} supported image(s).")
    return results


def _download_remote_image_bytes(client, source: ScreenshotSource) -> bytes:
    if not source.is_remote:
        raise ValueError("Remote download requested for a local screenshot source.")
    print(f"Downloading image: {source.display_name}")
    return _onedrive_request(
        client,
        "GET",
        f"/drives/{quote(source.remote_drive_id or '', safe='')}/items/{quote(source.remote_item_id or '', safe='')}/content",
        expect_json=False,
        ok_statuses={200},
    )


def _ensure_target_table(engine, table_name: str) -> None:
    if table_name != "profiles":
        raise SystemExit('ERROR: complete_process.py now uploads to the existing "profiles" table.')
    with engine.connect() as conn:
        row = conn.execute(uploader._text("""
            SELECT to_regclass('public.profiles')
        """)).scalar()
        if row is None:
            raise SystemExit('ERROR: Neon database is missing public.profiles.')

        columns = set(conn.execute(uploader._text("""
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = 'profiles'
        """)).scalars().all())
    missing = sorted(PROFILE_COLUMNS - columns)
    if missing:
        raise SystemExit("ERROR: profiles table is missing column(s): " + ", ".join(missing))


def _existing_candidate_id(conn, table_name: str, fields: dict, digest: str) -> str | None:
    table = _quoted_table(table_name)
    row = conn.execute(
        uploader._text(f"""
            SELECT profile_id
            FROM {table}
            WHERE resume_sections::text LIKE :digest_like
               OR (CAST(:npi_number AS TEXT) IS NOT NULL AND npi_number = CAST(:npi_number AS TEXT))
               OR (CAST(:email AS TEXT) IS NOT NULL AND lower(email) = lower(CAST(:email AS TEXT)))
               OR (
                    CAST(:phone AS TEXT) IS NOT NULL
                    AND CAST(:last_name AS TEXT) IS NOT NULL
                    AND regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g')
                        LIKE '%' || regexp_replace(CAST(:phone AS TEXT), '[^0-9]', '', 'g')
                    AND lower(last_name) = lower(CAST(:last_name AS TEXT))
               )
            LIMIT 1
        """),
        {
            "digest": digest,
            "digest_like": f"%{digest}%",
            "npi_number": fields.get("npi_number"),
            "email": fields.get("email"),
            "phone": fields.get("phone"),
            "last_name": fields.get("last_name"),
        },
    ).first()
    return str(row[0]) if row else None


def _upsert_candidate(
    conn,
    *,
    table_name: str,
    row_id: str | None,
    fields: dict,
    text: str,
    source_file: str,
    docx_file: str,
    digest: str,
    key: str,
    resume_url: str,
) -> str:
    table = _quoted_table(table_name)
    now = uploader._utcnow()
    profile_id = row_id or str(uuid.uuid4())
    resume_sections = {
        "raw_extracted_text": text,
        "source_file": source_file,
        "docx_file": docx_file,
        "docx_sha256": digest,
        "cloudflare_bucket": os.environ.get("S3_BUCKET"),
        "cloudflare_key": key,
    }
    params = {
        "profile_id": profile_id,
        "first_name": fields.get("first_name"),
        "last_name": fields.get("last_name"),
        "phone": fields.get("phone"),
        "email": fields.get("email"),
        "headline": fields.get("headline"),
        "bio": fields.get("bio"),
        "profession_type": fields.get("profession_type"),
        "specialty": fields.get("specialty"),
        "provider_category": fields.get("provider_category"),
        "american_board": fields.get("american_board"),
        "city": fields.get("city"),
        "state_code": fields.get("state_code"),
        "npi_number": fields.get("npi_number"),
        "years_experience": fields.get("years_experience") or 0,
        "resume_url": resume_url,
        "completion_score": uploader._completion(fields),
        "search_text": uploader._search_text(fields),
        "job_type_prefs": _json_value(fields.get("job_type_prefs")),
        "resume_sections": json.dumps(resume_sections),
        "education": _json_value(fields.get("education")),
        "source": "resume_parse",
        "capture_source": "complete_process",
        "is_listable": not uploader._bad_name(fields.get("first_name"), fields.get("last_name")),
        "zip_code": fields.get("zip_code"),
        "created_at": now,
        "updated_at": now,
        "captured_at": datetime.now(timezone.utc),
    }
    if row_id:
        conn.execute(uploader._text(f"""
            UPDATE {table}
            SET first_name = :first_name,
                last_name = :last_name,
                phone = :phone,
                email = :email,
                headline = :headline,
                bio = :bio,
                profession_type = :profession_type,
                specialty = :specialty,
                provider_category = :provider_category,
                american_board = :american_board,
                city = :city,
                state_code = :state_code,
                npi_number = :npi_number,
                years_experience = :years_experience,
                resume_url = :resume_url,
                completion_score = :completion_score,
                search_text = :search_text,
                job_type_prefs = CAST(:job_type_prefs AS JSON),
                resume_sections = CAST(:resume_sections AS JSON),
                education = CAST(:education AS JSON),
                capture_source = :capture_source,
                captured_at = :captured_at,
                is_listable = :is_listable,
                zip_code = :zip_code,
                updated_at = :updated_at
            WHERE profile_id = :profile_id
        """), params)
        return row_id

    inserted = conn.execute(uploader._text(f"""
        INSERT INTO {table} (
            profile_id, user_id, first_name, last_name, headline, bio, phone, email,
            specialty, profession_type, provider_category, american_board,
            years_experience, city, state_code, lat, lng, open_to_work,
            job_type_prefs, pay_min_hourly, available_date, npi_number,
            profile_photo_url, resume_url, completion_score, source, search_text,
            is_listable, zip_code, resume_sections, education, capture_source,
            captured_at, created_at, updated_at
        )
        VALUES (
            :profile_id, NULL, :first_name, :last_name, :headline, :bio, :phone, :email,
            :specialty, :profession_type, :provider_category, :american_board,
            :years_experience, :city, :state_code, NULL, NULL, TRUE,
            CAST(:job_type_prefs AS JSON), NULL, NULL, :npi_number,
            NULL, :resume_url, :completion_score, CAST(:source AS profilesource), :search_text,
            :is_listable, :zip_code, CAST(:resume_sections AS JSON), CAST(:education AS JSON),
            :capture_source, :captured_at, :created_at, :updated_at
        )
        RETURNING profile_id
    """), params).first()
    return str(inserted[0])


def _process_one(
    *,
    source: ScreenshotSource,
    output_dir: Path | None,
    onedrive_client,
    engine,
    args: argparse.Namespace,
) -> dict:
    docx_parse_path = Path(source.docx_name)
    docx_reference = str(_docx_path_for(source.local_path, output_dir)) if source.local_path and output_dir else _virtual_docx_reference(source)
    base = {
        "record_id": source.record_id,
        "source_file": source.source_file,
        "docx_file": docx_reference,
        "updated_at": _utc_iso(),
    }

    if source.is_remote:
        image_bytes = _download_remote_image_bytes(onedrive_client, source)
        print(f"Running OCR and DOCX conversion: {source.display_name}")
        docx_bytes = converter.convert_bytes(
            image_bytes,
            source.display_name,
            scale=args.scale,
            cutoff_ratio=args.cutoff_ratio,
            keep_promo=args.keep_promo,
            min_conf=args.min_conf,
            debug=args.debug,
        )
        digest = uploader._sha256_bytes(docx_bytes)
        text = uploader.extract_docx_text_bytes(docx_bytes)
    else:
        if output_dir is None or source.local_path is None:
            raise RuntimeError("Local processing requires a concrete output directory and source path.")
        docx_path = _docx_path_for(source.local_path, output_dir)
        base["docx_file"] = str(docx_path)
        if not docx_path.exists() or args.force_convert:
            converter.convert(
                source.local_path,
                docx_path,
                scale=args.scale,
                cutoff_ratio=args.cutoff_ratio,
                keep_promo=args.keep_promo,
                min_conf=args.min_conf,
                debug=args.debug,
            )
        digest = uploader._sha256_file(docx_path)
        text = uploader.extract_text(docx_path)
        docx_bytes = None

    fields = uploader.parse_resume_smart(text, docx_parse_path)
    label = f"{fields['first_name']} {fields['last_name']}"
    base["label"] = label
    key = _cloudflare_key_for(
        docx_parse_path,
        digest,
        fields.get("provider_category"),
        fields.get("profession_type"),
        fields.get("specialty"),
    )
    resume_url = uploader._url_for_key(key)
    base.update({
        "sha256": digest,
        "docx_sha256": digest,
        "cloudflare_key": key,
        "resume_url": resume_url,
        "specialty": fields.get("specialty") or "Unknown",
        "first_name": fields.get("first_name"),
        "last_name": fields.get("last_name"),
        "email": fields.get("email"),
        "phone": fields.get("phone"),
        "npi_number": fields.get("npi_number"),
        "profession_type": fields.get("profession_type"),
        "provider_category": fields.get("provider_category"),
        "city": fields.get("city"),
        "state_code": fields.get("state_code"),
    })

    if args.dry_run:
        return {**base, "status": "would_process"}

    print(f"Checking existing profile match: {source.display_name}")
    with engine.begin() as conn:
        candidate_id = _existing_candidate_id(conn, args.target_table, fields, digest)

    if source.is_remote:
        print(f"Uploading DOCX to Cloudflare R2: {source.display_name}")
        uploaded_url = uploader._upload_bytes_if_needed(docx_bytes or b"", key, content_type=DOCX_CONTENT_TYPE)
    else:
        uploaded_url = uploader._upload_if_needed(Path(base["docx_file"]), key)

    print(f"Upserting profile in Neon: {source.display_name}")
    with engine.begin() as conn:
        candidate_id = _upsert_candidate(
            conn,
            table_name=args.target_table,
            row_id=candidate_id,
            fields=fields,
            text=text,
            source_file=source.source_file,
            docx_file=base["docx_file"],
            digest=digest,
            key=key,
            resume_url=uploaded_url,
        )

    return {
        **base,
        "status": "processed",
        "profile_id": candidate_id,
        "candidate_id": candidate_id,
        "resume_url": uploaded_url,
        "uploaded_at": _utc_iso(),
    }


def _build_onedrive_client(args: argparse.Namespace):
    if not args.onedrive_client_id:
        raise SystemExit(
            "ERROR: OneDrive/SharePoint URL inputs require --onedrive-client-id "
            "or the ONEDRIVE_CLIENT_ID environment variable."
        )

    if args.onedrive_client_secret:
        client = ConfidentialOneDriveClient(
            client_id=args.onedrive_client_id,
            tenant=args.onedrive_tenant,
            client_secret=args.onedrive_client_secret,
        )
    else:
        client = converter.OneDriveUploader(
            client_id=args.onedrive_client_id,
            tenant=args.onedrive_tenant,
            remote_folder="",
            upload_delay=0.0,
            open_browser=not args.no_browser,
        )

    try:
        client.authenticate()
    except converter.OneDriveError as exc:
        hint = ""
        error_text = str(exc)
        if "AADSTS50059" in error_text:
            hint = (
                "\nHint: Try --onedrive-tenant common or your Microsoft Entra tenant ID, "
                "and confirm the app registration allows public client flows."
            )
        elif "AADSTS7000218" in error_text:
            hint = (
                "\nHint: pass --onedrive-client-secret with the secret value, "
                "or enable public client flows and retry without a secret."
            )
        raise SystemExit(f"OneDrive login failed: {exc}{hint}") from exc
    return client


def run(args: argparse.Namespace) -> dict:
    uploader._prepare_runtime_env(args)
    uploader._validate_upload_env(args)

    remote_input = _is_remote_input(args.input)
    manifest_path = _optional_local_path(args.manifest, None if remote_input else DEFAULT_PROCESS_MANIFEST)
    upload_log_path = _optional_local_path(args.upload_log, None if remote_input else DEFAULT_UPLOAD_LOG)

    if remote_input:
        if args.output_dir:
            raise SystemExit(
                "ERROR: --output-dir is not supported for OneDrive/SharePoint URL inputs; "
                "DOCX files are streamed directly without local staging."
            )
        onedrive_client = _build_onedrive_client(args)
        screenshots = _iter_remote_screenshots(onedrive_client, args.input, args.limit)
        output_dir = None
        print("Input mode: OneDrive/SharePoint shared URL")
        if manifest_path is None:
            print("Local manifest disabled for remote input.")
        if upload_log_path is None:
            print("Local upload log disabled for remote input.")
    else:
        input_path = Path(args.input).expanduser()
        output_dir = Path(args.output_dir).expanduser() if args.output_dir else _default_output_dir_for(input_path)
        output_dir.mkdir(parents=True, exist_ok=True)
        onedrive_client = None
        screenshots = _iter_local_screenshots(input_path, args.limit)

    latest = {} if args.ignore_manifest else _latest_manifest_status(manifest_path)
    terminal_statuses = {"processed", "skipped_duplicate"}
    required = [
        source for source in screenshots
        if args.retry_all or latest.get(source.record_id, {}).get("status") not in terminal_statuses
    ]

    print(f"Found {len(screenshots)} screenshot(s); {len(required)} require processing.")
    if args.dry_run:
        print("Dry run: conversion and parsing only; no database writes or Cloudflare uploads.")

    needs_conversion = bool(required) and (
        remote_input
        or any(
            source.local_path is not None
            and output_dir is not None
            and not (_docx_path_for(source.local_path, output_dir).exists() and not args.force_convert)
            for source in required
        )
    )
    if needs_conversion:
        converter.configure_tesseract(args.tesseract)

    engine = None if args.dry_run else uploader._create_engine()
    if engine is not None:
        _ensure_target_table(engine, args.target_table)
        print(f'Target table: "{args.target_table}"')

    stats = {"processed": 0, "skipped": 0, "failed": 0, "total": len(required)}
    started = monotonic()
    worker_count = max(1, min(args.workers, len(required) or 1))
    if required:
        print(f"Starting processing with {worker_count} worker(s)...")

    def _docx_reference_for(source: ScreenshotSource) -> str:
        return (
            str(_docx_path_for(source.local_path, output_dir))
            if source.local_path is not None and output_dir is not None
            else _virtual_docx_reference(source)
        )

    def _handle_success(index: int, source: ScreenshotSource, record: dict) -> None:
        if manifest_path is not None:
            _append_manifest(manifest_path, record)
        status = record["status"]
        if status in {"processed", "would_process"}:
            stats["processed"] += 1
            if status == "processed" and upload_log_path is not None:
                _append_upload_log(upload_log_path, record)
            profile_note = f" -> {record.get('profile_id')}" if record.get("profile_id") else ""
            print(f"[{index}/{len(required)}] {status.upper()} {record.get('label', _label_from_source(source))}{profile_note}")
        else:
            stats["skipped"] += 1
            print(f"[{index}/{len(required)}] SKIP {_label_from_source(source)}: {status}")

    def _handle_failure(index: int, source: ScreenshotSource, docx_reference: str, exc: Exception) -> None:
        stats["failed"] += 1
        failed_record = {
            "record_id": source.record_id,
            "source_file": source.source_file,
            "docx_file": docx_reference,
            "status": "failed",
            "error": str(exc),
            "updated_at": _utc_iso(),
            "uploaded_at": _utc_iso(),
        }
        if manifest_path is not None:
            _append_manifest(manifest_path, failed_record)
        if upload_log_path is not None:
            _append_upload_log(upload_log_path, failed_record)
        print(f"[{index}/{len(required)}] FAIL {_label_from_source(source)}: {exc}", file=sys.stderr)

    if worker_count == 1:
        for index, source in enumerate(required, start=1):
            docx_reference = _docx_reference_for(source)
            try:
                record = _process_one(
                    source=source,
                    output_dir=output_dir,
                    onedrive_client=onedrive_client,
                    engine=engine,
                    args=args,
                )
                _handle_success(index, source, record)
            except Exception as exc:
                _handle_failure(index, source, docx_reference, exc)
    else:
        future_map = {}
        with ThreadPoolExecutor(max_workers=worker_count, thread_name_prefix="process-worker") as executor:
            for index, source in enumerate(required, start=1):
                docx_reference = _docx_reference_for(source)
                future = executor.submit(
                    _process_one,
                    source=source,
                    output_dir=output_dir,
                    onedrive_client=onedrive_client,
                    engine=engine,
                    args=args,
                )
                future_map[future] = (index, source, docx_reference)

            for future in as_completed(future_map):
                index, source, docx_reference = future_map[future]
                try:
                    record = future.result()
                    _handle_success(index, source, record)
                except Exception as exc:
                    _handle_failure(index, source, docx_reference, exc)

    elapsed = max(monotonic() - started, 0.001)
    print(f"\nSummary: {stats}")
    print(f"Elapsed: {elapsed / 60:.1f} min ({len(required) / elapsed:.2f} records/sec)")
    return stats


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convert screenshots to DOCX, insert profile data, upload DOCX to Cloudflare R2, and track retry status."
    )
    parser.add_argument("input", help="Local screenshot image/folder, or a OneDrive/SharePoint shared folder URL")
    parser.add_argument("--output-dir", help=f"Folder for generated DOCX files for local inputs only; default: {DEFAULT_OUTPUT_FOLDER_NAME}")
    parser.add_argument("--dry-run", action="store_true", help="Convert/parse only; do not write DB or upload")
    parser.add_argument("--limit", type=int, default=None, help="Only process first N screenshots")
    parser.add_argument("--prefix", default=uploader.DEFAULT_PREFIX, help="Cloudflare R2 key prefix")
    parser.add_argument("--target-table", default=DEFAULT_TARGET_TABLE, help='Neon table for uploaded profiles, default "profiles"')
    parser.add_argument("--manifest", default=None, help='JSONL process status manifest path. Use "none" to disable; disabled by default for OneDrive/SharePoint URL inputs.')
    parser.add_argument("--upload-log", default=None, help='CSV log path. Use "none" to disable; disabled by default for OneDrive/SharePoint URL inputs.')
    parser.add_argument("--ignore-manifest", action="store_true", help="Do not skip records already marked processed")
    parser.add_argument("--retry-all", action="store_true", help="Retry every discovered record, including previously processed records")
    parser.add_argument("--force-convert", action="store_true", help="Regenerate DOCX even if it already exists")
    parser.add_argument("--env-file", default=DEFAULT_ENV_FILE, help="Env file path, relative to this script by default")
    parser.add_argument("--no-save-credentials", action="store_true", help="Do not offer to save prompted credentials")
    parser.add_argument("--install-deps", action="store_true", help="Install missing Python packages, then continue")
    parser.add_argument("--tesseract", help="Path to tesseract.exe")
    parser.add_argument("--scale", type=float, default=1.5, help="Upscale factor before OCR")
    parser.add_argument("--cutoff-ratio", type=float, default=0.66, help="Fallback main-column width fraction")
    parser.add_argument("--min-conf", type=int, default=35, help="Minimum OCR confidence")
    parser.add_argument("--keep-promo", action="store_true", help="Keep Doximity promo/navigation lines")
    parser.add_argument("--workers", type=int, default=1, help="Number of parallel workers for download/OCR/upload processing")
    parser.add_argument("--onedrive-client-id", default=os.environ.get("ONEDRIVE_CLIENT_ID"), help="Microsoft Graph client ID for OneDrive/SharePoint URL inputs")
    parser.add_argument("--onedrive-client-secret", default=DEFAULT_ONEDRIVE_CLIENT_SECRET, help="Microsoft Graph client secret value for app-only OneDrive/SharePoint URL access")
    parser.add_argument("--onedrive-tenant", default=DEFAULT_ONEDRIVE_TENANT, help="Microsoft tenant for OneDrive/SharePoint URL inputs. Use common or your tenant ID.")
    parser.add_argument("--no-browser", action="store_true", help="Print the Microsoft device login URL/code without trying to open a browser")
    parser.add_argument("--debug", action="store_true", help="Print converter structure details")
    args = parser.parse_args(argv)

    if args.workers < 1:
        raise SystemExit("ERROR: --workers must be 1 or greater.")

    if args.install_deps:
        uploader._install_missing_dependencies()

    stats = run(args)
    return 1 if stats["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
