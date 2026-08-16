from __future__ import annotations

import argparse
import csv
import importlib.util
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from time import monotonic
import uuid


ROOT = Path(__file__).resolve().parent
CONVERTER_DIR = ROOT / "Converter" / "Converter" / "Converter"
UPLOAD_SCRIPT = ROOT / "data_upload 5.py"
DEFAULT_OUTPUT_DIR = ROOT / "generated_docx"
DEFAULT_PROCESS_MANIFEST = ROOT / "complete_process_profiles_manifest.jsonl"
DEFAULT_UPLOAD_LOG = ROOT / "complete_process_upload_log.csv"
DEFAULT_TARGET_TABLE = "profiles"
DEFAULT_ENV_FILE = "env 1" if (ROOT / "env 1").exists() else ".env"

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


def _latest_manifest_status(path: Path) -> dict[str, dict]:
    latest: dict[str, dict] = {}
    if not path.exists():
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


def _iter_screenshots(input_path: Path, limit: int | None) -> list[Path]:
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
    return files[:limit] if limit else files


def _record_id(path: Path) -> str:
    return str(path.resolve()).lower()


def _docx_path_for(image_path: Path, output_dir: Path) -> Path:
    return output_dir / f"{image_path.stem}.docx"


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
    image_path: Path,
    docx_path: Path,
    digest: str,
    key: str,
    resume_url: str,
) -> str:
    table = _quoted_table(table_name)
    now = uploader._utcnow()
    profile_id = row_id or str(uuid.uuid4())
    resume_sections = {
        "raw_extracted_text": text,
        "source_file": str(image_path),
        "docx_file": str(docx_path),
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
    image_path: Path,
    docx_path: Path,
    engine,
    seen_identity: dict,
    args: argparse.Namespace,
) -> dict:
    record_id = _record_id(image_path)
    base = {
        "record_id": record_id,
        "source_file": str(image_path),
        "docx_file": str(docx_path),
        "updated_at": _utc_iso(),
    }

    if not docx_path.exists() or args.force_convert:
        converter.convert(
            image_path,
            docx_path,
            scale=args.scale,
            cutoff_ratio=args.cutoff_ratio,
            keep_promo=args.keep_promo,
            min_conf=args.min_conf,
            debug=args.debug,
        )

    digest = uploader._sha256_file(docx_path)
    text = uploader.extract_text(docx_path)
    fields = uploader.parse_resume_smart(text, docx_path)
    label = f"{fields['first_name']} {fields['last_name']}"
    base["label"] = label
    key = _cloudflare_key_for(
        docx_path,
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

    with engine.begin() as conn:
        candidate_id = _existing_candidate_id(conn, args.target_table, fields, digest)

    uploaded_url = uploader._upload_if_needed(docx_path, key)

    with engine.begin() as conn:
        candidate_id = _upsert_candidate(
            conn,
            table_name=args.target_table,
            row_id=candidate_id,
            fields=fields,
            text=text,
            image_path=image_path,
            docx_path=docx_path,
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


def run(args: argparse.Namespace) -> dict:
    uploader._prepare_runtime_env(args)
    uploader._validate_upload_env(args)

    input_path = Path(args.input).expanduser()
    output_dir = Path(args.output_dir).expanduser()
    output_dir.mkdir(parents=True, exist_ok=True)

    screenshots = _iter_screenshots(input_path, args.limit)
    latest = {} if args.ignore_manifest else _latest_manifest_status(Path(args.manifest))
    terminal_statuses = {"processed", "skipped_duplicate"}
    required = [
        p for p in screenshots
        if args.retry_all or latest.get(_record_id(p), {}).get("status") not in terminal_statuses
    ]

    print(f"Found {len(screenshots)} screenshot(s); {len(required)} require processing.")
    if args.dry_run:
        print("Dry run: conversion and parsing only; no database writes or Cloudflare uploads.")

    if any(not (_docx_path_for(p, output_dir).exists() and not args.force_convert) for p in required):
        converter.configure_tesseract(args.tesseract)

    engine = None if args.dry_run else uploader._create_engine()
    if engine is not None:
        _ensure_target_table(engine, args.target_table)
        seen_identity = {"npi": set(), "email": set(), "phone": set()}
        print(f'Target table: "{args.target_table}"')
    else:
        seen_identity = {"npi": set(), "email": set(), "phone": set()}

    stats = {"processed": 0, "skipped": 0, "failed": 0, "total": len(required)}
    started = monotonic()
    manifest_path = Path(args.manifest)
    upload_log_path = Path(args.upload_log)

    for index, image_path in enumerate(required, start=1):
        docx_path = _docx_path_for(image_path, output_dir)
        try:
            record = _process_one(
                image_path=image_path,
                docx_path=docx_path,
                engine=engine,
                seen_identity=seen_identity,
                args=args,
            )
            _append_manifest(manifest_path, record)
            status = record["status"]
            if status in {"processed", "would_process"}:
                stats["processed"] += 1
                if status == "processed":
                    _append_upload_log(upload_log_path, record)
                profile_note = f" -> {record.get('profile_id')}" if record.get("profile_id") else ""
                print(f"[{index}/{len(required)}] {status.upper()} {record.get('label', image_path.stem)}{profile_note}")
            else:
                stats["skipped"] += 1
                print(f"[{index}/{len(required)}] SKIP {image_path.name}: {status}")
        except Exception as exc:
            stats["failed"] += 1
            failed_record = {
                "record_id": _record_id(image_path),
                "source_file": str(image_path),
                "docx_file": str(docx_path),
                "status": "failed",
                "error": str(exc),
                "updated_at": _utc_iso(),
                "uploaded_at": _utc_iso(),
            }
            _append_manifest(manifest_path, failed_record)
            _append_upload_log(upload_log_path, failed_record)
            print(f"[{index}/{len(required)}] FAIL {image_path.name}: {exc}", file=sys.stderr)

    elapsed = max(monotonic() - started, 0.001)
    print(f"\nSummary: {stats}")
    print(f"Elapsed: {elapsed / 60:.1f} min ({len(required) / elapsed:.2f} records/sec)")
    return stats


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convert screenshots to DOCX, insert profile data, upload DOCX to Cloudflare R2, and track retry status."
    )
    parser.add_argument("input", help="Screenshot image or folder of screenshot images")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="Folder for generated DOCX files")
    parser.add_argument("--dry-run", action="store_true", help="Convert/parse only; do not write DB or upload")
    parser.add_argument("--limit", type=int, default=None, help="Only process first N screenshots")
    parser.add_argument("--prefix", default=uploader.DEFAULT_PREFIX, help="Cloudflare R2 key prefix")
    parser.add_argument("--target-table", default=DEFAULT_TARGET_TABLE, help='Neon table for uploaded profiles, default "profiles"')
    parser.add_argument("--manifest", default=str(DEFAULT_PROCESS_MANIFEST), help="JSONL process status manifest")
    parser.add_argument("--upload-log", default=str(DEFAULT_UPLOAD_LOG), help="CSV log of uploaded profiles and lookup fields")
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
    parser.add_argument("--debug", action="store_true", help="Print converter structure details")
    args = parser.parse_args(argv)

    if args.install_deps:
        uploader._install_missing_dependencies()

    stats = run(args)
    return 1 if stats["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
