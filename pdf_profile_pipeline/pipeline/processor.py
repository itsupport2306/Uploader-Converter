"""Orchestration: SharePoint PDF -> extract -> profile ID -> R2 -> Neon.

The PDF stays a PDF end to end. Nothing is converted to images or DOCX.
"""
from __future__ import annotations

import argparse
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from time import monotonic

from . import config, database, experience, llm_extract, pdf_text, run_log, sharepoint, storage
from .sharepoint import PdfSource

TERMINAL_STATUSES = {"processed", "skipped_duplicate"}


def _extract_fields(text: str, *, use_llm: bool) -> tuple[dict, str, str | None]:
    """Return (fields, extraction_mode, llm_error)."""
    raw: dict = {}
    mode = "regex"
    llm_error = None

    if use_llm:
        try:
            raw = llm_extract.extract_profile(text)
            mode = f"qwen:{llm_extract.model_name()}"
        except llm_extract.LlmError as exc:
            llm_error = str(exc)
            mode = "regex_fallback"
            print(f"  warning: model extraction failed, falling back to regex: {exc}")

    fields = llm_extract.normalize_profile(raw, text)
    years, detail = experience.resolve_years_experience(
        text,
        model_total=fields.pop("_model_total_years", None),
        model_positions=fields.pop("_model_positions", []),
    )
    fields["years_experience"] = years
    fields["_experience_detail"] = detail
    return fields, mode, llm_error


def process_one(
    *,
    source: PdfSource,
    graph_client,
    engine,
    args: argparse.Namespace,
) -> dict:
    record = {
        "record_id": source.record_id,
        "source_file": source.source_file,
        "display_name": source.display_name,
        "file_name": source.file_name,
        "logged_at": run_log.utc_iso(),
    }

    pdf_bytes = sharepoint.read_pdf_bytes(graph_client, source)
    digest = storage.sha256_bytes(pdf_bytes)
    pages = pdf_text.page_count(pdf_bytes)
    text = pdf_text.extract_text(pdf_bytes)

    record.update({"pdf_sha256": digest, "pdf_pages": pages, "pdf_bytes": len(pdf_bytes)})

    if pdf_text.looks_empty(text):
        # Image-only PDF: no text layer to read, and converting it is out of scope.
        return {
            **record,
            "status": "skipped_no_text",
            "error": "PDF has no extractable text layer (likely a scanned image PDF).",
        }

    fields, mode, llm_error = _extract_fields(text, use_llm=args.use_llm)
    experience_detail = fields.pop("_experience_detail", {})
    profile_id = database.profile_id_for(digest)

    key = storage.key_for(digest, fields.get("specialty"), fields.get("state_code"))
    record.update({
        "profile_id": profile_id,
        "first_name": fields.get("first_name"),
        "last_name": fields.get("last_name"),
        "full_name": fields.get("full_name"),
        "city": fields.get("city"),
        "state_code": fields.get("state_code"),
        "specialty": fields.get("specialty"),
        "years_experience": fields.get("years_experience"),
        "years_experience_source": experience_detail.get("chosen_kind"),
        "years_experience_detail": experience_detail,
        "bio_chars": len(fields.get("bio") or ""),
        "extraction_mode": mode,
        "llm_error": llm_error,
        "cloudflare_key": key,
        "resume_url": storage.url_for_key(key),
    })

    if args.dry_run:
        return {**record, "status": "would_process"}

    print(f"  uploading original PDF to Cloudflare R2: {key}")
    resume_url, uploaded_now = storage.upload_pdf(pdf_bytes, key)
    record.update({"resume_url": resume_url, "cloudflare_uploaded": uploaded_now})

    resume_sections = {
        "raw_extracted_text": text,
        "source_file": source.source_file,
        "source_display_name": source.display_name,
        "pdf_file_name": source.file_name,
        "pdf_sha256": digest,
        "pdf_pages": pages,
        "cloudflare_bucket": os.environ.get("S3_BUCKET"),
        "cloudflare_key": key,
        "extraction_mode": mode,
        "experience_decision": experience_detail,
        "full_name": fields.get("full_name"),
    }

    with engine.begin() as conn:
        existing_id = database.find_existing_profile_id(
            conn, args.target_table, profile_id=profile_id, digest=digest, fields=fields,
        )
        stored_id, db_action = database.upsert_profile(
            conn,
            table_name=args.target_table,
            profile_id=profile_id,
            existing_id=existing_id,
            fields=fields,
            resume_sections=resume_sections,
            resume_url=resume_url,
        )

    return {
        **record,
        "status": "processed",
        "profile_id": stored_id,
        "db_action": db_action,
        "uploaded_at": run_log.utc_iso(),
    }


def _label(source: PdfSource) -> str:
    return Path(source.display_name).name or source.display_name


def discover(args: argparse.Namespace) -> tuple[list[PdfSource], object]:
    if sharepoint.is_remote_input(args.input):
        client = sharepoint.build_client(
            client_id=args.onedrive_client_id,
            tenant=args.onedrive_tenant,
            client_secret=args.onedrive_client_secret,
            open_browser=not args.no_browser,
        )
        print("Input mode: SharePoint/OneDrive shared URL")
        return sharepoint.iter_remote_pdfs(client, args.input, args.limit), client

    print("Input mode: local PDF file/folder")
    return sharepoint.iter_local_pdfs(Path(args.input).expanduser(), args.limit), None


def run(args: argparse.Namespace) -> dict:
    env_file = config.load_env(args.env_file)
    print(f"Env file: {env_file}{'' if env_file.exists() else ' (not found; using process environment)'}")

    args.use_llm = llm_extract.llm_enabled() and not args.no_llm
    config.validate_env(dry_run=args.dry_run, use_llm=args.use_llm)

    if args.use_llm:
        print(f"Extraction model: {llm_extract.model_name()} @ {os.environ.get('LLM_BASE_URL')}")
    else:
        print("Extraction model: disabled (regex-only extraction)")
    if not args.dry_run:
        print(f"Database: {config.safe_db_label(config.normalize_db_url(os.environ.get('DATABASE_URL', '')))}")
        print(f"Cloudflare R2 bucket: {os.environ.get('S3_BUCKET')}")

    manifest_path = run_log.optional_path(args.manifest, config.ROOT / "logs" / "profiles_manifest.jsonl")
    csv_path = run_log.optional_path(args.upload_log, config.ROOT / "logs" / "upload_log.csv")

    sources, graph_client = discover(args)

    latest = {} if args.ignore_manifest else run_log.latest_status(manifest_path)
    required = [
        source for source in sources
        if args.retry_all or latest.get(source.record_id, {}).get("status") not in TERMINAL_STATUSES
    ]
    print(f"Found {len(sources)} PDF(s); {len(required)} require processing.")
    if args.dry_run:
        print("Dry run: read and extract only. No Cloudflare upload, no database write.")

    engine = None
    if not args.dry_run:
        engine = database.create_engine()
        database.ensure_table(engine, args.target_table)
        print(f'Target table: "{args.target_table}"')

    stats = {"processed": 0, "skipped": 0, "failed": 0, "total": len(required)}
    started = monotonic()
    workers = max(1, min(args.workers, len(required) or 1))
    if required:
        print(f"Starting processing with {workers} worker(s)...")

    def _log(record: dict) -> None:
        if manifest_path is not None:
            run_log.append_manifest(manifest_path, record)
        if csv_path is not None:
            run_log.append_csv(csv_path, record)

    def _on_success(index: int, source: PdfSource, record: dict) -> None:
        status = record["status"]
        _log(record)
        if status in {"processed", "would_process"}:
            stats["processed"] += 1
            note = f" -> {record.get('profile_id')} ({record.get('db_action', 'dry-run')})"
            years = record.get("years_experience")
            print(f"[{index}/{len(required)}] {status.upper()} {_label(source)}{note} "
                  f"| yrs={years} via {record.get('years_experience_source')}")
        else:
            stats["skipped"] += 1
            print(f"[{index}/{len(required)}] SKIP {_label(source)}: {status} - {record.get('error', '')}")

    def _on_failure(index: int, source: PdfSource, exc: Exception) -> None:
        stats["failed"] += 1
        record = {
            "record_id": source.record_id,
            "source_file": source.source_file,
            "display_name": source.display_name,
            "status": "failed",
            "error": str(exc),
            "logged_at": run_log.utc_iso(),
        }
        _log(record)
        print(f"[{index}/{len(required)}] FAIL {_label(source)}: {exc}")

    if workers == 1:
        for index, source in enumerate(required, start=1):
            print(f"[{index}/{len(required)}] Reading {_label(source)}")
            try:
                _on_success(index, source, process_one(
                    source=source, graph_client=graph_client, engine=engine, args=args,
                ))
            except Exception as exc:
                _on_failure(index, source, exc)
    else:
        futures = {}
        with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="pdf-worker") as executor:
            for index, source in enumerate(required, start=1):
                futures[executor.submit(
                    process_one, source=source, graph_client=graph_client, engine=engine, args=args,
                )] = (index, source)
            for future in as_completed(futures):
                index, source = futures[future]
                try:
                    _on_success(index, source, future.result())
                except Exception as exc:
                    _on_failure(index, source, exc)

    elapsed = max(monotonic() - started, 0.001)
    print(f"\nSummary: {stats}")
    print(f"Elapsed: {elapsed / 60:.1f} min ({len(required) / elapsed:.2f} records/sec)")
    if manifest_path is not None:
        print(f"Manifest: {manifest_path}")
    if csv_path is not None:
        print(f"Upload log: {csv_path}")
    return stats
