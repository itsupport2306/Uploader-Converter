from __future__ import annotations

import argparse
import importlib.util
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from time import monotonic


ROOT = Path(__file__).resolve().parent
CONVERTER_DIR = ROOT / "Converter" / "Converter" / "Converter"
UPLOAD_SCRIPT = ROOT / "data_upload 5.py"
DEFAULT_OUTPUT_DIR = ROOT / "generated_docx"
DEFAULT_PROCESS_MANIFEST = ROOT / "complete_process_manifest.jsonl"


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


def _set_profile_resume_url(conn, profile_id: str, resume_url: str) -> None:
    conn.execute(
        uploader._text("""
            UPDATE profiles
            SET resume_url = :resume_url, updated_at = :updated_at
            WHERE profile_id = :profile_id
        """),
        {
            "profile_id": profile_id,
            "resume_url": resume_url,
            "updated_at": uploader._utcnow(),
        },
    )


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
    key = uploader._key_for(docx_path, digest, args.prefix)
    resume_url = uploader._url_for_key(key)
    base.update({"sha256": digest, "cloudflare_key": key, "resume_url": resume_url})

    text = uploader.extract_text(docx_path)
    fields = uploader.parse_resume_smart(text, docx_path)
    label = f"{fields['first_name']} {fields['last_name']}"
    base["label"] = label

    if args.dry_run:
        return {**base, "status": "would_process"}

    with engine.begin() as conn:
        existing_id = uploader._existing_profile_id(conn, key)
        if existing_id:
            profile_id = existing_id
        else:
            dup_by = uploader._identity_dup(seen_identity, fields)
            if dup_by:
                return {**base, "status": "skipped_duplicate", "duplicate_by": dup_by}
            profile_id = uploader._insert_profile(conn, docx_path, key, resume_url, fields)
            uploader._identity_remember(seen_identity, fields)

    uploaded_url = uploader._upload_if_needed(docx_path, key)

    with engine.begin() as conn:
        _set_profile_resume_url(conn, profile_id, uploaded_url)

    return {
        **base,
        "status": "processed",
        "profile_id": profile_id,
        "resume_url": uploaded_url,
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
        uploader._validate_schema(engine)
        seen_identity = uploader._load_existing_identity(engine)
        print(f"Identity index: {len(seen_identity['email']):,} emails, "
              f"{len(seen_identity['phone']):,} phones, {len(seen_identity['npi']):,} NPIs.")
    else:
        seen_identity = {"npi": set(), "email": set(), "phone": set()}

    stats = {"processed": 0, "skipped": 0, "failed": 0, "total": len(required)}
    started = monotonic()
    manifest_path = Path(args.manifest)

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
                print(f"[{index}/{len(required)}] {status.upper()} {record.get('label', image_path.stem)}")
            else:
                stats["skipped"] += 1
                print(f"[{index}/{len(required)}] SKIP {image_path.name}: {status}")
        except Exception as exc:
            stats["failed"] += 1
            _append_manifest(manifest_path, {
                "record_id": _record_id(image_path),
                "source_file": str(image_path),
                "docx_file": str(docx_path),
                "status": "failed",
                "error": str(exc),
                "updated_at": _utc_iso(),
            })
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
    parser.add_argument("--manifest", default=str(DEFAULT_PROCESS_MANIFEST), help="JSONL process status manifest")
    parser.add_argument("--ignore-manifest", action="store_true", help="Do not skip records already marked processed")
    parser.add_argument("--retry-all", action="store_true", help="Retry every discovered record, including previously processed records")
    parser.add_argument("--force-convert", action="store_true", help="Regenerate DOCX even if it already exists")
    parser.add_argument("--env-file", default=".env", help="Env file path, relative to this script by default")
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
