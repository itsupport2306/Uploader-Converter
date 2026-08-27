"""CLI entry point for the SharePoint PDF -> Cloudflare R2 -> Neon pipeline.

    python run_pipeline.py "https://<sharepoint-share-link>"
    python run_pipeline.py "C:\\path\\to\\pdfs" --limit 5 --dry-run
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pipeline import config, processor, run_log  # noqa: E402


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Read PDFs from SharePoint/OneDrive (or a local folder), extract profile "
            "details with Qwen2.5, upload the original PDF to Cloudflare R2, and "
            "insert/update the profile in Neon."
        )
    )
    parser.add_argument("input", help="SharePoint/OneDrive shared folder URL, or a local PDF file/folder")
    parser.add_argument("--env-file", default=config.DEFAULT_ENV_FILE,
                        help="Env file path; relative paths resolve next to this script")
    parser.add_argument("--limit", type=int, default=None, help="Process only the first N PDFs")
    parser.add_argument("--workers", type=int, default=1, help="Parallel workers for download/extract/upload")
    parser.add_argument("--dry-run", action="store_true",
                        help="Read and extract only; no Cloudflare upload and no database write")
    parser.add_argument("--target-table", default=os.environ.get("TARGET_TABLE", "profiles"),
                        help='Neon table to write, default "profiles"')

    parser.add_argument("--manifest", default=None,
                        help='JSONL manifest path, or "none" to disable. Default: logs/profiles_manifest.jsonl')
    parser.add_argument("--upload-log", default=None,
                        help='CSV upload log path, or "none" to disable. Default: logs/upload_log.csv')
    parser.add_argument("--run-log", default=None,
                        help='Console transcript log path, or "none" to disable. Default: logs/run_<timestamp>.log')
    parser.add_argument("--ignore-manifest", action="store_true",
                        help="Do not skip records already marked processed")
    parser.add_argument("--retry-all", action="store_true",
                        help="Reprocess every discovered PDF, including previously processed ones")

    # OCR is on by default: a scanned resume must never be dropped for having
    # no text layer. --no-ocr is the opt-out.
    parser.add_argument("--ocr", action="store_true", default=True,
                        help="OCR scanned/image-only PDFs with Tesseract (default: on)")
    parser.add_argument("--no-ocr", dest="ocr", action="store_false",
                        help="Do not OCR scanned PDFs; use the text layer only")
    parser.add_argument("--require-ocr", action="store_true",
                        help="Fail the run if Tesseract cannot be started, instead of continuing without it")
    parser.add_argument("--no-llm", action="store_true",
                        help="Skip the Qwen2.5 call and use regex extraction only")

    parser.add_argument("--onedrive-client-id", default=os.environ.get("ONEDRIVE_CLIENT_ID"),
                        help="Microsoft Graph client ID (or set ONEDRIVE_CLIENT_ID)")
    parser.add_argument("--onedrive-client-secret", default=os.environ.get("ONEDRIVE_CLIENT_SECRET"),
                        help="Client secret for app-only access (or set ONEDRIVE_CLIENT_SECRET)")
    parser.add_argument("--onedrive-tenant", default=None,
                        help="Microsoft tenant: 'common' or your tenant ID (or set ONEDRIVE_TENANT)")
    parser.add_argument("--no-browser", action="store_true",
                        help="Print the device login URL/code instead of opening a browser")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.workers < 1:
        raise SystemExit("ERROR: --workers must be 1 or greater.")

    # The env file is loaded again inside run(); load it here so CLI defaults
    # that read the environment (client id, tenant, table) also pick it up.
    config.load_env(args.env_file)
    args.onedrive_client_id = args.onedrive_client_id or os.environ.get("ONEDRIVE_CLIENT_ID")
    args.onedrive_client_secret = args.onedrive_client_secret or os.environ.get("ONEDRIVE_CLIENT_SECRET")
    args.onedrive_tenant = args.onedrive_tenant or os.environ.get("ONEDRIVE_TENANT") or "common"

    default_run_log = config.ROOT / "logs" / f"run_{run_log.utc_iso().replace(':', '').replace('-', '')}.log"
    run_log_path = run_log.optional_path(args.run_log, default_run_log)

    with run_log.RunLog(run_log_path):
        if run_log_path is not None:
            print(f"Run log: {run_log_path}")
        stats = processor.run(args)
    return 1 if stats["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
