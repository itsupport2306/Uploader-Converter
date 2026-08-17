from __future__ import annotations

import argparse
import importlib.util
import re
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
UPLOAD_SCRIPT = SCRIPT_DIR / "data_upload 5.py"

# Default run configuration. Running this file with no arguments uses these
# values and uploads to Cloudflare R2 plus Neon.
DOCS_FOLDER_PATH = "generated_docx"
DRY_RUN = False
LIMIT = None
PREFIX = None
MANIFEST = None
IGNORE_MANIFEST = False
ENV_FILE = None
INSTALL_DEPS = False


def _load_uploader():
    if not UPLOAD_SCRIPT.exists():
        raise SystemExit(f"ERROR: uploader script not found: {UPLOAD_SCRIPT}")

    spec = importlib.util.spec_from_file_location("data_upload_5", UPLOAD_SCRIPT)
    if spec is None or spec.loader is None:
        raise SystemExit(f"ERROR: could not load uploader script: {UPLOAD_SCRIPT}")

    module = importlib.util.module_from_spec(spec)
    sys.modules["data_upload_5"] = module
    spec.loader.exec_module(module)
    return module


def _safe_segment(value: str) -> str:
    segment = re.sub(r"[^A-Za-z0-9._-]+", "-", value.strip())
    segment = re.sub(r"-{2,}", "-", segment).strip("-._")
    return segment.lower() or "docs"


def _resolve_docs_folder(folder_name: str) -> Path:
    candidate = Path(folder_name).expanduser()
    if not candidate.is_absolute():
        candidate = SCRIPT_DIR / candidate

    if candidate.is_dir():
        return candidate.resolve()

    matches = [
        path for path in SCRIPT_DIR.rglob("*")
        if path.is_dir() and path.name.lower() == folder_name.lower()
    ]

    if len(matches) == 1:
        return matches[0].resolve()

    if len(matches) > 1:
        options = "\n".join(f"  {path}" for path in matches[:20])
        raise SystemExit(
            "ERROR: more than one folder matched that name. Use the full path:\n"
            f"{options}"
        )

    raise SystemExit(f"ERROR: folder not found: {folder_name}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Upload all PDF/DOCX files from a named docs folder to Cloudflare R2 and Neon"
    )
    parser.add_argument(
        "folder_name",
        nargs="?",
        default=DOCS_FOLDER_PATH,
        help="Optional override for DOCS_FOLDER_PATH",
    )
    parser.add_argument("--dry-run", action="store_true", default=DRY_RUN, help="Parse only; do not upload or write DB rows")
    parser.add_argument("--limit", type=int, default=LIMIT, help="Only process first N files")
    parser.add_argument("--prefix", default=PREFIX, help="Cloudflare R2 key prefix; default is docs/<folder-name>")
    parser.add_argument("--manifest", default=MANIFEST, help="JSONL progress manifest path")
    parser.add_argument("--ignore-manifest", action="store_true", default=IGNORE_MANIFEST, help="Do not skip hashes in manifest")
    parser.add_argument("--env-file", default=ENV_FILE, help="Env file path, relative to the uploader by default")
    parser.add_argument("--install-deps", action="store_true", default=INSTALL_DEPS, help="Install missing Python packages, then continue")
    args = parser.parse_args()

    uploader = _load_uploader()
    folder = _resolve_docs_folder(args.folder_name)
    segment = _safe_segment(folder.name)

    upload_args = argparse.Namespace(
        folder=str(folder),
        dry_run=args.dry_run,
        limit=args.limit,
        prefix=args.prefix or f"docs/{segment}",
        manifest=args.manifest or f"data_upload_{segment}_manifest.jsonl",
        ignore_manifest=args.ignore_manifest,
        env_file=args.env_file or uploader.DEFAULT_ENV_FILE,
        install_deps=args.install_deps,
        no_save_credentials=True,
    )

    if upload_args.install_deps:
        uploader._install_missing_dependencies()

    print(f"Docs folder: {folder}")
    print(f"Cloudflare prefix: {upload_args.prefix}")
    print(f"Manifest: {upload_args.manifest}")

    uploader._prepare_runtime_env(upload_args)
    uploader.import_folder(upload_args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
