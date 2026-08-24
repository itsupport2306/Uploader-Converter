# Complete Screenshot to Cloudflare Upload Process

This project uses `complete_process.py` to connect the screenshot-to-DOCX converter with the Neon profile insert/update flow and the Cloudflare R2 resume upload.

The converter logic is reused from `Converter/Converter/Converter/screenshot_to_word.py`. The upload and database helpers are reused from `data_upload 5.py`.

## What It Does

For each screenshot, the runner:

1. Finds source images from either a local folder/file or a OneDrive/SharePoint shared folder URL.
2. Converts the screenshot to DOCX.
3. Extracts text and profile fields from that DOCX.
4. Checks for an existing profile match by resume hash, NPI, email, or phone plus last name.
5. Uploads the DOCX to Cloudflare R2.
6. Inserts or updates the `profiles` row in Neon.

## Input Modes

### Local folder or file

```powershell
python complete_process.py "C:\path\to\screenshots"
```

This keeps the previous local staging behavior and writes generated `.docx` files to the configured output directory.

### OneDrive or SharePoint shared folder URL

```powershell
python complete_process.py "https://<your-share-link>" --onedrive-client-id "<app-client-id>" --onedrive-tenant common
```

For URL inputs:

- Screenshot images are downloaded directly from Microsoft Graph.
- DOCX conversion runs in memory.
- DOCX upload to Cloudflare R2 runs from memory.
- No local DOCX files are created.
- Local manifest and upload-log files are disabled by default.

If you still want a local manifest or CSV log for URL inputs, pass `--manifest <path>` and/or `--upload-log <path>`.

## Configuration

Create a `.env` file next to `complete_process.py`, or pass `--env-file PATH`.

Required settings:

```env
DATABASE_URL=postgresql://...
STORAGE_ENABLED=true
S3_ENDPOINT_URL=https://<account-id>.r2.cloudflarestorage.com
S3_REGION=auto
S3_BUCKET=<bucket-name>
S3_ACCESS_KEY=<r2-access-key>
S3_SECRET_KEY=<r2-secret-key>
S3_PUBLIC_BASE_URL=https://<public-domain-or-r2-dev-url>
S3_PUBLIC=true
S3_ACL=
```

If `S3_PUBLIC=true` and `S3_PUBLIC_BASE_URL` is set, the database stores a full public URL. Otherwise it stores an internal path like `/files/resumes/...`.

Tesseract OCR must also be installed. If it is not in a standard Windows location:

```powershell
python complete_process.py "C:\path\to\screenshots" --tesseract "C:\path\to\tesseract.exe"
```

## OneDrive / SharePoint Auth

Shared-folder URL inputs require a Microsoft Graph public client app registration.

Use either:

```powershell
python complete_process.py "https://<your-share-link>" --onedrive-client-id "<app-client-id>" --onedrive-tenant common
```

or environment variables:

```env
ONEDRIVE_CLIENT_ID=<app-client-id>
ONEDRIVE_TENANT=common
```

Notes:

- `common` is usually the best starting value.
- If your org requires a specific tenant, pass that tenant ID instead.
- `--no-browser` prints the device-login URL/code without trying to open a browser.

## Useful Commands

Local input:

```powershell
python complete_process.py "C:\path\to\screenshots" --limit 10
python complete_process.py "C:\path\to\screenshots" --force-convert
python complete_process.py "C:\path\to\screenshots" --output-dir "C:\path\to\docx_out"
python complete_process.py "C:\path\to\screenshots" --dry-run
```

OneDrive / SharePoint input:

```powershell
python complete_process.py "https://<your-share-link>" --onedrive-client-id "<app-client-id>" --onedrive-tenant common --dry-run
python complete_process.py "https://<your-share-link>" --onedrive-client-id "<app-client-id>" --onedrive-tenant common --limit 5
python complete_process.py "https://<your-share-link>" --onedrive-client-id "<app-client-id>" --onedrive-tenant common --manifest "complete_process_profiles_manifest.jsonl" --upload-log "complete_process_upload_log.csv"
```

Install dependencies:

```powershell
python complete_process.py "C:\path\to\screenshots" --install-deps --dry-run --limit 1
```

## Retry Behavior

For local inputs, the default manifest is `complete_process_profiles_manifest.jsonl`, so reruns skip records already marked `processed` or `skipped_duplicate`.

Examples:

```powershell
python complete_process.py "C:\path\to\screenshots" --retry-all
python complete_process.py "C:\path\to\screenshots" --ignore-manifest
python complete_process.py "C:\path\to\screenshots" --manifest none
```

For OneDrive / SharePoint URL inputs, manifest tracking is off by default unless you explicitly pass `--manifest`.

## Verification

Check the latest uploaded profile in Neon:

```sql
SELECT profile_id, first_name, last_name, resume_url, source, created_at, updated_at
FROM profiles
WHERE resume_url IS NOT NULL
ORDER BY created_at DESC
LIMIT 10;
```

Check a specific uploaded object:

```sql
SELECT profile_id, first_name, last_name, resume_url
FROM profiles
WHERE resume_url LIKE '%<cloudflare_key_or_hash>%';
```

The command prints a final summary like:

```text
Summary: {'processed': 10, 'skipped': 0, 'failed': 0, 'total': 10}
```

Cloudflare keys are hash-based, so if the same DOCX already exists in R2, the upload step reuses the existing object instead of creating a duplicate.
