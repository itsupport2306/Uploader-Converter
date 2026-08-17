# Complete Screenshot to Cloudflare Upload Process

This project now has a single runner, `complete_process.py`, that connects the existing screenshot-to-DOCX converter with the existing Neon database and Cloudflare R2 uploader.

The converter logic is reused from `Converter/Converter/Converter/screenshot_to_word.py`. The upload and database helpers are reused from `data_upload 5.py`.

## How The Complete Process Works

Run the process with one command:

```powershell
python complete_process.py "C:\path\to\screenshots"
```

For each screenshot record, the runner works in this order:

1. Finds required screenshot records in the input file or folder.
2. Skips records already marked `processed` or `skipped_duplicate` in `complete_process_manifest.jsonl`.
3. Converts the screenshot into a local `.docx` file in `generated_docx` inside the input folder.
4. Reads the generated `.docx` and extracts profile fields.
5. Checks existing database identities once, then avoids duplicate records by NPI, email, or last name plus phone.
6. Inserts the profile data into Neon.
7. Uploads the generated `.docx` to Cloudflare R2.
8. Saves the Cloudflare path or public URL in `profiles.resume_url`.
9. Appends `processed`, `skipped_duplicate`, `failed`, or `would_process` status to the JSONL manifest.

Failed records are not terminal, so running the same command again retries them.

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

If `S3_PUBLIC=true` and `S3_PUBLIC_BASE_URL` is set, the database stores a full public URL. Otherwise it stores an internal path like `/files/resumes/import/...`.

Tesseract OCR must also be installed for screenshot conversion. If it is not in a standard Windows location, pass:

```powershell
python complete_process.py "C:\path\to\screenshots" --tesseract "C:\path\to\tesseract.exe"
```

Install Python dependencies:

```powershell
python complete_process.py "C:\path\to\screenshots" --install-deps --dry-run --limit 1
```

## Run The Full Process

Process all screenshots in a folder:

```powershell
python complete_process.py "C:\path\to\screenshots"
```

Useful options:

```powershell
python complete_process.py "C:\path\to\screenshots" --output-dir "C:\path\to\docx_out"
python complete_process.py "C:\path\to\screenshots" --limit 10
python complete_process.py "C:\path\to\screenshots" --prefix "resumes/import"
python complete_process.py "C:\path\to\screenshots" --force-convert
```

## Test With A Sample Record

Use `--dry-run` first. This creates/parses the DOCX but does not write to Neon and does not upload to Cloudflare:

```powershell
python complete_process.py "C:\path\to\one_screenshot.png" --dry-run --limit 1
```

Then run the same sample against the real services:

```powershell
python complete_process.py "C:\path\to\one_screenshot.png" --limit 1
```

## Verify The DOCX Was Created Correctly

Check the configured output folder. By default it is created inside the input folder:

```text
C:\path\to\screenshots\generated_docx
```

Open the generated `.docx` and confirm the physician name, specialty, headings, education, licensing, and publication sections look correct. The conversion logic is unchanged from the existing converter.

## Verify The Database Upload

In Neon or any Postgres client, check the inserted profile:

```sql
SELECT profile_id, first_name, last_name, resume_url, source, created_at, updated_at
FROM profiles
WHERE resume_url IS NOT NULL
ORDER BY created_at DESC
LIMIT 10;
```

To check a specific Cloudflare key from the manifest:

```sql
SELECT profile_id, first_name, last_name, resume_url
FROM profiles
WHERE resume_url LIKE '%<cloudflare_key_or_hash>%';
```

Related certification, license, and specialty records are inserted into `certifications`, `licenses`, and `profile_skills` when the parser finds those values.

## Verify The Cloudflare Upload

Open `complete_process_manifest.jsonl` and find the latest record with:

```json
"status": "processed"
```

That record contains:

```json
"cloudflare_key": "resumes/import/...",
"resume_url": "..."
```

Verify in Cloudflare R2 that the object exists at `cloudflare_key`. If `S3_PUBLIC=true`, open `resume_url` in a browser. If public access is disabled, use Cloudflare R2 or your application file-serving route to verify the object.

## Check Succeeded Or Failed

The command prints a final summary:

```text
Summary: {'processed': 1, 'skipped': 0, 'failed': 0, 'total': 1}
```

The manifest is append-only and keeps the latest status per source screenshot:

```text
complete_process_manifest.jsonl
```

Statuses:

```text
processed          DOCX created, database updated, Cloudflare upload complete
failed             One step failed; error text is recorded
skipped_duplicate  A matching person already exists in the database
would_process      Dry-run result only
```

## Retry Failed Records

Run the same command again:

```powershell
python complete_process.py "C:\path\to\screenshots"
```

Records already marked `processed` or `skipped_duplicate` are skipped. Records marked `failed` are retried automatically.

To retry everything, including records already marked processed:

```powershell
python complete_process.py "C:\path\to\screenshots" --retry-all
```

To ignore the manifest completely:

```powershell
python complete_process.py "C:\path\to\screenshots" --ignore-manifest
```

Cloudflare keys are based on the generated DOCX hash. If the same DOCX was already uploaded, the uploader reuses the existing object instead of uploading a duplicate.
