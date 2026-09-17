# delete_physical_a_profiles

Deletes every profile whose `resume_sections` mentions `Physical_A`, together
with its resume object in Cloudflare R2. Standalone script, does not import
or modify anything else in the repo. Reads `../word_profile_pipeline/.env` by
default (same `DATABASE_URL` / `S3_*` settings as the rest of the pipeline).

## Match query

```sql
SELECT profile_id, resume_sections::text
FROM profiles
WHERE resume_sections::text ILIKE '%Physical_A%';
```

## Flow per profile

1. Read `resume_sections.cloudflare_bucket` and `resume_sections.cloudflare_key`
   (falls back to the `.env` `S3_BUCKET` if `cloudflare_bucket` is missing).
2. Delete that object from Cloudflare R2 (S3-compatible API). Idempotent —
   if the object is already gone, this is treated as success.
3. Only if step 2 did not raise: `DELETE FROM profiles WHERE profile_id = ...`.
4. Append one line to `logs/delete_manifest.jsonl` regardless of outcome.

If `cloudflare_key` (or the bucket) is missing, the profile is **skipped** and
left in the database — nothing is deleted for that row.

## Run

```powershell
cd D:\Radixsol\delete_physical_a_profiles
python delete_physical_a.py --dry-run --limit 5   # check a few matches first
python delete_physical_a.py --dry-run              # see every match with no changes
python delete_physical_a.py                        # actually delete everything matched
```

## Logs

`logs/delete_manifest.jsonl` — one line per profile (`deleted` / `would_delete` /
`skipped` / `failed`), with bucket, key, and any error. Re-running skips
`profile_id`s already marked `deleted`; `--retry-all` ignores the manifest and
redoes every match.
