# Word Profile Pipeline

Reads resume **DOCX files directly from SharePoint/OneDrive or a local folder**,
extracts profile details with **Qwen2.5**, uploads the **original, untouched
DOCX** to **Cloudflare R2**, and inserts/updates the profile in **Neon
Postgres**.

```
SharePoint DOCX
      -> read DOCX text (in memory)
      -> extract profile details (Qwen2.5)
      -> resolve years of experience
      -> generate Profile ID
      -> upload the original DOCX to Cloudflare R2
      -> insert/update the profile in Neon
      -> log every record
```

## Setup

```powershell
cd word_profile_pipeline
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env
```

Use the same `.env` values as `pdf_profile_pipeline`.

## Usage

```powershell
python run_pipeline.py "https://<your-share-link>"
python run_pipeline.py "C:\path\to\docx"
python run_pipeline.py "C:\path\to\docx" --limit 5 --dry-run
python run_pipeline.py "C:\path\to\docx" --no-llm
```

The pipeline accepts `.docx` files only.
