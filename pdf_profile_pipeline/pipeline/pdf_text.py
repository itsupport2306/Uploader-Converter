"""Read text out of a PDF, in memory. The PDF itself is never modified."""
from __future__ import annotations

import io
import re


class PdfTextError(RuntimeError):
    pass


def _import_pypdf():
    try:
        import pypdf  # noqa: PLC0415
    except ImportError as exc:  # pragma: no cover - dependency guard
        raise SystemExit("ERROR: pypdf is required. Run: pip install -r requirements.txt") from exc
    return pypdf


def extract_text(pdf_bytes: bytes, *, max_pages: int | None = None) -> str:
    """Return the concatenated text layer of the PDF.

    Returns an empty string for scanned/image-only PDFs; the caller decides
    whether that is a failure. No OCR and no format conversion happens here.
    """
    pypdf = _import_pypdf()
    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    except Exception as exc:
        raise PdfTextError(f"Could not open PDF: {exc}") from exc

    if getattr(reader, "is_encrypted", False):
        try:
            reader.decrypt("")
        except Exception as exc:
            raise PdfTextError(f"PDF is encrypted and could not be opened: {exc}") from exc

    pages = reader.pages if max_pages is None else reader.pages[:max_pages]
    chunks: list[str] = []
    for index, page in enumerate(pages, start=1):
        try:
            chunks.append(page.extract_text() or "")
        except Exception as exc:
            chunks.append("")
            print(f"  warning: page {index} text extraction failed: {exc}")
    return normalize_text("\n".join(chunks))


def page_count(pdf_bytes: bytes) -> int:
    pypdf = _import_pypdf()
    try:
        return len(pypdf.PdfReader(io.BytesIO(pdf_bytes)).pages)
    except Exception:
        return 0


def normalize_text(text: str) -> str:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace(" ", " ").replace("–", "-").replace("—", "-")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return "\n".join(line.strip() for line in text.split("\n")).strip()


def looks_empty(text: str, *, min_chars: int = 40) -> bool:
    return len(re.sub(r"\s", "", text or "")) < min_chars
