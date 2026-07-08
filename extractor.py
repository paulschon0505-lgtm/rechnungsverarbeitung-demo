"""
Text-Extraktion aus Dokumenten (PDF, PNG, JPG).

Strategie:
1. PDF mit eingebettetem Text (z.B. aus Buchhaltungssoftware/Word exportiert)
   -> Text wird direkt ausgelesen (pdfplumber). Kostenlos, schnell, sehr genau.
2. PDF ohne Text (eingescanntes Dokument) oder Bilddatei
   -> Seiten werden gerendert bzw. das Bild wird direkt an Tesseract OCR
      übergeben (komplett lokal, keine Cloud, keine Kosten).

Tesseract selbst ist eine separate, kostenlose Installation (siehe README).
Ist es nicht installiert, funktioniert die App weiterhin für alle
textbasierten PDFs - nur der OCR-Fallback für Scans/Fotos entfällt dann
mit einer verständlichen Fehlermeldung statt eines Absturzes.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO

import fitz  # PyMuPDF
import pdfplumber
from PIL import Image

try:
    import pytesseract

    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False


OCR_DPI = 300
MIN_CHARS_FOR_DIGITAL_PDF = 40  # darunter gilt eine PDF-Seite als "leer" -> OCR nötig


@dataclass
class ExtractionResult:
    filename: str
    text: str
    method: str  # "pdf-text" | "ocr" | "pdf-text+ocr" | "fehler"
    page_count: int = 0
    warnings: list[str] = field(default_factory=list)


def _ocr_image(image: Image.Image) -> str:
    if not TESSERACT_AVAILABLE:
        return ""
    try:
        return pytesseract.image_to_string(image, lang="deu+eng")
    except Exception:
        # z.B. wenn Tesseract-Sprachpakete fehlen oder die Binary nicht im PATH liegt
        try:
            return pytesseract.image_to_string(image)
        except Exception:
            return ""


def _render_pdf_page_to_image(page: fitz.Page) -> Image.Image:
    zoom = OCR_DPI / 72
    matrix = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=matrix)
    return Image.open(BytesIO(pix.tobytes("png")))


def extract_from_pdf(file_bytes: bytes, filename: str) -> ExtractionResult:
    warnings: list[str] = []
    texts: list[str] = []
    used_ocr = False

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        page_count = len(pdf.pages)
        digital_texts = [(p.extract_text() or "") for p in pdf.pages]

    needs_ocr_pages = [
        i for i, t in enumerate(digital_texts) if len(t.strip()) < MIN_CHARS_FOR_DIGITAL_PDF
    ]

    if needs_ocr_pages:
        if not TESSERACT_AVAILABLE:
            warnings.append(
                "Dokument enthält (teilweise) keinen eingebetteten Text - vermutlich "
                "ein Scan/Foto. Für die Texterkennung wird Tesseract OCR benötigt, "
                "das lokal nicht installiert ist. Siehe README für die (kostenlose) "
                "Installation."
            )
        else:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for i in needs_ocr_pages:
                image = _render_pdf_page_to_image(doc[i])
                ocr_text = _ocr_image(image)
                if ocr_text.strip():
                    digital_texts[i] = ocr_text
                    used_ocr = True
            doc.close()

    texts = digital_texts
    method = "pdf-text"
    if used_ocr and any(len(t.strip()) >= MIN_CHARS_FOR_DIGITAL_PDF for t in digital_texts):
        method = "pdf-text+ocr"
    elif used_ocr:
        method = "ocr"

    return ExtractionResult(
        filename=filename,
        text="\n".join(texts),
        method=method,
        page_count=page_count,
        warnings=warnings,
    )


def extract_from_image(file_bytes: bytes, filename: str) -> ExtractionResult:
    if not TESSERACT_AVAILABLE:
        return ExtractionResult(
            filename=filename,
            text="",
            method="fehler",
            page_count=1,
            warnings=[
                "Bilddatei kann nicht gelesen werden: Tesseract OCR ist lokal nicht "
                "installiert. Siehe README für die (kostenlose) Installation."
            ],
        )
    image = Image.open(BytesIO(file_bytes))
    text = _ocr_image(image)
    warnings = [] if text.strip() else ["OCR konnte keinen Text im Bild erkennen."]
    return ExtractionResult(
        filename=filename, text=text, method="ocr", page_count=1, warnings=warnings
    )


def extract(file_bytes: bytes, filename: str) -> ExtractionResult:
    """Erkennt den Dateityp anhand der Endung und wählt die passende Extraktion."""
    lower = filename.lower()
    try:
        if lower.endswith(".pdf"):
            return extract_from_pdf(file_bytes, filename)
        if lower.endswith((".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp")):
            return extract_from_image(file_bytes, filename)
        return ExtractionResult(
            filename=filename,
            text="",
            method="fehler",
            warnings=[f"Dateityp von '{filename}' wird nicht unterstützt."],
        )
    except Exception as exc:  # z.B. beschädigte Datei
        return ExtractionResult(
            filename=filename,
            text="",
            method="fehler",
            warnings=[f"Datei konnte nicht gelesen werden: {exc}"],
        )
