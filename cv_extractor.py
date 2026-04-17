import io
from typing import Optional


class CVExtractor:
    """Extrae texto de CVs en formato PDF o DOCX."""

    MAX_CHARS = 15000  # Límite para no exceder el contexto del modelo

    def extract(self, content: bytes, mime_type: str) -> str:
        if mime_type == "application/pdf":
            text = self._extract_pdf(content)
        elif mime_type in (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/msword",
        ):
            text = self._extract_docx(content)
        else:
            text = content.decode("utf-8", errors="ignore")

        text = text.strip()
        if len(text) > self.MAX_CHARS:
            text = text[: self.MAX_CHARS] + "\n[... texto truncado ...]"

        return text

    def _extract_pdf(self, content: bytes) -> str:
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                pages = []
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        pages.append(page_text)
                return "\n\n".join(pages)
        except Exception as e:
            return f"[Error al leer PDF: {e}]"

    def _extract_docx(self, content: bytes) -> str:
        try:
            from docx import Document
            doc = Document(io.BytesIO(content))
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
            return "\n".join(paragraphs)
        except Exception as e:
            return f"[Error al leer DOCX: {e}]"
