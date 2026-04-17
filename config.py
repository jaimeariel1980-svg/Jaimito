import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()


class Config:
    ANTHROPIC_API_KEY: str = os.getenv("ANTHROPIC_API_KEY", "")
    GOOGLE_CREDENTIALS_FILE: str = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials/credentials.json")
    GOOGLE_TOKEN_FILE: str = os.getenv("GOOGLE_TOKEN_FILE", "credentials/token.json")

    # Gmail: busca emails con adjuntos que no estén etiquetados como procesados
    GMAIL_SEARCH_QUERY: str = os.getenv(
        "GMAIL_QUERY",
        "has:attachment -label:cv-procesado"
    )
    PROCESSED_LABEL_NAME: str = os.getenv("PROCESSED_LABEL", "cv-procesado")

    # Drive
    DRIVE_ROOT_FOLDER_NAME: str = os.getenv("DRIVE_FOLDER_NAME", "CVs - Proceso de Selección")

    # Sheets: si está vacío, se crea un spreadsheet nuevo automáticamente
    SPREADSHEET_ID: str = os.getenv("SPREADSHEET_ID", "")

    CLAUDE_MODEL: str = "claude-sonnet-4-6"

    GOOGLE_SCOPES = [
        "https://www.googleapis.com/auth/gmail.modify",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/spreadsheets",
    ]

    CV_MIME_TYPES = {
        "application/pdf",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/msword",
    }

    def validate(self):
        errors = []
        if not self.ANTHROPIC_API_KEY:
            errors.append("ANTHROPIC_API_KEY no configurada")
        if not Path(self.GOOGLE_CREDENTIALS_FILE).exists():
            errors.append(f"Credenciales de Google no encontradas en: {self.GOOGLE_CREDENTIALS_FILE}")
        if errors:
            raise ValueError("Errores de configuración:\n  - " + "\n  - ".join(errors))
