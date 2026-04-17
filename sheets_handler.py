from datetime import datetime
from typing import Optional

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

from config import Config

HEADERS = [
    "Fecha Postulación",
    "Nombre",
    "Email Candidato",
    "Teléfono",
    "Posición",
    "Link CV (Drive)",
    "Años Experiencia",
    "Habilidades Clave",
    "Educación",
    "Puntaje IA (1-10)",
    "Recomendación",
    "Fortalezas",
    "Debilidades",
    "Resumen del Análisis",
    "Email Postulación",
    "Asunto Email",
    "Estado",
]


class SheetsHandler:
    def __init__(self, config: Config, credentials: Credentials):
        self.config = config
        self.service = build("sheets", "v4", credentials=credentials)

    def get_or_create_spreadsheet(self, job_title: str) -> str:
        if self.config.SPREADSHEET_ID:
            return self.config.SPREADSHEET_ID
        return self._create_spreadsheet(job_title)

    def _create_spreadsheet(self, job_title: str) -> str:
        spreadsheet = self.service.spreadsheets().create(
            body={
                "properties": {"title": f"Selección: {job_title}"},
                "sheets": [{"properties": {"title": "Candidatos"}}],
            }
        ).execute()

        spreadsheet_id = spreadsheet["spreadsheetId"]
        self._write_headers(spreadsheet_id)
        self._apply_formatting(spreadsheet_id)
        return spreadsheet_id

    def _write_headers(self, spreadsheet_id: str):
        self.service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range="Candidatos!A1",
            valueInputOption="RAW",
            body={"values": [HEADERS]},
        ).execute()

    def _apply_formatting(self, spreadsheet_id: str):
        sheet_id = self._get_sheet_id(spreadsheet_id, "Candidatos")
        requests = [
            # Encabezado en negrita con fondo azul oscuro
            {
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.2, "green": 0.4, "blue": 0.7},
                            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                            "horizontalAlignment": "CENTER",
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                }
            },
            # Fila fija (freeze)
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                    "fields": "gridProperties.frozenRowCount",
                }
            },
            # Formato condicional: APLICA → verde
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": sheet_id, "startColumnIndex": 10, "endColumnIndex": 11}],
                        "booleanRule": {
                            "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "APLICA"}]},
                            "format": {"backgroundColor": {"red": 0.7, "green": 0.9, "blue": 0.7}},
                        },
                    },
                    "index": 0,
                }
            },
            # Formato condicional: NO APLICA → rojo
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": sheet_id, "startColumnIndex": 10, "endColumnIndex": 11}],
                        "booleanRule": {
                            "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "NO APLICA"}]},
                            "format": {"backgroundColor": {"red": 0.95, "green": 0.7, "blue": 0.7}},
                        },
                    },
                    "index": 1,
                }
            },
        ]

        self.service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()

    def _get_sheet_id(self, spreadsheet_id: str, sheet_name: str) -> int:
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sheet in spreadsheet["sheets"]:
            if sheet["properties"]["title"] == sheet_name:
                return sheet["properties"]["sheetId"]
        return 0

    def add_candidate(self, spreadsheet_id: str, data: dict):
        analysis = data["analysis"]
        email_data = data["email_data"]

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            analysis.get("candidate_name", ""),
            analysis.get("candidate_email", ""),
            analysis.get("candidate_phone", ""),
            data.get("job_title", ""),
            data.get("drive_url", ""),
            str(analysis.get("years_experience", "")),
            ", ".join(analysis.get("key_skills", [])),
            analysis.get("education", ""),
            str(analysis.get("score", "")),
            "APLICA" if analysis.get("applies") else "NO APLICA",
            " | ".join(analysis.get("strengths", [])),
            " | ".join(analysis.get("weaknesses", [])),
            analysis.get("summary", ""),
            email_data.get("from", ""),
            email_data.get("subject", ""),
            "Pendiente revisión",
        ]

        self.service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range="Candidatos!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [row]},
        ).execute()

    def get_spreadsheet_url(self, spreadsheet_id: str) -> str:
        return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
