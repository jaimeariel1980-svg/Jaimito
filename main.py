#!/usr/bin/env python3
"""
Automatización de proceso de selección: Gmail → Drive → Sheets + análisis IA.

Uso:
    python main.py job_profiles/dev_senior.json
    python main.py job_profiles/dev_senior.json --dry-run
"""

import argparse
import json
import sys
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from config import Config
from gmail_handler import GmailHandler
from drive_handler import DriveHandler
from sheets_handler import SheetsHandler
from cv_extractor import CVExtractor
from ai_analyzer import AIAnalyzer


def get_google_credentials(config: Config) -> Credentials:
    token_path = Path(config.GOOGLE_TOKEN_FILE)
    creds_path = Path(config.GOOGLE_CREDENTIALS_FILE)
    creds = None

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), config.GOOGLE_SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), config.GOOGLE_SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json())

    return creds


def process_applications(job_profile_path: str, dry_run: bool = False):
    with open(job_profile_path, encoding="utf-8") as f:
        job_profile = json.load(f)

    config = Config()
    config.validate()

    print(f"\n{'='*60}")
    print(f"Proceso de selección: {job_profile['title']}")
    print(f"Modo: {'DRY RUN (sin guardar)' if dry_run else 'PRODUCCIÓN'}")
    print(f"{'='*60}\n")

    creds = get_google_credentials(config)

    gmail = GmailHandler(config)
    drive = DriveHandler(config, creds)
    sheets = SheetsHandler(config, creds)
    extractor = CVExtractor()
    analyzer = AIAnalyzer(config)

    spreadsheet_id = None if dry_run else sheets.get_or_create_spreadsheet(job_profile["title"])

    emails = gmail.get_unprocessed_applications()
    print(f"Postulaciones encontradas: {len(emails)}\n")

    if not emails:
        print("No hay nuevas postulaciones para procesar.")
        return

    processed = 0
    errors = 0

    for email_data in emails:
        sender = email_data.get("from", "desconocido")
        subject = email_data.get("subject", "")
        print(f"Procesando: {sender}")
        print(f"  Asunto: {subject}")

        cv_attachments = email_data.get("cv_attachments", [])
        if not cv_attachments:
            print("  Sin adjuntos de CV válidos, omitiendo.\n")
            continue

        attachment_meta = cv_attachments[0]

        try:
            # 1. Descargar adjunto
            cv_content = gmail.download_attachment(attachment_meta)
            print(f"  CV descargado: {attachment_meta['filename']}")

            # 2. Extraer texto
            cv_text = extractor.extract(cv_content, attachment_meta["mime_type"])
            if cv_text.startswith("[Error"):
                print(f"  Advertencia: {cv_text}")

            # 3. Subir a Drive
            if not dry_run:
                drive_url = drive.upload_cv(
                    cv_content,
                    attachment_meta["filename"],
                    job_profile["title"],
                    attachment_meta["mime_type"],
                )
                print(f"  Subido a Drive: {drive_url}")
            else:
                drive_url = "DRY_RUN_URL"

            # 4. Análisis con IA
            print("  Analizando con IA...")
            analysis = analyzer.analyze(cv_text, job_profile)

            recommendation = "APLICA ✓" if analysis["applies"] else "NO APLICA ✗"
            print(f"  → {recommendation} | Puntaje: {analysis['score']}/10")
            print(f"     {analysis.get('summary', '')[:100]}")

            # 5. Guardar en Sheets
            if not dry_run:
                sheets.add_candidate(spreadsheet_id, {
                    "email_data": email_data,
                    "drive_url": drive_url,
                    "analysis": analysis,
                    "job_title": job_profile["title"],
                })

                # 6. Marcar email como procesado
                gmail.mark_as_processed(email_data["id"])

            processed += 1

        except Exception as e:
            print(f"  ERROR: {e}")
            errors += 1

        print()

    print(f"{'='*60}")
    print(f"Resumen: {processed} procesados, {errors} errores")
    if not dry_run and spreadsheet_id:
        print(f"Resultados: {sheets.get_spreadsheet_url(spreadsheet_id)}")
    print(f"{'='*60}\n")


def main():
    parser = argparse.ArgumentParser(
        description="Automatización de proceso de selección con IA"
    )
    parser.add_argument(
        "job_profile",
        help="Ruta al perfil del puesto en formato JSON (ej: job_profiles/dev_senior.json)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Procesa pero no guarda en Drive/Sheets ni etiqueta los emails"
    )
    args = parser.parse_args()

    if not Path(args.job_profile).exists():
        print(f"Error: No se encontró el archivo: {args.job_profile}")
        sys.exit(1)

    process_applications(args.job_profile, args.dry_run)


if __name__ == "__main__":
    main()
