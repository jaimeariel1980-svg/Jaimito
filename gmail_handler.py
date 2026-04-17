import base64
import os
from pathlib import Path
from typing import Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from config import Config


class GmailHandler:
    def __init__(self, config: Config):
        self.config = config
        self.service = self._build_service()
        self._processed_label_id: Optional[str] = None

    def _build_service(self):
        creds = self._get_credentials()
        return build("gmail", "v1", credentials=creds)

    def _get_credentials(self):
        token_path = Path(self.config.GOOGLE_TOKEN_FILE)
        creds_path = Path(self.config.GOOGLE_CREDENTIALS_FILE)
        creds = None

        if token_path.exists():
            creds = Credentials.from_authorized_user_file(str(token_path), self.config.GOOGLE_SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), self.config.GOOGLE_SCOPES)
                creds = flow.run_local_server(port=0)
            token_path.parent.mkdir(parents=True, exist_ok=True)
            token_path.write_text(creds.to_json())

        return creds

    def get_unprocessed_applications(self) -> list[dict]:
        results = self.service.users().messages().list(
            userId="me",
            q=self.config.GMAIL_SEARCH_QUERY,
            maxResults=50
        ).execute()

        messages = results.get("messages", [])
        applications = []

        for msg_ref in messages:
            msg = self.service.users().messages().get(
                userId="me",
                id=msg_ref["id"],
                format="full"
            ).execute()

            email_data = self._parse_message(msg)
            if email_data:
                applications.append(email_data)

        return applications

    def _parse_message(self, msg: dict) -> Optional[dict]:
        headers = {h["name"]: h["value"] for h in msg["payload"].get("headers", [])}

        email_data = {
            "id": msg["id"],
            "thread_id": msg["threadId"],
            "subject": headers.get("Subject", "(sin asunto)"),
            "from": headers.get("From", ""),
            "date": headers.get("Date", ""),
            "attachments": [],
        }

        self._extract_attachments(msg["payload"], msg["id"], email_data["attachments"])

        cv_attachments = [
            a for a in email_data["attachments"]
            if a["mime_type"] in self.config.CV_MIME_TYPES
        ]

        if not cv_attachments:
            return None

        email_data["cv_attachments"] = cv_attachments
        return email_data

    def _extract_attachments(self, payload: dict, msg_id: str, attachments: list):
        if "parts" in payload:
            for part in payload["parts"]:
                self._extract_attachments(part, msg_id, attachments)
        elif payload.get("filename") and payload.get("body", {}).get("attachmentId"):
            attachments.append({
                "filename": payload["filename"],
                "mime_type": payload.get("mimeType", "application/octet-stream"),
                "attachment_id": payload["body"]["attachmentId"],
                "message_id": msg_id,
            })

    def download_attachment(self, attachment_meta: dict) -> bytes:
        attachment = self.service.users().messages().attachments().get(
            userId="me",
            messageId=attachment_meta["message_id"],
            id=attachment_meta["attachment_id"]
        ).execute()

        return base64.urlsafe_b64decode(attachment["data"])

    def mark_as_processed(self, message_id: str):
        label_id = self._get_or_create_label()
        self.service.users().messages().modify(
            userId="me",
            id=message_id,
            body={"addLabelIds": [label_id]}
        ).execute()

    def _get_or_create_label(self) -> str:
        if self._processed_label_id:
            return self._processed_label_id

        labels = self.service.users().labels().list(userId="me").execute()
        for label in labels.get("labels", []):
            if label["name"].lower() == self.config.PROCESSED_LABEL_NAME.lower():
                self._processed_label_id = label["id"]
                return self._processed_label_id

        new_label = self.service.users().labels().create(
            userId="me",
            body={
                "name": self.config.PROCESSED_LABEL_NAME,
                "labelListVisibility": "labelShow",
                "messageListVisibility": "show",
            }
        ).execute()
        self._processed_label_id = new_label["id"]
        return self._processed_label_id
