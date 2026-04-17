import io
from typing import Optional

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

from config import Config


class DriveHandler:
    def __init__(self, config: Config, credentials: Credentials):
        self.config = config
        self.service = build("drive", "v3", credentials=credentials)
        self._root_folder_id: Optional[str] = None

    def upload_cv(self, content: bytes, filename: str, job_title: str, mime_type: str) -> str:
        """Sube el CV a Drive y retorna la URL pública de visualización."""
        job_folder_id = self._get_or_create_job_folder(job_title)

        file_metadata = {
            "name": filename,
            "parents": [job_folder_id],
        }

        media = MediaIoBaseUpload(
            io.BytesIO(content),
            mimetype=mime_type,
            resumable=True
        )

        file = self.service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink"
        ).execute()

        return file.get("webViewLink", "")

    def _get_or_create_job_folder(self, job_title: str) -> str:
        root_id = self._get_or_create_root_folder()
        safe_title = job_title[:100]

        query = (
            f"name='{safe_title}' "
            f"and '{root_id}' in parents "
            f"and mimeType='application/vnd.google-apps.folder' "
            f"and trashed=false"
        )
        results = self.service.files().list(q=query, fields="files(id)").execute()
        files = results.get("files", [])

        if files:
            return files[0]["id"]

        folder = self.service.files().create(
            body={
                "name": safe_title,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [root_id],
            },
            fields="id"
        ).execute()
        return folder["id"]

    def _get_or_create_root_folder(self) -> str:
        if self._root_folder_id:
            return self._root_folder_id

        name = self.config.DRIVE_ROOT_FOLDER_NAME
        query = (
            f"name='{name}' "
            f"and mimeType='application/vnd.google-apps.folder' "
            f"and trashed=false"
        )
        results = self.service.files().list(q=query, fields="files(id)").execute()
        files = results.get("files", [])

        if files:
            self._root_folder_id = files[0]["id"]
        else:
            folder = self.service.files().create(
                body={
                    "name": name,
                    "mimeType": "application/vnd.google-apps.folder",
                },
                fields="id"
            ).execute()
            self._root_folder_id = folder["id"]

        return self._root_folder_id
