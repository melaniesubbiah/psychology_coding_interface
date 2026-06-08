from __future__ import annotations

import io

import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]

# Required for all calls so the service account can operate on Shared Drives.
_SD = {"supportsAllDrives": True}
_SD_LIST = {"supportsAllDrives": True, "includeItemsFromAllDrives": True}


@st.cache_resource(ttl=3600)
def _service():
    """Build and cache the Drive API client (shared across reruns for 1 hour)."""
    creds = Credentials.from_service_account_info(
        dict(st.secrets["google_service_account"]),
        scopes=SCOPES,
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def list_files(folder_id: str, mime_filter: str | None = None) -> list[dict]:
    svc = _service()
    q = f"'{folder_id}' in parents and trashed=false"
    if mime_filter:
        q += f" and mimeType='{mime_filter}'"
    result = (
        svc.files()
        .list(q=q, fields="files(id,name)", pageSize=1000, **_SD_LIST)
        .execute()
    )
    return result.get("files", [])


def download_bytes(file_id: str) -> bytes:
    svc = _service()
    req = svc.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue()


def update_by_id(file_id: str, data: bytes, mime_type: str) -> None:
    print("DEBUG!!!", file_id)
    """Update an existing file by known ID — no list query needed."""
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime_type, resumable=False)
    _service().files().update(fileId=file_id, media_body=media, **_SD).execute()


def upload_or_update(name: str, parent_id: str, data: bytes, mime_type: str) -> str:
    svc = _service()
    q = f"'{parent_id}' in parents and name='{name}' and trashed=false"
    existing = (
        svc.files().list(q=q, fields="files(id)", **_SD_LIST).execute().get("files", [])
    )
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime_type, resumable=False)
    if existing:
        svc.files().update(
            fileId=existing[0]["id"], media_body=media, **_SD
        ).execute()
        return existing[0]["id"]
    meta = {"name": name, "parents": [parent_id]}
    return (
        svc.files().create(body=meta, media_body=media, fields="id", **_SD).execute()["id"]
    )


def ensure_subfolder(name: str, parent_id: str) -> str:
    svc = _service()
    q = (
        f"'{parent_id}' in parents and name='{name}' "
        f"and mimeType='application/vnd.google-apps.folder' and trashed=false"
    )
    existing = (
        svc.files().list(q=q, fields="files(id)", **_SD_LIST).execute().get("files", [])
    )
    if existing:
        return existing[0]["id"]
    meta = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    return svc.files().create(body=meta, fields="id", **_SD).execute()["id"]
