from __future__ import annotations

import io

import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]


def _service():
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
    result = svc.files().list(q=q, fields="files(id,name)", pageSize=1000).execute()
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


def upload_or_update(name: str, parent_id: str, data: bytes, mime_type: str) -> str:
    svc = _service()
    q = f"'{parent_id}' in parents and name='{name}' and trashed=false"
    existing = svc.files().list(q=q, fields="files(id)").execute().get("files", [])
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime_type, resumable=False)
    if existing:
        svc.files().update(fileId=existing[0]["id"], media_body=media).execute()
        return existing[0]["id"]
    meta = {"name": name, "parents": [parent_id]}
    return svc.files().create(body=meta, media_body=media, fields="id").execute()["id"]


def ensure_subfolder(name: str, parent_id: str) -> str:
    svc = _service()
    q = (f"'{parent_id}' in parents and name='{name}' "
         f"and mimeType='application/vnd.google-apps.folder' and trashed=false")
    existing = svc.files().list(q=q, fields="files(id)").execute().get("files", [])
    if existing:
        return existing[0]["id"]
    meta = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    return svc.files().create(body=meta, fields="id").execute()["id"]
