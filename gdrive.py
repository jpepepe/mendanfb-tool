import json
import io
import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.service_account import Credentials

FOLDER_ID = "1mzuALou_MSN6Fs5ac6fEsB1aBTUsArAg"
SCOPES = ["https://www.googleapis.com/auth/drive"]

def get_drive_service():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)

def upload_json(filename: str, data: dict, subfolder: str = "json"):
    service = get_drive_service()
    # サブフォルダ取得 or 作成
    folder_id = _get_or_create_subfolder(service, subfolder)
    # 既存ファイルがあれば上書き
    existing_id = _find_file(service, filename, folder_id)
    content = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype="application/json")
    if existing_id:
        service.files().update(fileId=existing_id, media_body=media).execute()
    else:
        meta = {"name": filename, "parents": [folder_id]}
        service.files().create(body=meta, media_body=media).execute()

def list_json_files(subfolder: str = "json") -> list[dict]:
    service = get_drive_service()
    folder_id = _get_or_create_subfolder(service, subfolder)
    results = service.files().list(
        q=f"'{folder_id}' in parents and name contains '.json' and trashed=false",
        fields="files(id, name)",
        pageSize=1000,
    ).execute()
    return results.get("files", [])

def download_json(file_id: str) -> dict:
    service = get_drive_service()
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, service.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return json.loads(buf.read().decode("utf-8"))

def _get_or_create_subfolder(service, name: str) -> str:
    res = service.files().list(
        q=f"'{FOLDER_ID}' in parents and name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id)",
    ).execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [FOLDER_ID]}
    f = service.files().create(body=meta, fields="id").execute()
    return f["id"]

def _find_file(service, filename: str, folder_id: str):
    res = service.files().list(
        q=f"'{folder_id}' in parents and name='{filename}' and trashed=false",
        fields="files(id)",
    ).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None
