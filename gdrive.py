import json
import io
from typing import Optional
import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.service_account import Credentials

SHARED_DRIVE_ID = "0ADyp-MOPyOQ5Uk9PVA"
SCOPES = ["https://www.googleapis.com/auth/drive"]

def get_drive_service():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)

def upload_json(filename: str, data: dict, subfolder: str = "json"):
    service = get_drive_service()
    folder_id = _get_or_create_subfolder(service, subfolder)
    existing_id = _find_file(service, filename, folder_id)
    content = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype="application/json")
    if existing_id:
        service.files().update(fileId=existing_id, media_body=media, supportsAllDrives=True).execute()
    else:
        meta = {"name": filename, "parents": [folder_id]}
        service.files().create(body=meta, media_body=media, supportsAllDrives=True).execute()

def list_json_files(subfolder: str = "json") -> list[dict]:
    service = get_drive_service()
    folder_id = _get_or_create_subfolder(service, subfolder)
    results = service.files().list(
        q=f"'{folder_id}' in parents and name contains '.json' and trashed=false",
        fields="files(id, name)",
        pageSize=1000,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    return results.get("files", [])

def download_json(file_id: str) -> dict:
    service = get_drive_service()
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, service.files().get_media(fileId=file_id, supportsAllDrives=True))
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return json.loads(buf.read().decode("utf-8"))

def _get_or_create_subfolder(service, name: str) -> str:
    res = service.files().list(
        q=f"'{SHARED_DRIVE_ID}' in parents and name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora="drive",
        driveId=SHARED_DRIVE_ID,
    ).execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [SHARED_DRIVE_ID]}
    f = service.files().create(body=meta, fields="id", supportsAllDrives=True).execute()
    return f["id"]

def _find_file(service, filename: str, folder_id: str):
    res = service.files().list(
        q=f"'{folder_id}' in parents and name='{filename}' and trashed=false",
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None

# ── サマリーキャッシュ（ダッシュボード高速化用） ─────────────
SUMMARY_FILENAME = "_dashboard_summary.json"

def download_json_by_name(filename: str, subfolder: str = "json") -> Optional[dict]:
    """ファイル名でJSONをDriveから取得。なければNone"""
    service = get_drive_service()
    folder_id = _get_or_create_subfolder(service, subfolder)
    file_id = _find_file(service, filename, folder_id)
    if not file_id:
        return None
    return download_json(file_id)

def upload_summary(records: list) -> None:
    """ダッシュボード用サマリーJSONをDriveに保存（_rawは除外）"""
    slim = [{k: v for k, v in r.items() if k != "_raw"} for r in records]
    upload_json(SUMMARY_FILENAME, {"records": slim, "version": 1}, subfolder="json")

def download_summary() -> Optional[list]:
    """サマリーJSONをDriveから取得。なければNone"""
    try:
        data = download_json_by_name(SUMMARY_FILENAME, subfolder="json")
        if data and "records" in data:
            return data["records"]
    except Exception:
        pass
    return None

PROPOSAL_SUMMARY_FILENAME = "_proposal_summary.json"

def upload_proposal_summary(records: list) -> None:
    """求人提案ダッシュボード用サマリーJSONをDriveに保存（_rawは除外）"""
    slim = [{k: v for k, v in r.items() if k != "_raw"} for r in records]
    upload_json(PROPOSAL_SUMMARY_FILENAME, {"records": slim, "version": 1}, subfolder="json_proposal")

def download_proposal_summary() -> Optional[list]:
    """求人提案サマリーJSONをDriveから取得。なければNone"""
    try:
        data = download_json_by_name(PROPOSAL_SUMMARY_FILENAME, subfolder="json_proposal")
        if data and "records" in data:
            return data["records"]
    except Exception:
        pass
    return None
