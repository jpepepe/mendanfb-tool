# -*- coding: utf-8 -*-
"""
補正版（全文再分析）の酒井8件を共有Driveの json/ に上書きし、
ダッシュボードのサマリーキャッシュを無効化して再構築させる。
"""
import io, json, toml
from pathlib import Path
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2.service_account import Credentials

SHARED = "0ADyp-MOPyOQ5Uk9PVA"
SRC = Path("/tmp/saito_full")

sec = toml.load(".streamlit/secrets.toml")
cd = dict(sec["gcp_service_account"]); cd["private_key"] = cd["private_key"].replace("\\n", "\n")
creds = Credentials.from_service_account_info(cd, scopes=["https://www.googleapis.com/auth/drive"])
svc = build("drive", "v3", credentials=creds)


def subfolder(name):
    r = svc.files().list(
        q=f"'{SHARED}' in parents and name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True,
        corpora="drive", driveId=SHARED).execute()
    return r["files"][0]["id"]


def find_file(name, folder_id):
    r = svc.files().list(
        q=f"'{folder_id}' in parents and name='{name}' and trashed=false",
        fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    return r["files"][0]["id"] if r["files"] else None


def upload_json(name, data, folder_id):
    content = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype="application/json")
    existing = find_file(name, folder_id)
    if existing:
        svc.files().update(fileId=existing, media_body=media, supportsAllDrives=True).execute()
        return "更新"
    meta = {"name": name, "parents": [folder_id]}
    svc.files().create(body=meta, media_body=media, supportsAllDrives=True).execute()
    return "新規"


json_fid = subfolder("json")
files = sorted(SRC.glob("*.json"))
print(f"対象 {len(files)} 件を json/ に反映:")
for f in files:
    data = json.loads(f.read_text(encoding="utf-8"))
    act = upload_json(f.name, data, json_fid)
    print(f"  [{act}] {f.name}  (条件={data['grip_drivers'].get('条件',{}).get('score')}, grade={data['overall'].get('grade')})")

print("\n完了。※ダッシュボードに反映するには画面の「データを再読み込み」ボタンを押してください")
print("（json/ から全件再構築され、補正後スコアに更新されます）")
