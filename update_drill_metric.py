# -*- coding: utf-8 -*-
"""
縦深掘り指標バグ修正後のロジックで、酒井8件の『縦深掘り最大』を再計算し
Drive json/ のJSONを上書き更新する。他の指標・分析結果は変更しない。
"""
import io, json, re, toml
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.service_account import Credentials

SHARED = "0ADyp-MOPyOQ5Uk9PVA"
DRILL_PAT = re.compile(r'なぜ|なんで|どうして|どんな.*とき|一番.*何|きっかけ|どうやって|どのよう|どういう.*気持|何が.*よかっ|何が.*楽し|どんな.*感じ')
TARGETS = [
    "酒井_C_岡田さん_初回面談.json", "酒井_C_佐藤さん_その他.json", "酒井_C_加藤_初回面談.json",
    "酒井_C_菅さん_その他.json", "酒井_C_櫻井 美彩稀様_その他.json", "酒井_C_辻 結愛_初回面談.json",
    "酒井_B_神谷_初回面談.json", "酒井_C_児子_その他.json",
]

sec = toml.load(".streamlit/secrets.toml")
cd = dict(sec["gcp_service_account"]); cd["private_key"] = cd["private_key"].replace("\\n", "\n")
creds = Credentials.from_service_account_info(cd, scopes=["https://www.googleapis.com/auth/drive"])
svc = build("drive", "v3", credentials=creds)


def subfolder(name):
    return svc.files().list(q=f"'{SHARED}' in parents and name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True, corpora="drive", driveId=SHARED).execute()["files"][0]["id"]

def list_files(fid):
    files, tok = [], None
    while True:
        r = svc.files().list(q=f"'{fid}' in parents and name contains '.json' and trashed=false",
            fields="nextPageToken,files(id,name)", pageSize=1000, supportsAllDrives=True, includeItemsFromAllDrives=True, pageToken=tok).execute()
        files += r.get("files", []); tok = r.get("nextPageToken")
        if not tok: break
    return {f["name"]: f["id"] for f in files}

def dl(fid):
    buf = io.BytesIO(); d = MediaIoBaseDownload(buf, svc.files().get_media(fileId=fid, supportsAllDrives=True))
    done = False
    while not done: _, done = d.next_chunk()
    buf.seek(0); return json.loads(buf.read().decode("utf-8"))

def upload(fid, data):
    media = MediaIoBaseUpload(io.BytesIO(json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")), mimetype="application/json")
    svc.files().update(fileId=fid, media_body=media, supportsAllDrives=True).execute()


def vertical_drill_max(utt, is_ca):
    """修正後ロジック：求職者の回答ではstreakを切らず、CAの非深掘り発話で終了"""
    ms = cs = 0
    for i in range(1, len(utt)):
        u = utt[i]; pv = utt[i-1]
        if not is_ca(u["speaker"]):
            continue
        if DRILL_PAT.search(u["text"]) and not is_ca(pv["speaker"]) and len(u["text"]) > 8:
            cs += 1; ms = max(ms, cs)
        else:
            cs = 0
    return ms


json_fid = subfolder("json"); utt_fid = subfolder("utterances")
json_map = list_files(json_fid); utt_map = list_files(utt_fid)

for name in TARGETS:
    if name not in json_map:
        print("❌ json未検出:", name); continue
    rec = dl(json_map[name])
    cand = rec.get("candidate", "")
    uid = utt_map.get(name) or next((fid for n, fid in utt_map.items() if "酒井" in n and cand and cand in n and "初回" in n), None) \
                            or next((fid for n, fid in utt_map.items() if "酒井" in n and cand and cand in n), None)
    if not uid:
        print("⚠️ utterances未検出（スキップ）:", name); continue
    utt = dl(uid)["utterances"]
    spk = set(u["speaker"] for u in utt); ca_spk = {s for s in spk if "酒井" in s or s == "CA"}
    is_ca = lambda s: s in ca_spk
    old = rec.get("behaviors", {}).get("縦深掘り最大")
    new = vertical_drill_max(utt, is_ca)
    rec.setdefault("behaviors", {})["縦深掘り最大"] = new
    upload(json_map[name], rec)
    print(f"  {cand}: 縦深掘り最大 {old} → {new}  ✅上書き")

print("完了。ダッシュボードは「データを再読み込み」で反映されます。")
