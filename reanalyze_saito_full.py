# -*- coding: utf-8 -*-
"""
酒井CA 直近10件を「全文トランスクリプト」で再分析するワンショットスクリプト。

背景: 旧来ツールは AI に渡すトランスクリプトを 15,000 字で打ち切っていたため、
面談後半（条件確認・クロージング）が分析されず誤判定していた。
analysis_core.py の打ち切りを 60,000 字へ拡張済みなので、その関数をそのまま使って
共有Drive上の全文 utterances を再採点し、補正版JSONを /tmp/saito_full/ に保存する。

行動指標（compute_metrics由来）はもともと全発話から計算され打ち切りの影響を受けないため、
既存JSONの値を温存し、LLM判定（grip_drivers / overall / notes / behaviors のフェーズ系 /
深掘り詳細）のみ全文で上書きする。
"""
import os, io, json, toml
from pathlib import Path
import anthropic
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials

import analysis_core as core  # 打ち切り拡張済み（MAX_TRANSCRIPT_CHARS=60000）

SHARED = "0ADyp-MOPyOQ5Uk9PVA"
OUT = Path("/tmp/saito_full"); OUT.mkdir(parents=True, exist_ok=True)

# 対象10件（Drive上の正確なファイル名）
TARGETS = [
    "酒井_C_岡田さん_初回面談.json",
    "酒井_C_佐藤さん_その他.json",
    "酒井_C_加藤_初回面談.json",
    "酒井_C_菅さん_その他.json",
    "酒井_C_櫻井 美彩稀様_その他.json",
    "酒井_C_辻 結愛_初回面談.json",
    "酒井_B_神谷_初回面談.json",
    "酒井_C_児子_その他.json",
    "酒井_D_梅本_初回面談.json",
    "酒井_D_尾高_初回面談.json",
]

# ── 認証 ─────────────────────────────────────────────
sec = toml.load(".streamlit/secrets.toml")
cd = dict(sec["gcp_service_account"]); cd["private_key"] = cd["private_key"].replace("\\n", "\n")
creds = Credentials.from_service_account_info(cd, scopes=["https://www.googleapis.com/auth/drive"])
svc = build("drive", "v3", credentials=creds)
client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])


def subfolder(name):
    r = svc.files().list(
        q=f"'{SHARED}' in parents and name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True,
        corpora="drive", driveId=SHARED).execute()
    return r["files"][0]["id"]


def list_files(folder_id):
    files, tok = [], None
    while True:
        r = svc.files().list(
            q=f"'{folder_id}' in parents and name contains '.json' and trashed=false",
            fields="nextPageToken, files(id,name)", pageSize=1000,
            supportsAllDrives=True, includeItemsFromAllDrives=True, pageToken=tok).execute()
        files += r.get("files", []); tok = r.get("nextPageToken")
        if not tok:
            break
    return files


def dl(file_id):
    buf = io.BytesIO()
    d = MediaIoBaseDownload(buf, svc.files().get_media(fileId=file_id, supportsAllDrives=True))
    done = False
    while not done:
        _, done = d.next_chunk()
    buf.seek(0)
    return json.loads(buf.read().decode("utf-8"))


json_fid = subfolder("json")
utt_fid = subfolder("utterances")
json_files = {f["name"]: f["id"] for f in list_files(json_fid)}
utt_files = {f["name"]: f["id"] for f in list_files(utt_fid)}


def find_utterances(json_name, old):
    """utterances を厳密名→候補者名一致の順で探す"""
    if json_name in utt_files:
        return utt_files[json_name]
    cand = old.get("candidate", "")
    for n, fid in utt_files.items():
        if cand and cand in n and "酒井" in n:
            return fid
    return None


rows = []
for name in TARGETS:
    if name not in json_files:
        print(f"❌ json未検出: {name}"); continue
    old = dl(json_files[name])
    uid = find_utterances(name, old)
    if not uid:
        rows.append((old.get("candidate", name), "—", "utterances欠落で再分析不可", old, None))
        print(f"⚠️ utterances未検出（スキップ）: {name}")
        continue

    utt = dl(uid)["utterances"]
    formatted_len = len("\n".join(f"[{u['speaker']}] {u['text']}" for u in utt))

    new = core.score_with_claude(utt, old.get("ca", "酒井"), old.get("candidate", ""), old.get("format", "docx"), client)
    deep = core.deep_analysis_with_claude(utt, old.get("ca", "酒井"), old.get("candidate", ""), client)

    if not new.get("overall"):
        print(f"⚠️ 再採点失敗（overall空）: {name}")

    corrected = dict(old)  # 既存スキーマを温存
    corrected["grip_drivers"] = new.get("grip_drivers", old.get("grip_drivers", {}))
    corrected["overall"] = new.get("overall", {})
    corrected["notes"] = new.get("notes", "")
    # 行動指標: 既存（rule-based含む全発話計算）に、新LLMのフェーズ系を上書きマージ
    corrected["behaviors"] = {**old.get("behaviors", {}), **new.get("behaviors", {})}
    corrected["emotion_drill_analysis"] = deep.get("emotion_drill_analysis", {})
    corrected["self_disclosure_analysis"] = deep.get("self_disclosure_analysis", {})
    corrected["backtrack_analysis"] = deep.get("backtrack_analysis", {})
    corrected["next_phrases"] = deep.get("next_phrases", [])
    corrected["_reanalyzed_full_transcript"] = True
    corrected["_formatted_transcript_len"] = formatted_len

    (OUT / name).write_text(json.dumps(corrected, ensure_ascii=False, indent=2), encoding="utf-8")
    rows.append((old.get("candidate", name), formatted_len, "OK", old, corrected))
    print(f"✅ {name}  (整形{formatted_len}字)")


def g(d, axis):
    return (d.get("grip_drivers", {}).get(axis, {}) or {}).get("score")


def ov(d, k):
    return (d.get("overall", {}) or {}).get(k)


print("\n================ before → after 比較 ================")
hdr = f"{'候補者':<8}{'grade':>7}{'条件':>6}{'認識統一':>8}{'気づき':>7}{'次回アポ':>9}  フェーズ網羅(after)"
print(hdr)
for cand, flen, status, old, new in rows:
    if new is None:
        print(f"{cand[:7]:<8}  {status}")
        continue
    grade = f"{ov(old,'grade')}→{ov(new,'grade')}"
    cond = f"{g(old,'条件')}→{g(new,'条件')}"
    nin = f"{g(old,'認識統一')}→{g(new,'認識統一')}"
    kiz = f"{g(old,'気づき')}→{g(new,'気づき')}"
    apo_o = old.get("behaviors", {}).get("次回アポ確定")
    apo_n = new.get("behaviors", {}).get("次回アポ確定")
    apo = f"{apo_o}→{apo_n}"
    ph = "/".join(new.get("behaviors", {}).get("フェーズ網羅", []))
    print(f"{cand[:7]:<8}{grade:>7}{cond:>6}{nin:>8}{kiz:>7}{apo:>9}  {ph[:60]}")

print(f"\n保存先: {OUT}  / 件数: {sum(1 for r in rows if r[4] is not None)}")
