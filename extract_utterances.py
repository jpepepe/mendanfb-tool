# -*- coding: utf-8 -*-
"""
話者分離テキスト抽出スクリプト
- docx: APIなしで即時抽出・保存
- txt: Claude Haiku（低コスト）で話者推定・保存
- 処理済みファイルはスキップ（冪等）
"""

import os
import re
import json
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
import anthropic

TRANSCRIPT_BASE = Path("/Users/takahashijunpei/文字起こし")
UTTERANCES_DIR = Path("/Users/takahashijunpei/Downloads/人材紹介分析フォルダ/output/utterances")

client = anthropic.Anthropic()


def find_files(meeting_type: str = None) -> list[dict]:
    files = []
    for root, _, fnames in os.walk(TRANSCRIPT_BASE):
        for fn in fnames:
            if fn.startswith("."):
                continue
            fn_nfc = unicodedata.normalize("NFC", fn)
            if not (fn_nfc.endswith(".txt") or fn_nfc.endswith(".docx")):
                continue
            stem = fn_nfc.rsplit(".", 1)[0]
            ext = fn_nfc.rsplit(".", 1)[1]
            stem_norm = re.sub(r"-グリップ", "_グリップ", stem)
            parts = re.split(r"[_＿]", stem_norm)
            grip = ca = candidate = mt = None
            for i, p in enumerate(parts):
                m = re.search(r"グリップ([ABCD])", p)
                if m:
                    grip = m.group(1)
                    ca = parts[0]
                    candidate = parts[i + 1] if i + 1 < len(parts) else None
                    mt = parts[i + 2] if i + 2 < len(parts) else None
                    break
            if grip and ca and mt:
                if meeting_type is None or mt == meeting_type:
                    files.append({
                        "path": os.path.join(root, fn),
                        "ca": ca, "grip": grip, "candidate": candidate,
                        "meeting_type": mt, "format": ext,
                    })
    return files


def _detect_speakers(full_text: str) -> list[str]:
    from collections import Counter
    candidates = re.findall(r"(?<![^\s。！？\n])([぀-鿿a-zA-Zー]{2,10}):\s", full_text)
    freq = Counter(candidates)
    return [name for name, cnt in freq.most_common(10) if cnt >= 3][:2]


def parse_docx(path: str) -> list[dict]:
    with zipfile.ZipFile(path) as z:
        with z.open("word/document.xml") as f:
            tree = ET.parse(f)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs = tree.getroot().findall(".//w:p", ns)
    full_text = ""
    for p in paragraphs:
        texts = p.findall(".//w:t", ns)
        line = "".join(t.text or "" for t in texts).strip()
        if line:
            full_text += line + "\n"
    full_text = re.sub(r"^\d{2}:\d{2}:\d{2}\n", "", full_text, flags=re.MULTILINE)
    full_text = re.sub(r"^\d{4}年.+\n", "", full_text, flags=re.MULTILINE)
    full_text = re.sub(r"^会議\s.+\n", "", full_text, flags=re.MULTILINE)

    speakers = _detect_speakers(full_text)
    if not speakers:
        return [{"speaker": "不明", "text": full_text}]

    speaker_pattern = re.compile(
        r"(?:" + "|".join(re.escape(s) for s in speakers) + r"):\s*"
    )
    utterances = []
    matches = list(speaker_pattern.finditer(full_text))
    for i, m in enumerate(matches):
        speaker = m.group().rstrip(": \t")
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(full_text)
        text = full_text[start:end].strip().replace("\n", " ")
        text = re.sub(r"\s+", " ", text).strip()
        if text:
            utterances.append({"speaker": speaker, "text": text})
    return utterances


def parse_txt_with_haiku(path: str, ca: str, candidate: str) -> list[dict]:
    """Haikuモデルで話者推定（低コスト）"""
    with open(path, encoding="utf-8") as f:
        raw = f.read()
    raw = re.sub(r"^\d+\t", "", raw, flags=re.MULTILINE)

    prompt = f"""以下は人材紹介会社の面談音声文字起こしです。
話者は2名です：
- CA（キャリアアドバイザー）: {ca}
- 求職者: {candidate}

各発言をCAまたは求職者に割り当てて、JSON配列のみ返してください。
形式:
[
  {{"speaker": "CA", "text": "発話内容"}},
  {{"speaker": "求職者", "text": "発話内容"}}
]

文字起こし:
{raw[:40000]}
"""
    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=16000,
        messages=[{"role": "user", "content": prompt}],
    )
    content = response.content[0].text.strip()
    content = re.sub(r"```(?:json)?\s*", "", content).strip()
    try:
        parsed = json.loads(content)
        if isinstance(parsed, list) and parsed:
            return parsed
    except json.JSONDecodeError:
        pass
    return [{"speaker": "不明", "text": raw}]


def utterance_path(meta: dict) -> Path:
    return UTTERANCES_DIR / f"{meta['ca']}_{meta['grip']}_{meta['candidate']}_{meta['meeting_type']}.json"


def save(meta: dict, utterances: list[dict]):
    out = utterance_path(meta)
    speakers = list(set(u["speaker"] for u in utterances))
    out.write_text(json.dumps({
        "ca": meta["ca"],
        "grip": meta["grip"],
        "candidate": meta["candidate"],
        "meeting_type": meta["meeting_type"],
        "format": meta["format"],
        "speakers_detected": speakers,
        "utterance_count": len(utterances),
        "utterances": utterances,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    return speakers, len(utterances)


def main():
    UTTERANCES_DIR.mkdir(parents=True, exist_ok=True)
    files = find_files()
    docx_files = [f for f in files if f["format"] == "docx"]
    txt_files  = [f for f in files if f["format"] == "txt"]

    # docx未処理分（念のため）
    docx_pending = [f for f in docx_files if not utterance_path(f).exists()]
    txt_pending   = [f for f in txt_files  if not utterance_path(f).exists()]

    print(f"docx: {len(docx_pending)}件未処理（API不要）")
    print(f"txt:  {len(txt_pending)}件未処理（Haiku使用）")
    print()

    ok = skip = err = 0
    all_pending = docx_pending + txt_pending
    total = len(all_pending)

    for i, meta in enumerate(all_pending, 1):
        out = utterance_path(meta)
        label = f"[{i:03d}/{total}] {meta['ca']}_{meta['grip']}_{meta['candidate']}_{meta['meeting_type']}({meta['format']})"

        if out.exists():
            print(f"{label} → スキップ")
            skip += 1
            continue

        try:
            if meta["format"] == "docx":
                utterances = parse_docx(meta["path"])
            else:
                print(f"{label} → Haiku話者推定中...")
                utterances = parse_txt_with_haiku(meta["path"], meta["ca"], meta["candidate"])

            speakers, count = save(meta, utterances)
            print(f"{label} → 保存（{count}発話, 話者:{speakers}）")
            ok += 1
        except Exception as e:
            print(f"{label} → エラー: {e}")
            err += 1

    print(f"\n完了: 成功{ok}件 / スキップ{skip}件 / エラー{err}件")


if __name__ == "__main__":
    main()
