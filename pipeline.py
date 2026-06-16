# -*- coding: utf-8 -*-
"""
面談文字起こし分析パイプライン
- docx: 話者ラベルをパースして分離
- txt: Claude APIで話者推定
- 全ファイルをClaude APIでルーブリック採点しJSON保存
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

# ── 設定 ──────────────────────────────────────────────
TRANSCRIPT_BASE = Path("/Users/takahashijunpei/文字起こし")
OUTPUT_DIR = Path("/Users/takahashijunpei/Downloads/人材紹介分析フォルダ/output/json")
UTTERANCES_DIR = Path("/Users/takahashijunpei/Downloads/人材紹介分析フォルダ/output/utterances")
RUBRIC_PATH = Path("/Users/takahashijunpei/Downloads/人材紹介分析フォルダ/rubric_初回面談.md")
TARGET_MEETING_TYPE = "初回面談"  # 今回の対象

FILLER_PATTERN = re.compile(r'(?:えー+|あのー+|えっと+|まあ+|うーん+|んー+|あー+|そのー+|なんか(?!ら)|ちょっと(?=\s|、|。|$))')
NAME_CALL_SUFFIXES = ["さん"]

client = anthropic.Anthropic()

# ── ファイル探索・パース ───────────────────────────────
def find_files(meeting_type: str) -> list[dict]:
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
            if grip and ca and mt == meeting_type:
                files.append({
                    "path": os.path.join(root, fn),
                    "ca": ca,
                    "grip": grip,
                    "candidate": candidate,
                    "meeting_type": mt,
                    "format": ext,
                })
    return files


# ── docxパーサー ──────────────────────────────────────
def _detect_speakers(full_text: str) -> list[str]:
    """全文から話者名候補を頻度で検出し、上位2名を返す"""
    # 「2〜10文字の日本語系文字 + コロン + スペース」パターンで候補抽出
    candidates = re.findall(r"(?<![^\s。！？\n])([぀-鿿a-zA-Zー]{2,10}):\s", full_text)
    from collections import Counter
    freq = Counter(candidates)
    # 上位2名（最低5回以上出現）
    speakers = [name for name, cnt in freq.most_common(10) if cnt >= 3][:2]
    return speakers


def parse_docx(path: str) -> list[dict]:
    """話者ラベル付きdocxを [{speaker, text}, ...] に変換"""
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

    # ヘッダー・タイムスタンプ除去
    full_text = re.sub(r"^\d{2}:\d{2}:\d{2}\n", "", full_text, flags=re.MULTILINE)
    full_text = re.sub(r"^\d{4}年.+\n", "", full_text, flags=re.MULTILINE)
    full_text = re.sub(r"^会議\s.+\n", "", full_text, flags=re.MULTILINE)

    # 話者名を自動検出
    speakers = _detect_speakers(full_text)
    if not speakers:
        return [{"speaker": "不明", "text": full_text}]

    # 検出した話者名で厳密にsplit（話者名 + コロン + スペースの手前で分割）
    speaker_pattern = re.compile(
        r"(?:" + "|".join(re.escape(s) for s in speakers) + r"):\s*"
    )

    utterances = []
    matches = list(speaker_pattern.finditer(full_text))
    for i, m in enumerate(matches):
        # 話者名はマッチした文字列から抽出
        matched_str = m.group()
        speaker = matched_str.rstrip(": \t")
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(full_text)
        text = full_text[start:end].strip().replace("\n", " ")
        text = re.sub(r"\s+", " ", text).strip()
        if text:
            utterances.append({"speaker": speaker, "text": text})

    return utterances


# ── txtパーサー（Claude APIで話者推定） ──────────────────
def parse_txt(path: str, ca: str, candidate: str) -> list[dict]:
    """話者ラベルなしtxtをClaude APIで話者推定して [{speaker, text}, ...] に変換"""
    with open(path, encoding="utf-8") as f:
        raw = f.read()

    # 行番号プレフィックス（「1\t」形式）を除去
    raw = re.sub(r"^\d+\t", "", raw, flags=re.MULTILINE)

    prompt = f"""以下は人材紹介会社の面談音声文字起こしです。
話者は2名います：
- CA（キャリアアドバイザー）: {ca}
- 求職者: {candidate}

文字起こし全文を読み、各発言をCAまたは求職者に割り当ててください。
句読点なし・誤変換あり・話者ラベルなしですが、文脈から判断してください。

出力はJSON配列のみ。他の文字は一切含めないでください。
形式:
[
  {{"speaker": "CA", "text": "発話内容"}},
  {{"speaker": "求職者", "text": "発話内容"}},
  ...
]

文字起こし:
{raw[:40000]}
"""
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=16000,
        messages=[{"role": "user", "content": prompt}],
    )
    content = response.content[0].text.strip()
    # Markdownコードブロックを除去
    content = re.sub(r"```(?:json)?\s*", "", content).strip()
    try:
        parsed = json.loads(content)
        if isinstance(parsed, list) and parsed:
            return parsed
    except json.JSONDecodeError:
        pass
    # フォールバック：話者なし全文
    return [{"speaker": "不明", "text": raw}]


# ── 行動指標の計算（ルールベース） ──────────────────────
def compute_behavior_metrics(utterances: list[dict], ca: str, candidate_name: str) -> dict:
    # 話者名は「下川玖瑠実」「馬場さやか」のようにフルネームのケースがあるため前方一致で判定
    ca_base = re.sub(r"さん$", "", ca)
    cand_base = re.sub(r"さん$", "", candidate_name)

    def is_ca(speaker: str) -> bool:
        return ca_base in speaker or speaker in ("CA", "キャリアアドバイザー")

    def is_cand(speaker: str) -> bool:
        return cand_base in speaker or speaker in ("求職者",)

    ca_texts = [u["text"] for u in utterances if is_ca(u["speaker"])]
    cand_texts = [u["text"] for u in utterances if is_cand(u["speaker"])]
    all_ca = " ".join(ca_texts)
    all_cand = " ".join(cand_texts)

    # 求職者発話比率
    ca_len = len(all_ca)
    cand_len = len(all_cand)
    total_len = ca_len + cand_len
    cand_ratio = round(cand_len / total_len, 3) if total_len > 0 else 0.0

    # 会話回数（10文字以上の実質的な発話で話者が交代した回数）
    # 「はい」「。」などの相づちはカウントしない
    turn_count = 0
    prev_speaker = None
    for u in utterances:
        if len(u["text"].replace(" ", "")) < 10:
            continue
        sp = "CA" if is_ca(u["speaker"]) else ("求職者" if is_cand(u["speaker"]) else None)
        if sp and sp != prev_speaker:
            turn_count += 1
            prev_speaker = sp

    # 求職者名の呼称回数（CAの発話内で「〇〇さん」）
    # candidate_name から「さん」を除いた姓を使う
    name_base = re.sub(r"さん$", "", candidate_name)
    name_pattern = re.compile(re.escape(name_base) + r"さん")
    name_count = sum(len(name_pattern.findall(t)) for t in ca_texts)

    # フィラー回数（CAの発話内）
    filler_count = sum(len(FILLER_PATTERN.findall(t)) for t in ca_texts)

    return {
        "求職者発話比率": cand_ratio,
        "会話回数": turn_count,
        "名前呼称回数": name_count,
        "フィラー回数": filler_count,
    }


# ── Claude APIによるルーブリック採点 ─────────────────────
def extract_with_claude(utterances: list[dict], meta: dict, rubric: str) -> dict:
    # 発話を読みやすい形式に整形（長すぎる場合は末尾カット）
    transcript_formatted = "\n".join(
        f"[{u['speaker']}] {u['text']}" for u in utterances
    )
    if len(transcript_formatted) > 60000:
        transcript_formatted = transcript_formatted[:60000] + "\n...(以下省略)"

    prompt = f"""あなたは人材紹介会社のマネージャーです。以下のルーブリックに従い、面談文字起こしを分析してください。

## ルーブリック
{rubric}

## メタ情報
- CA名: {meta['ca']}
- 求職者名: {meta['candidate']}
- グリップランク: {meta['grip']}（分析には使わない。答え合わせ用）
- 面談種別: {meta['meeting_type']}
- 形式: {meta['format']}

## 面談文字起こし
{transcript_formatted}

## 出力指示
以下のJSONスキーマで出力してください。JSON以外の文字は一切含めないでください。

{{
  "grip_drivers": {{
    "意向": {{"score": 0, "evidence": ["引用1（30字以内）", "引用2（任意）"], "confidence": "高|中|低"}},
    "適正": {{"score": 0, "evidence": [], "confidence": "高|中|低"}},
    "条件": {{"score": 0, "evidence": [], "confidence": "高|中|低"}},
    "認識統一": {{"score": 0, "evidence": [], "confidence": "高|中|低"}},
    "気づき": {{"score": 0, "evidence": [], "confidence": "高|中|低"}}
  }},
  "behaviors": {{
    "深掘り_価値観": 0,
    "深掘り_実績": 0,
    "バックトラッキング": 0,
    "共感自己開示": 0,
    "ポジティブ反応": 0,
    "NG_急かし": 0,
    "NG_権威": 0,
    "NG_感情無視": 0,
    "フェーズ網羅": ["冒頭"],
    "MUST提案": false,
    "MUST同意": false,
    "次回アポ確定": false
  }},
  "notes": "文字起こし品質・話者分離の懸念・特記事項など"
}}

スコア基準:
- 3: 明確な根拠引用あり＋求職者の同意・反応が確認できる
- 2: 把握・実施できているが確認が弱い or 求職者の反応が「はい」止まり
- 1: 触れているが浅い、または一方的な説明のみ
- 0: 未実施

MUST提案: エンジニア意向が弱い→別職種提案、エンジニア意向が強い→エンジニア路線確認、どちらかができていればtrue
"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    content = response.content[0].text.strip()
    json_match = re.search(r"\{.*\}", content, re.DOTALL)
    if json_match:
        return json.loads(json_match.group())
    raise ValueError(f"JSON解析失敗: {content[:200]}")


# ── メイン処理 ────────────────────────────────────────
def process_file(file_meta: dict, rubric: str) -> dict:
    path = file_meta["path"]
    ca = file_meta["ca"]
    candidate = file_meta["candidate"]
    fmt = file_meta["format"]

    print(f"  → 話者分離中...")
    if fmt == "docx":
        utterances = parse_docx(path)
        base_confidence = "高"
    else:
        utterances = parse_txt(path, ca, candidate)
        base_confidence = "中"

    print(f"  → 行動指標計算中...")
    behavior_metrics = compute_behavior_metrics(utterances, ca, candidate)

    print(f"  → ルーブリック採点中...")
    claude_result = extract_with_claude(utterances, file_meta, rubric)

    # 結果をマージ
    result = {
        "ca": ca,
        "grip": file_meta["grip"],
        "candidate": candidate,
        "meeting_type": file_meta["meeting_type"],
        "format": fmt,
        "grip_drivers": claude_result.get("grip_drivers", {}),
        "behaviors": {
            **behavior_metrics,
            **claude_result.get("behaviors", {}),
        },
        "notes": claude_result.get("notes", ""),
    }

    # txtの場合、grip_driversの確信度を強制的に「中」に設定
    if fmt == "txt":
        for key in result["grip_drivers"]:
            result["grip_drivers"][key]["confidence"] = "中"

    return result


def output_path(file_meta: dict) -> Path:
    ca = file_meta["ca"]
    grip = file_meta["grip"]
    candidate = file_meta["candidate"]
    mt = file_meta["meeting_type"]
    return OUTPUT_DIR / f"{ca}_{grip}_{candidate}_{mt}.json"


def save_utterances(file_meta: dict, utterances: list[dict]):
    """話者分離テキストをutterances/に保存（API不要の再分析用）"""
    UTTERANCES_DIR.mkdir(parents=True, exist_ok=True)
    ca = file_meta["ca"]
    grip = file_meta["grip"]
    candidate = file_meta["candidate"]
    mt = file_meta["meeting_type"]
    utt_path = UTTERANCES_DIR / f"{ca}_{grip}_{candidate}_{mt}.json"
    if not utt_path.exists():
        speakers = list(set(u["speaker"] for u in utterances))
        utt_path.write_text(json.dumps({
            "ca": ca, "grip": grip, "candidate": candidate,
            "meeting_type": mt, "format": file_meta["format"],
            "speakers_detected": speakers,
            "utterance_count": len(utterances),
            "utterances": utterances,
        }, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    UTTERANCES_DIR.mkdir(parents=True, exist_ok=True)

    rubric = RUBRIC_PATH.read_text(encoding="utf-8")

    files = find_files(TARGET_MEETING_TYPE)
    print(f"対象ファイル数: {len(files)}件 ({TARGET_MEETING_TYPE})")

    ok = skip = err = 0
    for i, file_meta in enumerate(files, 1):
        out_path = output_path(file_meta)
        label = f"[{i:03d}/{len(files)}] {file_meta['ca']}_{file_meta['grip']}_{file_meta['candidate']}"

        if out_path.exists():
            print(f"{label} → スキップ（処理済み）")
            skip += 1
            continue

        print(f"{label} → 処理開始")
        try:
            # 話者分離
            print(f"  → 話者分離中...")
            fmt = file_meta["format"]
            if fmt == "docx":
                utterances = parse_docx(file_meta["path"])
            else:
                utterances = parse_txt(file_meta["path"], file_meta["ca"], file_meta["candidate"])

            # utterances保存（API不要の再分析用）
            save_utterances(file_meta, utterances)

            # 行動指標計算
            print(f"  → 行動指標計算中...")
            behavior_metrics = compute_behavior_metrics(utterances, file_meta["ca"], file_meta["candidate"])

            # Claude採点
            print(f"  → ルーブリック採点中...")
            claude_result = extract_with_claude(utterances, file_meta, rubric)

            result = {
                "ca": file_meta["ca"],
                "grip": file_meta["grip"],
                "candidate": file_meta["candidate"],
                "meeting_type": file_meta["meeting_type"],
                "format": fmt,
                "grip_drivers": claude_result.get("grip_drivers", {}),
                "behaviors": {**behavior_metrics, **claude_result.get("behaviors", {})},
                "notes": claude_result.get("notes", ""),
            }
            if fmt == "txt":
                for key in result["grip_drivers"]:
                    result["grip_drivers"][key]["confidence"] = "中"

            out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"  → 保存: {out_path.name}")
            ok += 1
        except Exception as e:
            print(f"  → エラー: {e}")
            err += 1

    print(f"\n完了: 成功{ok}件 / スキップ{skip}件 / エラー{err}件")


if __name__ == "__main__":
    main()
