# -*- coding: utf-8 -*-
"""
分析コア関数（面談FB_ツール.py と ダッシュボードで共用）
"""
import json, re
from pathlib import Path
import anthropic

OUTPUT_JSON = Path(__file__).parent / "output" / "json"
OUTPUT_UTT  = Path(__file__).parent / "output" / "utterances"
OUTPUT_JSON.mkdir(parents=True, exist_ok=True)
OUTPUT_UTT.mkdir(parents=True, exist_ok=True)


def score_with_claude(utterances, ca_name, cand_name, fmt, client):
    """Call 1: ルーブリックスコアリング"""
    transcript = '\n'.join(f"[{u['speaker']}] {u['text']}" for u in utterances)[:15000]
    prompt = f"""あなたは人材紹介会社の敏腕トレーナーです。以下の面談文字起こしを分析し、
CAへの具体的なフィードバックをJSONで生成してください。

## ★最重要：この会社の初回面談の目的（評価の大前提）
この人材紹介会社では、初回面談のゴールは「応募企業への魅力づけ・囲い込み」ではありません。
応募企業はあくまで求職者に来てもらうための入口であり、面談の本当の目的は次の2つです。
1. 求職者の価値観・本音を引き出し、意向を広げる（意向変え）
2. 応募企業に固執させず、複数企業・新しい選択肢に目を向けてもらう（視野拡大）
したがって全軸を以下の方針で評価してください。
- 応募企業への魅力づけ・志望度の確認・囲い込みは評価対象外。やっていなくても減点しない。
- 応募企業だけに話を限定している場合はマイナス、視野を広げられている場合はプラス。
- 求職者の価値観を引き出し、新しい視点・他の可能性に気づかせられたかを高く評価する。

## 各評価軸の定義（この定義に厳密に従うこと）
- 意向把握：価値観・やりがいを引き出し、応募企業に固執させず意向を広げられたか（意向変えの土台）
- 適正把握：経験・強み・適性を具体的に把握できたか
- 条件把握：Must/Betterを確認し、期待値を調整できたか
- 認識統一：①求職者の価値観・強みの要約に本人の明示的同意が得られたか、②今後のキャリアの方向性（複数社を見て進めること等）について合意できたか。※応募企業への同意ではない
- 気づき付与：求職者に新しい視点・他の選択肢の可能性に気づかせられたか

## スコア基準（0〜3点）
- 3: 根拠引用あり＋求職者の明示的な同意・反応が確認できる
- 2: 把握できているが確認が弱い or 求職者が「はい」止まり
- 1: 触れているが浅い・一方的な説明のみ
- 0: 未実施

## グリップA基準値
- ポジティブ反応：5回以上 / 価値観深掘り：4回以上
- バックトラッキング：5回以上 / 感情スルー率：50%以下 / 縦深掘り：同テーマ3回以上連続

## メタ情報
CA名: {ca_name} / 求職者名: {cand_name} / 形式: {fmt}

## 面談文字起こし
{transcript}

## 出力（JSONのみ）
{{
  "grip_drivers": {{
    "意向":    {{"score":0,"evidence":[],"strength":"","weakness":"","next_action":""}},
    "適正":    {{"score":0,"evidence":[],"strength":"","weakness":"","next_action":""}},
    "条件":    {{"score":0,"evidence":[],"strength":"","weakness":"","next_action":""}},
    "認識統一":{{"score":0,"evidence":[],"strength":"","weakness":"","next_action":""}},
    "気づき":  {{"score":0,"evidence":[],"strength":"","weakness":"","next_action":""}}
  }},
  "behaviors": {{
    "深掘り_価値観":0,"深掘り_実績":0,
    "バックトラッキング":0,"共感自己開示":0,"ポジティブ反応":0,
    "NG_急かし":0,"NG_権威":0,"NG_感情無視":0,
    "フェーズ網羅":["冒頭"],
    "MUST提案":false,"MUST同意":false,"次回アポ確定":false
  }},
  "overall": {{
    "grade":"A",
    "grade_reason":"グレード判定の理由を1文で",
    "top_strength":"この面談の最大の強みを具体的に1文で",
    "top_issues":[
      {{"issue":"課題1","detail":"何が問題か","fix":"代わりにこう言う（例文）"}},
      {{"issue":"課題2","detail":"","fix":""}},
      {{"issue":"課題3","detail":"","fix":""}}
    ],
    "missed_moment":"感情ワードをスルーした最も惜しかった場面",
    "best_exchange":"最も良かったやり取り",
    "closing_eval":"クロージング評価",
    "one_thing":"次の面談で必ず1つ試してほしいこと（例文付き）"
  }},
  "notes":""
}}
grade基準: S=全軸2.5以上, A=10以上, B=7〜9, C=4〜6, D=3以下"""

    resp = client.messages.create(
        model='claude-sonnet-4-6', max_tokens=4000,
        messages=[{'role': 'user', 'content': prompt}])
    content = re.sub(r'```(?:json)?\s*', '', resp.content[0].text.strip()).strip()
    try:
        return json.loads(content)
    except Exception:
        m = re.search(r'\{.*\}', content, re.DOTALL)
        if m:
            try: return json.loads(m.group())
            except: pass
    return {}


def deep_analysis_with_claude(utterances, ca_name, cand_name, client):
    """Call 2: 深掘り・自己開示・バックトラッキング分析"""
    transcript = '\n'.join(f"[{u['speaker']}] {u['text']}" for u in utterances)[:14000]
    prompt = f"""あなたは人材紹介会社の面談コーチです。以下の面談文字起こしを分析し、
3つの観点から詳細なフィードバックをJSONで返してください。

CA名: {ca_name} / 求職者名: {cand_name}

## 面談文字起こし
{transcript}

## 出力（JSONのみ）
{{
  "emotion_drill_analysis": {{
    "summary": "感情深掘り全体の評価（2〜3文）",
    "missed_scenes": [
      {{"cd_text":"スルーされた求職者発話","emotion_word":"感情ワード",
        "ca_actual":"実際のCAの返し","ca_suggested":"こう返すべきだった（例文）","why":"理由"}}
    ],
    "good_scenes": [
      {{"cd_text":"うまく深掘りできた場面","ca_text":"CAの深掘り発話","why_good":"理由"}}
    ],
    "vertical_drill_comment": "縦の深掘りの評価"
  }},
  "self_disclosure_analysis": {{
    "summary": "自己開示全体の評価",
    "found_scenes": [
      {{"ca_text":"CAの自己開示","timing_eval":"良い/普通/改善余地あり","effect":"効果"}}
    ],
    "missed_opportunities": [
      {{"cd_text":"ここで自己開示できた","ca_suggested":"こう自己開示できた（例文）"}}
    ],
    "advice": "アドバイス"
  }},
  "backtrack_analysis": {{
    "summary": "バックトラッキング全体の評価",
    "found_scenes": [
      {{"ca_text":"CAの発話","referenced_cd":"参照した求職者発言","effect":"良い/普通/惜しい"}}
    ],
    "missed_opportunities": [
      {{"cd_keyword":"引用できたキーワード","ca_suggested":"こう使えた（例文）"}}
    ],
    "advice": "アドバイス"
  }},
  "next_phrases": [
    {{"situation":"感情ワードが出た直後","phrase":"フレーズ例","why":"理由"}},
    {{"situation":"縦の深掘り2回目","phrase":"","why":""}},
    {{"situation":"自己開示のタイミング","phrase":"","why":""}},
    {{"situation":"バックトラッキングで引用","phrase":"","why":""}},
    {{"situation":"クロージング問いかけ","phrase":"","why":""}}
  ]
}}"""

    try:
        resp = client.messages.create(
            model='claude-sonnet-4-6', max_tokens=8000,
            messages=[{'role': 'user', 'content': prompt}])
        content = re.sub(r'```(?:json)?\s*', '', resp.content[0].text.strip()).strip()
        if resp.stop_reason == 'max_tokens':
            # 途中で切れた場合は閉じ括弧を補完して無理やりパースを試みる
            content = content.rstrip(',\n ')
            for _ in range(10):
                content += '}'
            content += ']}}}'
        try:
            return json.loads(content)
        except Exception:
            m = re.search(r'\{.*\}', content, re.DOTALL)
            if m:
                try: return json.loads(m.group())
                except: pass
    except Exception:
        pass
    return {}


def save_analysis(ca, grip, candidate, meeting_type, fmt,
                  behaviors_calc, claude_result, deep_result=None):
    """分析結果をJSONに保存（utterancesは別途保存済みの前提）"""
    safe_grip = grip if grip != '未入力' else 'X'
    key = f"{ca}_{safe_grip}_{candidate}_{meeting_type}"
    json_path = OUTPUT_JSON / f"{key}.json"
    json_path.write_text(json.dumps({
        "ca": ca, "grip": safe_grip, "candidate": candidate,
        "meeting_type": meeting_type, "format": fmt,
        "grip_drivers":  claude_result.get('grip_drivers', {}),
        "behaviors":     {**behaviors_calc, **claude_result.get('behaviors', {})},
        "overall":       claude_result.get('overall', {}),
        "notes":         claude_result.get('notes', ''),
        "emotion_drill_analysis":  (deep_result or {}).get('emotion_drill_analysis', {}),
        "self_disclosure_analysis":(deep_result or {}).get('self_disclosure_analysis', {}),
        "backtrack_analysis":      (deep_result or {}).get('backtrack_analysis', {}),
        "next_phrases":            (deep_result or {}).get('next_phrases', []),
    }, ensure_ascii=False, indent=2), encoding='utf-8')
    return json_path


def find_utterances_file(ca, grip, candidate, meeting_type):
    """対応するutterancesファイルを探す"""
    safe_grip = grip if grip != '未入力' else 'X'
    key = f"{ca}_{safe_grip}_{candidate}_{meeting_type}"
    p = OUTPUT_UTT / f"{key}.json"
    return p if p.exists() else None
