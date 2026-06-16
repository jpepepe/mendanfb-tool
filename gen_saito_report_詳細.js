const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, Footer,
} = require("docx");

const FONT = "Yu Gothic";
const CW = 9360;
const border = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
const borders = { top: border, bottom: border, left: border, right: border };
const HEAD_FILL = "1F4E79", ALT_FILL = "EEF3F9";

function t(text, opts = {}) { return new TextRun({ text, font: FONT, ...opts }); }
function p(runs, opts = {}) {
  return new Paragraph({ children: Array.isArray(runs) ? runs : [t(runs)],
    spacing: { after: opts.after ?? 110, before: opts.before ?? 0, line: 270 }, alignment: opts.align });
}
function bullet(runs, lvl = 0) {
  return new Paragraph({ numbering: { reference: "b", level: lvl },
    children: Array.isArray(runs) ? runs : [t(runs)], spacing: { after: 70, line: 264 } });
}
function quote(label, body) {
  return new Paragraph({
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: "9DB8D2", space: 8 } },
    indent: { left: 220 }, spacing: { after: 80, line: 264 },
    children: [t(label + "｜", { bold: true, color: "1F4E79", size: 19 }), t(body, { size: 19, italics: true })] });
}
function cell(content, { w, fill, headerCell = false, align } = {}) {
  const runs = Array.isArray(content) ? content
    : [t(String(content), { bold: headerCell, color: headerCell ? "FFFFFF" : "000000", size: 18 })];
  return new TableCell({ borders, width: { size: w, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 100, right: 100 },
    children: [new Paragraph({ children: runs, alignment: align, spacing: { line: 248 } })] });
}
function table(widths, header, rows, headFill = HEAD_FILL) {
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((l, i) => cell(l, { w: widths[i], fill: headFill, headerCell: true, align: AlignmentType.CENTER })) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => cell(c, { w: widths[i], fill: ri % 2 ? ALT_FILL : undefined, align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })) })) ] });
}
function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(text, { bold: true })] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t(text, { bold: true })] }); }

const C = [];
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [t("酒井CA 初回面談 詳細分析レポート", { bold: true, size: 34, color: "1F4E79" })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 180 }, children: [t("多角版 — 行動指標の定量 ＋ 質的深掘り（全文再分析・有効8件）", { size: 21, color: "555555" })] }));
C.push(p([t("対象：", { bold: true }), t("岡田・佐藤・加藤・菅・櫻井・辻・神谷・児子（2026/6/2〜6/12）。各レコードの未活用フィールド（行動指標25種・感情ドリル・自己開示・バックトラッキング・軸別コーチング・フレーズ集）まで分析。")], { after: 160 }));

// 1
C.push(h1("1. 定量行動プロファイル（8件横断）"));
C.push(table([1500, 1300, 1100, 1300, 1320, 1320, 1520],
  ["候補者", "求職者発話比", "感情場面", "スルー率", "価値観深掘り", "バックトラック", "フィラー"],
  [["岡田","46%","18","89%","7","6","60"],["佐藤","52%","36","89%","6","7","59"],["加藤","51%","23","96%","3","4","57"],
   ["菅","60%","16","94%","5","4","58"],["櫻井","51%","14","100%","3","4","55"],["辻","32%","11","91%","6","4","82"],
   ["神谷","34%","17","88%","4","4","97"],["児子","41%","9","89%","4","4","70"],
   ["平均","50%","18","92%","4.8","4.6","67"]]));
C.push(p([t("目標値との照合：", { bold: true })], { before: 100, after: 60 }));
C.push(bullet([t("❌ 感情スルー率 平均92%", { bold: true, color: "B00000" }), t("（目標50%以下）— 8/8で大幅未達。最も深刻かつ最も再現性が高い。1面談平均18回の感情シグナルのうち16.5回をスルー")]));
C.push(bullet([t("❌ フィラー 平均67回", { bold: true }), t("（目標30以下）— 8/8未達（神谷97・辻82）")]));
C.push(bullet([t("❌ 名前呼称 ほぼ0回", { bold: true }), t("（目標3+）— 佐藤の1回以外すべて0")]));
C.push(bullet([t("⚠️ バックトラッキング 平均4.6", {}), t("（目標5+）— 達成は岡田・佐藤のみ")]));
C.push(bullet([t("✅ 価値観深掘り 平均4.8・ポジティブ反応 全件5+", { bold: true, color: "1E7A1E" }), t("— ラポール・傾聴は良好")]));
C.push(p([t("※「縦深掘り最大」は全8件1だが、これは指標の仕様（CAが答えを挟まず連続質問した時だけカウント＝健全な一問一答では常に1）。酒井さんの深掘りの弱さを示すものではない。", { size: 18, italics: true, color: "555555" })], { before: 60 }));

// 2
C.push(h1("2. 感情スルーの解剖（スルー率92%の中身）"));
C.push(p("8件の missed_scenes（計37場面）を分析すると、スルーの仕方が4類型に分かれる。これが92%の正体。"));
C.push(h2("類型① ポジティブ変換・即否定でスルー（最頻・酒井の“クセ”）"));
C.push(p("ネガティブな自己開示を、励まし・称賛で打ち消してしまう。", { after: 60 }));
C.push(quote("辻", "求職者「人見知りでコミュ力もダメだなと」→ 酒井「コミ力あるじゃないですか」"));
C.push(quote("佐藤", "「自己肯定感だいぶ低い方」→「すごいですね（笑いに転換）」"));
C.push(quote("神谷", "「フラストレーションも溜まっていく」→「めちゃくちゃすごいですよ」"));
C.push(h2("類型② 相槌のみでスルー"));
C.push(quote("櫻井(100%)", "「女性はキャリアという概念がない職場だった」→「うん。うん。うん。」"));
C.push(quote("菅", "「（漫画家を）突きつけられまして」→「うん。うん。うん」"));
C.push(h2("類型③ 事実質問へ即ジャンプ"));
C.push(quote("岡田", "「最後3人だけでやめれない」→「どれぐらい契約取れたんですか？」"));
C.push(quote("加藤", "「気づいたら4年走ってた／体が完全になってた（燃え尽き）」→「うん。なるほど。（退職理由へ）」"));
C.push(h2("類型④ 自分のフレームへ即・要約"));
C.push(quote("児子", "「できるようになるのが楽しくて」→「自己成長にやりがいを感じるタイプなんですね」（本人が言語化する前にCAがラベル）"));
C.push(p([t("→ スルーされた感情語は「体を壊した」「燃え尽き」「ラストチャンス」「突きつけられた」「久しぶりに褒められた」など、", {}), t("グリップの最重要ポイントばかり。", { bold: true }), t("ここを1ターン拾えるかが課題の核心。")], { before: 80 }));
C.push(p([t("できている深掘り（強み＝価値観ラベリング）：", { bold: true, color: "1E7A1E" }), t("辻「料理とトロンボーンから“支えたい”を見抜き言語化」／神谷「陸上の“一瞬に賭ける”本質を言語化」。強み（価値観の言語化）と弱み（感情の深掘り）は表裏——価値観は拾えるのに、その手前の“生の感情”を拾えていない。")]));

// 3
C.push(h1("3. 自己開示の質"));
C.push(bullet([t("量は多い", { bold: true }), t("（平均6.8回、岡田・菅9回）。場を和ませラポールには寄与")]));
C.push(bullet([t("ただし8件中6件で「過多／改善余地」。", { bold: true }), t("特徴は“共感型”でなく“場を和ませる型・自分語り型”。櫻井「自社・経歴説明の文脈ばかりで感情に共鳴した自己開示が乏しい」")]));
C.push(bullet([t("求職者の発話機会を圧迫", {}), t("（辻32%・神谷34%の低発話比率と連動）")]));
C.push(p([t("改善：", { bold: true }), t("自己開示の“量”でなく“狙い”。場を和ませる開示を減らし、求職者がネガティブを話した直後に「私も似た経験が」と共感を渡す自己開示へ振り替える。")]));

// 4
C.push(h1("4. バックトラッキングの質"));
C.push(bullet([t("短期ミラーリング（直後の言い換え）はできる", { bold: true }), t("が、数ターン後に伏線回収する長期引用が弱い（8件中6件で指摘）")]));
C.push(bullet([t("好例：神谷", { bold: true, color: "1E7A1E" }), t("「ロープレ全国9位・500人中23位を繰り返し引用して自己肯定感を高めた」← この“時間差引用”を全件で再現したい")]));
C.push(p([t("改善：", { bold: true }), t("前半で出た印象的なワードをメモし、後半で「さっき◯◯っておっしゃってましたよね」と引用する。")]));

// 5
C.push(h1("5. 軸別コーチング（5軸 × 8件の弱点パターン）"));
C.push(p("弱点が軸ごとにほぼ同一文言で8件に出現＝完全に固定化したクセ。", { after: 80 }));
C.push(table([1300, 1100, 6960], ["軸", "スコア傾向", "共通する弱点（8件横断）"],
  [["意向","多くs2","価値観は引き出すが「意向変え・視野拡大」に接続しない。応募企業の範囲で意向確認を終える"],
   ["適正","全件s2","強みは把握するが「あなたの強みはこれですね」と本人に返して確認するフィードバックが無い"],
   ["条件","s1〜2","確認はするがMust/Betterの優先順位・年収下限の数値・期待値調整が無い（“浅さ”）"],
   ["認識統一","全件s1","①価値観・強みの要約への明示的同意なし、②複数社並行方針の合意なし（8件で同じ2点）"],
   ["気づき","s1〜2","気づき提供がCAの一方向情報提供になり、「聞いてどう感じましたか？」の内省・反応引き出しが無い"]]));
C.push(p([t("→ 共通構造は", {}), t("「引き出す（得意）→ 本人に返して合意/接続する（欠落）」。", { bold: true }), t("前半の傾聴で得た材料を、後半で“本人の言葉での確認・視野拡大”に変換するステップが丸ごと抜けている。")], { before: 80 }));

// 6
C.push(h1("6. 明日使えるフレーズ集"));
C.push(p("各候補者に最適化された next_phrases（計40個）から汎用フレーズを抽出。", { after: 60 }));
C.push(bullet([t("感情ワード直後：", { bold: true }), t("「今“◯◯”っておっしゃいましたよね。その時って、どんな気持ちでしたか？」")]));
C.push(bullet([t("縦の深掘り2回目：", { bold: true }), t("「それって、もともとそういうタイプ？ 何かきっかけがあった感じですか？」")]));
C.push(bullet([t("共感の自己開示：", { bold: true }), t("「私も転職の時、何が合うか分からず遠回りしたんです。だから今の気持ち分かる気がします。◯◯さんはどうでした？」")]));
C.push(bullet([t("時間差バックトラッキング：", { bold: true }), t("「さっき“気づいたら4年走ってた”っておっしゃったの印象的で。それ、次に求めるものと繋がってたりします？」")]));
C.push(bullet([t("クロージング問いかけ：", { bold: true }), t("「今日話してみて、◯◯さんが一番大事にしたいことって、言葉にするとどんな感じになりそうですか？」")]));
C.push(bullet([t("認識統一（締めの合意）：", { bold: true, color: "B00000" }), t("「今日伺った“◯◯（価値観）・△△（強み）”で合っていますか？ では次回はこの方向で、複数社を並行して見ていきましょう。」")]));

// 7
C.push(h1("7. 総括"));
C.push(p([t("強みは「価値観ラベリング」", { bold: true }), t("＝複数エピソードから一貫した価値観を見抜き言語化する力（good_scenes・深掘り4.8・ポジティブ反応5+で裏付け）。")]));
C.push(p([t("弱みは一点に集約される——", {}), t("「生の感情」と「本人の合意・視野拡大」を素通りすること。", { bold: true })]));
C.push(p("最大の数値的証拠が感情スルー率92%で、中身は4類型（ポジティブ変換／相槌のみ／事実質問へジャンプ／自分のフレームへ要約）。とりわけ“ネガティブな自己開示を励ましで打ち消すクセ”は辻・佐藤・神谷で繰り返され、最も矯正効果が高い。軸別では「引き出す→返す」の“返す”側（適性FB・認識統一・気づきの双方向化）が8件で固定的に欠落。"));
C.push(p([t("伸ばし方は明快：", { bold: true }), t("①感情語が出たら1ターンだけ「その時どんな気持ち？」で受ける（励まし禁止）、②締めに「◯◯で合ってますか？→この方向で複数社」の合意フレーズを固定。この2つで、強い前半（価値観把握）が初めてグリップ＝意向変えに変換される。")]));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, font: FONT, color: "1F4E79" },
        paragraph: { spacing: { before: 260, after: 120 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 2 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 21, bold: true, font: FONT, color: "2E5E8C" },
        paragraph: { spacing: { before: 140, after: 60 }, outlineLevel: 1 } } ] },
  numbering: { config: [{ reference: "b", levels: [
    { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 240 } } } } ] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1300, right: 1440, bottom: 1300, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [t("酒井CA 初回面談 詳細分析レポート（多角版）　/　", { size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "888888" })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync("酒井CA_初回面談分析_詳細版.docx", buf); console.log("✅ 酒井CA_初回面談分析_詳細版.docx 生成（" + buf.length + " bytes）"); });
