const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, Header, Footer,
} = require("docx");

const FONT = "Yu Gothic";
const CW = 9360; // content width (US Letter, 1" margins)
const border = { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB" };
const borders = { top: border, bottom: border, left: border, right: border };
const HEAD_FILL = "1F4E79";
const ALT_FILL = "EEF3F9";

function t(text, opts = {}) { return new TextRun({ text, font: FONT, ...opts }); }
function p(text, opts = {}) {
  return new Paragraph({ children: Array.isArray(text) ? text : [t(text, opts.run || {})],
    spacing: { after: opts.after ?? 120, before: opts.before ?? 0, line: 276 },
    alignment: opts.align, ...opts.paraExtra });
}
function bullet(runs) {
  return new Paragraph({ numbering: { reference: "b", level: 0 },
    children: Array.isArray(runs) ? runs : [t(runs)], spacing: { after: 80, line: 276 } });
}
function cell(content, { w, fill, headerCell = false, align } = {}) {
  const runs = Array.isArray(content) ? content : [t(String(content), { bold: headerCell, color: headerCell ? "FFFFFF" : "000000", size: 19 })];
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 110, right: 110 },
    children: [new Paragraph({ children: runs, alignment: align, spacing: { line: 252 } })],
  });
}
function headerRow(labels, widths) {
  return new TableRow({ tableHeader: true,
    children: labels.map((l, i) => cell(l, { w: widths[i], fill: HEAD_FILL, headerCell: true, align: AlignmentType.CENTER })) });
}
function dataRow(cells, widths, idx) {
  const fill = idx % 2 === 1 ? ALT_FILL : undefined;
  return new TableRow({ children: cells.map((c, i) => {
    const isArr = Array.isArray(c);
    return cell(isArr ? c : String(c), { w: widths[i], fill, align: i === 0 ? AlignmentType.LEFT : undefined });
  }) });
}
function table(widths, header, rows) {
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: widths,
    rows: [headerRow(header, widths), ...rows.map((r, i) => dataRow(r, widths, i))] });
}
function h(text, level) { return new Paragraph({ heading: level, children: [t(text, { bold: true })] }); }
function spacer() { return new Paragraph({ children: [t("")], spacing: { after: 60 } }); }

// ── 補正 before→after（8件）──
const cmpW = [1700, 1500, 1300, 1700, 3160];
const cmpRows = [
  ["岡田", "B → A", "0 → 2", "0 → 1", "✗ → ✓"],
  ["佐藤", "C → B", "0 → 1", "0 → 1", "✗ → ✓"],
  ["加藤", "C → B", "0 → 1", "0 → 1", "✗ → ✗"],
  ["菅",   "B → B", "0 → 2", "1 → 1", "✗ → ✓"],
  ["櫻井", "C → B", "0 → 2", "0 → 1", "✗ → ✓"],
  ["辻",   "B → A", "0 → 1", "0 → 1", "✗ → ✓"],
  ["神谷", "C → B", "0 → 1", "0 → 1", "✗ → ✓"],
  ["児子", "B → B", "0 → 1", "1 → 1", "✗ → ✓"],
];

// ── 課題（n=8）──
const issW = [3300, 1100, 4960];
const issRows = [
  ["認識統一の欠如（価値観・強みの要約＋本人の明示的同意を取らない）", "8 / 8", "ほぼ全件の課題に「価値観要約・認識統一の欠如／明示的同意の未取得」"],
  ["意向変え・視野拡大（気づき付与）が不安定・ゼロの回が多い", "約6 / 8", "佐藤・加藤・菅・櫻井・児子で「意向変え・視野拡大がゼロ／皆無」。初回面談の本来目的"],
  ["条件把握が“浅い”（数値・期待値調整の不足）※未実施ではない", "5 / 8", "加藤「Must数値が未確認」／神谷「年収・休日・雇用形態の確認ゼロ」／辻「期待値調整ゼロ」"],
  ["感情ワードのスルー", "8 / 8", "NG_感情無視 計19件。加藤・菅・櫻井・神谷で各3件"],
  ["CA自己開示・雑談が長くCAが主語化", "約4 / 8", "岡田「共感自己開示の過多による主客転倒」／佐藤「自己開示・雑談が長すぎ」／辻"],
  ["フィラー過多", "8 / 8", "目標30回以下に対し全件55〜97回（神谷97・辻82）"],
  ["名前呼称がほぼゼロ", "7 / 8", "佐藤の1回以外すべて0回"],
];

// ── 改善アクション ──
const actW = [800, 2400, 3360, 2800];
const actRows = [
  ["1", "認識統一の欠如", "面談終盤に「ここまでの理解＝価値観◯◯・強み◯◯で合っていますか？今後はこの方向で進めましょう」と要約→明示的に同意を取るルーティンを固定", "8/8で欠落。前半で引き出した価値観が“合意”に変換されていない"],
  ["2", "視野拡大の不安定さ", "引き出した価値観を「だとすると応募先以外に◯◯という選択肢も」と毎回接続。神谷の成功例（個人目標と紐付け）を全件で再現", "約6/8でゼロ。本来目的かつ“やれる時はやれる”＝再現性の問題"],
  ["3", "条件把握の浅さ", "既にやれている条件確認を、数値（年収下限）＋Must/Better＋期待値調整まで一段深く", "5/8がscore1止まり。未実施ではなく“浅さ”が課題"],
  ["4", "感情ワードのスルー", "ネガ／重い発言は事実確認に進む前に1ターン受け止める（感情に名前をつけて返す）", "8/8・計19件"],
  ["5", "CA主語化・フィラー", "自己開示は要点1つに圧縮、フィラーを意識的に削減", "自己開示過多4/8、フィラー8/8超過"],
];

const children = [];
children.push(new Paragraph({ children: [t("酒井CA 初回面談 個別分析レポート", { bold: true, size: 36, color: "1F4E79" })], spacing: { after: 60 }, alignment: AlignmentType.CENTER }));
children.push(new Paragraph({ children: [t("【改訂版】全文トランスクリプトで再分析（直近10件 → 有効8件）", { size: 22, color: "555555" })], alignment: AlignmentType.CENTER, spacing: { after: 200 } }));

// 改訂理由 box
children.push(new Paragraph({
  shading: { fill: "FBE9D0", type: ShadingType.CLEAR, color: "auto" },
  border: { top: { style: BorderStyle.SINGLE, size: 4, color: "E0A458" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "E0A458" }, left: { style: BorderStyle.SINGLE, size: 4, color: "E0A458" }, right: { style: BorderStyle.SINGLE, size: 4, color: "E0A458" } },
  children: [
    t("【改訂理由】", { bold: true }),
    t("初版はツールが各面談の前半54%前後しかAIに渡しておらず「条件把握10/10未実施」「クロージング10/10未到達」と誤判定していた。トランスクリプトの打ち切り（15,000字→60,000字）を修正し全文で再分析した結果が本版。Drive該当8件も補正後データに更新済み。"),
  ], spacing: { after: 240, before: 60, line: 276 },
}));

children.push(h("1. 分析対象サマリ", HeadingLevel.HEADING_1));
children.push(bullet([t("参照データ：", { bold: true }), t("共有Drive json/（初回面談FBツール出力）＋ utterances/（全文トランスクリプト）")]));
children.push(bullet([t("ツール修正：", { bold: true }), t("AIへ渡す文字起こしの打ち切りを 15,000字 → 60,000字 に拡張（analysis_core.py ほか）")]));
children.push(bullet([t("再分析：", { bold: true }), t("全文（整形24,000〜32,000字）で再採点。行動指標は元々全文計算で不変")]));
children.push(bullet([t("対象：", { bold: true }), t("岡田(A)・佐藤(B)・加藤(B)・菅(B)・櫻井(B)・辻(A)・神谷(B)・児子(B)　→ グレード A2 / B6 / C0（初版の誤った A0/B5/C5 から上方修正）")]));
children.push(bullet([t("除外：", { bold: true }), t("梅本・尾高＝話者分離トランスクリプトがDrive・ローカル双方に存在せず全文再分析が不可能なため除外（2件）")]));
children.push(spacer());
children.push(p([t("補正の効果（before → after・8件）", { bold: true, size: 22 })], { after: 80 }));
children.push(table(cmpW, ["候補者", "グレード", "条件", "認識統一", "次回アポ"], cmpRows));
children.push(p([t("※ 条件は8/8で誤判定（全件0→1〜2）、次回アポは7/8がFalse→True、グレードは6/8が上昇・低下ゼロ。", { italics: true, size: 18, color: "555555" })], { before: 80 }));

children.push(h("2. 酒井の特徴（強み）", HeadingLevel.HEADING_1));
children.push(p([t("全文確認後、むしろ強みが鮮明になった。", { bold: true })]));
children.push(bullet([t("本質を突く価値観の縦深掘りが突出（8/8の最大の強み）：", { bold: true }), t("岡田『なぜそんなに頑張れるんですか？』で3層の価値観を一連で引き出す／辻『本人が気づいていなかった自己像を鮮やかに』／菅『価値観の軸を可視化』")]));
children.push(bullet([t("傾聴・共感・自己開示によるラポール構築：", { bold: true }), t("加藤『自然に話せる雰囲気で短時間に職歴と価値観』／児子『安心して経歴を話せる場』")]));
children.push(bullet([t("条件確認・クロージング・次回アポは実際にはほぼ実施できている：", { bold: true }), t("条件 score2＝岡田・菅・櫻井／次回アポ確定は8件中7件で取得（初版の誤りを補正）")]));
children.push(bullet([t("気づき付与は『できる時はできる』：", { bold: true }), t("神谷『人材紹介という新しい選択肢を個人目標（車・家・ハワイ）と結びつけて提示』（気づき=2）＝スキルが無いのではなく安定しない")]));

children.push(h("3. 改善すべき点（全文でも残った“本物の”課題）", HeadingLevel.HEADING_1));
children.push(table(issW, ["課題", "件数", "根拠"], issRows));

children.push(h("4. 改善アクション（優先度順）", HeadingLevel.HEADING_1));
children.push(table(actW, ["優先", "課題", "具体策", "根拠"], actRows));

children.push(h("5. 総括", HeadingLevel.HEADING_1));
children.push(p("全文で再分析すると、酒井さんの評価は大きく変わった。初版の「条件・クロージングが全滅」はツールの打ち切りによる誤判定で、実際には条件確認も次回アポ取得もほぼできている（次回アポ7/8）。本当の強みは前半の本質的な価値観の縦深掘りとラポール構築で、グレードもA2件を含むB以上に上振れした。"));
children.push(p("一方で全文でも残る本物の課題は、(1) 引き出した価値観を“要約＋明示的合意”に変える認識統一（8/8欠落）、(2) 意向変え・視野拡大の不安定さ、(3) 条件把握の“浅さ”、(4) 感情ワードのスルー。いずれも『能力不足』ではなく面談終盤の型と再現性の問題である。まず『要約して合意を取る』一手を固定するだけで、強い前半が成果（グリップ・意向変え）に直結する。"));
children.push(spacer());
children.push(p([t("補足：", { bold: true }), t("このトランスクリプト打ち切りバグは酒井さん以外の全CA・全初回面談の既存データにも同じ影響が出ている（修正は適用済みのため、今後の分析と再分析は正しくなる）。", { color: "555555", size: 19 })]));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: FONT, color: "1F4E79" },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 2 } } } },
    ],
  },
  numbering: { config: [
    { reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 460, hanging: 260 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1300, right: 1440, bottom: 1300, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [t("酒井CA 初回面談分析レポート（全文再分析版）　/　", { size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "888888" })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("酒井CA_初回面談分析レポート.docx", buf);
  console.log("✅ 酒井CA_初回面談分析レポート.docx を生成しました（" + buf.length + " bytes）");
});
