const pptx = require('pptxgenjs');
const pres = new pptx();
pres.layout = 'LAYOUT_16x9'; // 10" × 5.625"

// ─── Palette（元FMT準拠） ──────────────────────────────────
const CORAL = 'DD7E6B';  // サーモン/メインブランドカラー
const DKCRL = 'AA5848';  // ダークコーラル
const GREEN = '156818';  // 濃緑（ポジティブ）
const DKGRN = '115000';  // さらに暗い緑
const LTGRN = 'D9EAD3';  // 薄緑背景
const LTRED = 'F4CCCC';  // 薄赤背景
const RED   = 'DD001B';  // 赤（ネガティブ/コスト）
const AMBER = 'E65100';  // アンバー（補助色）
const LTAM  = 'FFF3E0';  // 薄アンバー
const LTBL  = 'C9DAF8';  // 薄青
const WHITE = 'FFFFFF';
const OFFWH = 'F3F3F3';  // オフホワイト背景
const DTXT  = '333333';  // ダークテキスト
const MTXT  = '595959';  // ミディアムテキスト
const MUTED = '999999';  // ミュートテキスト
const LGREY = 'CCCCCC';  // ライトグレー罫線
const POSGR = '188038';  // 黒字/ポジ数値

// ─── Helpers ──────────────────────────────────────────────
function card(s, x, y, w, h, fill, lineColor) {
  s.addShape(pres.ShapeType.rect, { x, y, w, h,
    fill: { color: fill || WHITE },
    line: lineColor ? { color: lineColor, width: 0.75 } : { type: 'none' } });
}
function txt(s, text, x, y, w, h, opts) {
  s.addText(text, { x, y, w, h, fontFace: 'Calibri', ...opts });
}
// 全幅コーラルヘッダーバー（元FMT）
function pgTitle(s, text, secNum) {
  card(s, 0, 0, 10, 0.7, CORAL);
  if (secNum !== undefined) {
    s.addShape(pres.ShapeType.ellipse, { x: 0.12, y: 0.11, w: 0.48, h: 0.48,
      fill: { color: WHITE }, line: { type: 'none' } });
    txt(s, String(secNum), 0.12, 0.11, 0.48, 0.48,
      { fontSize: 14, bold: true, color: CORAL, align: 'center', valign: 'middle' });
    txt(s, text, 0.72, 0, 9.2, 0.7,
      { fontSize: 16, bold: true, color: WHITE, valign: 'middle' });
  } else {
    txt(s, text, 0.2, 0, 9.7, 0.7,
      { fontSize: 16, bold: true, color: WHITE, valign: 'middle' });
  }
}
function slab(s, text, x, y, w, h, col) {
  card(s, x, y, w, h, col || CORAL);
  txt(s, text, x, y, w, h,
    { fontSize: 9, bold: true, color: WHITE, align: 'center', valign: 'middle', charSpacing: 0.5 });
}
function dot(s, x, y, col) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: 0.13, h: 0.13,
    fill: { color: col || GREEN }, line: { type: 'none' } });
}
function numCircle(s, n, x, y, col) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: 0.48, h: 0.48,
    fill: { color: col || CORAL }, line: { type: 'none' } });
  txt(s, String(n), x, y, 0.48, 0.48,
    { fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}
function hrule(s, x, y, w, col) {
  s.addShape(pres.ShapeType.line, { x, y, w, h: 0,
    line: { color: col || LGREY, width: 0.5 } });
}

// ══════════════════════════════════════════════════════════
// SLIDE 1: 表紙
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  // 左縦ストライプ
  card(s, 0, 0, 0.08, 5.625, CORAL);
  // 上部コーラルバー
  card(s, 0.08, 0, 9.92, 0.9, CORAL);

  txt(s, '学生特化型 家電レンタル事業', 0.32, 0.15, 9, 0.3,
    { fontSize: 12, color: WHITE });
  txt(s, '事業計画書', 0.32, 0.95, 6.0, 0.72,
    { fontSize: 36, bold: true, color: DTXT });
  txt(s, '営業戦略プラン', 0.32, 1.62, 6.0, 0.56,
    { fontSize: 28, bold: true, color: CORAL });

  s.addShape(pres.ShapeType.line, { x: 0.32, y: 2.32, w: 3.5, h: 0,
    line: { color: LGREY, width: 1.0 } });

  txt(s, '株式会社サキュレ', 0.32, 2.48, 5, 0.38,
    { fontSize: 14, bold: true, color: DTXT });
  txt(s, '2026年7月　作成', 0.32, 2.9, 5, 0.3,
    { fontSize: 11, color: MTXT });

  // right side: key numbers
  const nums = [
    { val: '1,000万円', sub: 'KGI：1年以内独立採算' },
    { val: '400件', sub: '稼働目標（3点セット）' },
    { val: '13〜15校', sub: '大学生協提携目標' },
  ];
  nums.forEach((n, i) => {
    const y = 1.05 + i * 1.18;
    card(s, 6.8, y, 2.7, 1.0, OFFWH, LGREY);
    card(s, 6.8, y, 0.06, 1.0, CORAL);
    txt(s, n.val, 6.86, y + 0.08, 2.64, 0.52,
      { fontSize: 22, bold: true, color: CORAL, align: 'center' });
    txt(s, n.sub, 6.86, y + 0.6, 2.64, 0.34,
      { fontSize: 10, color: MTXT, align: 'center' });
  });

  // footer
  card(s, 0.08, 4.85, 9.92, 0.775, OFFWH);
  txt(s, 'KGI  ×  KPI  ×  行動量を一本化した実行計画', 0.08, 4.85, 9.92, 0.775,
    { fontSize: 12, bold: true, color: CORAL, align: 'center', valign: 'middle' });
}

// ══════════════════════════════════════════════════════════
// SLIDE 2: 目次
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '目次');

  const items = [
    { n: 1, text: 'KGI / KPI 全体像' },
    { n: 2, text: '外部環境分析（機会・脅威）' },
    { n: 3, text: '内部環境分析（強み・弱み）' },
    { n: 4, text: 'クロスSWOT → 戦略結論' },
    { n: 5, text: 'ユニットエコノミクス（リユース vs 新品）' },
    { n: 6, text: '営業戦略：チャンネル設計と生協アプローチ' },
    { n: 7, text: 'KPI連鎖・週次行動量' },
    { n: 8, text: '商品収集戦略' },
    { n: 9, text: '財務シミュレーション' },
    { n: 10, text: '7月アクションプラン' },
  ];

  items.forEach((item, i) => {
    const col = i < 5 ? 0.45 : 5.35;
    const row = i < 5 ? i : i - 5;
    const y = 0.9 + row * 0.88;
    card(s, col, y, 4.65, 0.72, WHITE, LGREY);
    card(s, col, y, 0.06, 0.72, CORAL);
    numCircle(s, item.n, col + 0.12, y + 0.12, CORAL);
    txt(s, item.text, col + 0.72, y + 0.14, 3.82, 0.44,
      { fontSize: 11, bold: true, color: DTXT, valign: 'middle' });
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 3: KGI / KPI 全体像
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, 'KGI / KPI 全体像', 1);

  // 3-phase milestone boxes
  const milestones = [
    { yr: '1年目', kpi: '稼働400件 / 売上1,000万', sub: '事業基盤確立・ユニット黒字', col: CORAL },
    { yr: '3年目', kpi: '単年度黒字　達成', sub: '稼働1,000件・固定費完全回収', col: GREEN },
    { yr: '5年目', kpi: '累損一掃　安定成長', sub: '稼働1,500件・継続投資フェーズへ', col: DKGRN },
  ];
  const mw = 2.96;
  milestones.forEach((m, i) => {
    const mx = 0.4 + i * 3.07;
    card(s, mx, 0.84, mw, 0.78, m.col);
    txt(s, m.yr, mx + 0.12, 0.88, mw - 0.24, 0.24, { fontSize: 10, bold: true, color: WHITE });
    txt(s, m.kpi, mx, 1.09, mw, 0.32,
      { fontSize: 13, bold: true, color: WHITE, align: 'center' });
    txt(s, m.sub, mx, 1.43, mw, 0.18, { fontSize: 9, color: WHITE, align: 'center' });
  });

  // Funnel levels
  const levels = [
    { label: '月次稼働台数', val: '400件', sub: '3点セット × 月2,083円 × 12ヶ月 ≒ 1,000万円', col: GREEN, w: 8.0 },
    { label: '大学生協 提携数', val: '13〜15校', sub: '1校あたり平均30件獲得（新入生300人×自宅外率30%×獲得率10%）', col: DKGRN, w: 7.0 },
    { label: '商談数', val: '40〜50回', sub: '提携締結率30〜40% → 13校到達', col: AMBER, w: 6.0 },
    { label: 'アポ獲得数', val: '80〜100件', sub: '商談化率50%', col: '9E4500', w: 5.0 },
    { label: '初回アプローチ数', val: '1,500〜2,000件', sub: 'アポ化率5%', col: CORAL, w: 4.2 },
  ];

  levels.forEach((lv, i) => {
    const x = (10 - lv.w) / 2;
    const y = 1.74 + i * 0.72;
    card(s, x, y, lv.w, 0.62, lv.col);
    txt(s, lv.label, x + 0.15, y + 0.04, 2.5, 0.28,
      { fontSize: 9, color: WHITE, charSpacing: 0.3 });
    txt(s, lv.val, x + 0.15, y + 0.3, 2.5, 0.28,
      { fontSize: 14, bold: true, color: WHITE });
    txt(s, lv.sub, x + 2.8, y + 0.14, lv.w - 3.0, 0.34,
      { fontSize: 10, color: WHITE });
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 4: 外部環境分析（OT）
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '外部環境分析（OT）', 2);

  slab(s, '機会（Opportunities）', 0.4, 0.84, 4.42, 0.32, GREEN);
  const opps = [
    '①毎年15万台の廃棄家電が発生。回収事業とのシナジーで調達コストがほぼゼロ',
    '②3〜4月の新入生需要が集中。在庫・物流を計画的に動かしやすい',
    '③競合未提携の都内中堅大学生協が多数。早期開拓で独占的チャンネルを確保できる',
    '④物価高・奨学金負担により初期費用削減ニーズが増加（保護者負担11万円/家電）',
    '⑤Z世代のリユース・SDGs意識が高く、リユース家電への抵抗感が低下中',
  ];
  opps.forEach((t, i) => {
    const y = 1.24 + i * 0.78;
    card(s, 0.4, y, 4.42, 0.68, LTGRN, LGREY);
    dot(s, 0.55, y + 0.28, GREEN);
    txt(s, t, 0.76, y + 0.08, 3.96, 0.52,
      { fontSize: 11, color: DTXT, lineSpacingMultiple: 1.3 });
  });

  slab(s, '脅威（Threats）', 5.18, 0.84, 4.42, 0.32, RED);
  const threats = [
    '①かして！どっとこむ・てぶらでどっとこむが生協提携で先行。後発の参入障壁',
    '②3〜4月の物流集中。引越し繁忙期と重なり配送・設置の人員不足リスク',
    '③「他人が使った家電」への心理的抵抗。品質保証・見た目の担保が必須',
    '④古物商許可・PSE・電安法など複数法令の遵守体制整備コストが重い',
  ];
  threats.forEach((t, i) => {
    const y = 1.24 + i * 0.78;
    card(s, 5.18, y, 4.42, 0.68, LTRED, LGREY);
    dot(s, 5.33, y + 0.28, RED);
    txt(s, t, 5.55, y + 0.08, 3.96, 0.52,
      { fontSize: 11, color: DTXT, lineSpacingMultiple: 1.3 });
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 5: 内部環境分析（SW）
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '内部環境分析（SW）', 3);

  slab(s, '強み（Strengths）', 0.4, 0.84, 4.42, 0.32, GREEN);
  const strengths = [
    { h: '仕入れコスト＝ほぼゼロ（リユース品）', b: '廃棄物回収網からリユース家電を調達（コストほぼゼロ）。ただし初期は台数にばらつきがあり、新品（調達コスト発生）で補う可能性がある。' },
    { h: '与野倉庫（既存インフラ）', b: '都心60分圏の倉庫確保済み。整備・保管・配送をワンストップ化。倉庫賃料の追加負担なし' },
    { h: '自社車両・ドライバーで配送回収', b: '廃棄物回収の既存車両を活用。外注なしで設置・回収が完結できる' },
    { h: '行政認定業者としての信頼性', b: '廃棄物処理許認可・実績あり。大学生協に「信頼できる法人」として提案できる' },
  ];
  strengths.forEach((st, i) => {
    const y = 1.24 + i * 0.98;
    card(s, 0.4, y, 4.42, 0.88, LTGRN, LGREY);
    card(s, 0.4, y, 0.06, 0.88, GREEN);
    txt(s, st.h, 0.55, y + 0.06, 4.18, 0.3,
      { fontSize: 12, bold: true, color: GREEN });
    txt(s, st.b, 0.55, y + 0.4, 4.18, 0.44,
      { fontSize: 10.5, color: MTXT, lineSpacingMultiple: 1.3 });
  });

  slab(s, '弱み（Weaknesses）', 5.18, 0.84, 4.42, 0.32, MTXT);
  const weaknesses = [
    { h: '認知ゼロ・ブランドなし', b: '学生・保護者に知られていない。生協提案で「聞いたことがない」と言われるリスク' },
    { h: '学生向け営業・マーケ未経験', b: '廃棄物回収はBtoB営業。大学生協交渉・SNS集客は新規領域' },
    { h: '品質ばらつきと整備基準未整備', b: '回収品はコンディションに差がある。\nチェックシートで品質基準を統一する必要がある' },
    { h: '繁忙期の物流キャパ制約', b: '3〜4月に400件を納品すると廃棄物回収業務と競合する可能性がある' },
  ];
  weaknesses.forEach((wk, i) => {
    const y = 1.24 + i * 0.98;
    card(s, 5.18, y, 4.42, 0.88, 'F5F5F5', LGREY);
    card(s, 5.18, y, 0.06, 0.88, MTXT);
    txt(s, wk.h, 5.33, y + 0.06, 4.18, 0.3,
      { fontSize: 12, bold: true, color: DTXT });
    txt(s, wk.b, 5.33, y + 0.4, 4.18, 0.44,
      { fontSize: 10.5, color: MTXT, lineSpacingMultiple: 1.3 });
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 6: クロスSWOT → 戦略結論
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, 'クロスSWOT → 戦略結論', 4);

  const cells = [
    { x: 0.4,  y: 1.04, w: 4.3, h: 1.82, bg: LTGRN, hdr: 'SO戦略（強み × 機会）', hcol: GREEN,
      body: '仕入れ0円＋既存物流で業界最安値を実現。\n競合未提携の生協を先取りし学生を囲い込む。\n→ 低価格学生囲い込み戦略' },
    { x: 5.3,  y: 1.04, w: 4.3, h: 1.82, bg: LTAM,  hdr: 'WO戦略（弱み × 機会）', hcol: AMBER,
      body: 'ブランドなし → 生協の看板を借りる提案で信頼を補完。\nSNSは2年目以降。初年度は生協BtoBに集中。' },
    { x: 0.4,  y: 3.04, w: 4.3, h: 1.88, bg: LTBL,  hdr: 'ST戦略（強み × 脅威）', hcol: CORAL,
      body: '品質認知リスク → 整備済み証明・チェックシートで可視化。\n繁忙期集中リスク → 台数上限を設定し先着予約制で分散。' },
    { x: 5.3,  y: 3.04, w: 4.3, h: 1.88, bg: LTRED, hdr: 'WT戦略（弱み × 脅威）', hcol: RED,
      body: '法令対応（古物商・PSE）は提携法律事務所・行政書士と早期整備。\n競合後発を品質管理の明文化でブランド化する。' },
  ];
  cells.forEach(c => {
    card(s, c.x, c.y, c.w, c.h, c.bg, LGREY);
    slab(s, c.hdr, c.x, c.y, c.w, 0.28, c.hcol);
    txt(s, c.body, c.x + 0.15, c.y + 0.36, c.w - 0.3, c.h - 0.46,
      { fontSize: 11.5, color: DTXT, lineSpacingMultiple: 1.45 });
  });

  // Strategy conclusion box
  card(s, 2.5, 4.96, 5.0, 0.5, CORAL);
  txt(s, '「低価格学生囲い込み戦略」　業界最安値 × 生協BtoBで初年度400件を獲得する', 2.5, 4.96, 5.0, 0.5,
    { fontSize: 11.5, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}

// ══════════════════════════════════════════════════════════
// SLIDE 7: ユニットエコノミクス
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, 'ユニットエコノミクス（3点セット・2年サイクルあたりの収支）', 5);

  const cols = [
    { label: '項目', x: 0.4,  w: 2.2 },
    { label: 'リユース\n（回収品）', x: 2.65, w: 2.28, col: GREEN },
    { label: '新品\n（回収品）', x: 5.0,  w: 2.28, col: CORAL },
    { label: '新品\n（市場購入）', x: 7.35, w: 2.25, col: RED },
  ];
  cols.forEach(c => {
    if (c.col) {
      slab(s, c.label, c.x, 0.84, c.w, 0.52, c.col);
    } else {
      card(s, c.x, 0.84, c.w, 0.52, DTXT);
      txt(s, c.label, c.x, 0.84, c.w, 0.52,
        { fontSize: 11, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    }
  });

  const rows = [
    { label: '調達コスト', v1: '0円', v2: '0円', v3: '▲50,000円', neg3: true },
    { label: '整備・クリーニング', v1: '▲5,000円', v2: '▲1,000円', v3: '0円', neg1: true, neg2: true },
    { label: '配達・設置', v1: '▲10,000円', v2: '▲10,000円', v3: '▲10,000円', neg1:true, neg2:true, neg3:true },
    { label: '回収（2年後）', v1: '▲10,000円', v2: '▲10,000円', v3: '▲10,000円', neg1:true, neg2:true, neg3:true },
    { label: 'レンタル収入（2年間）', v1: '+50,000円', v2: '+50,000円', v3: '+50,000円', pos1:true, pos2:true, pos3:true },
  ];
  rows.forEach((r, i) => {
    const y = 1.44 + i * 0.54;
    const bg = i % 2 === 0 ? OFFWH : WHITE;
    card(s, 0.4, y, 2.2, 0.5, bg);
    txt(s, r.label, 0.52, y, 2.08, 0.5,
      { fontSize: 11, color: DTXT, valign: 'middle' });
    [[2.65, r.v1, r.neg1, r.pos1], [5.0, r.v2, r.neg2, r.pos2], [7.35, r.v3, r.neg3, r.pos3]].forEach(([x, v, neg, pos]) => {
      card(s, x, y, 2.28, 0.5, bg);
      txt(s, v, x, y, 2.28, 0.5,
        { fontSize: 12, bold: (neg || pos), color: neg ? RED : (pos ? POSGR : DTXT), align: 'center', valign: 'middle' });
    });
    hrule(s, 0.4, y + 0.5, 9.2);
  });

  const ry = 1.44 + rows.length * 0.54;
  card(s, 0.4, ry, 2.2, 0.72, CORAL);
  txt(s, '初回サイクル利益', 0.52, ry, 2.08, 0.72,
    { fontSize: 12, bold: true, color: WHITE, valign: 'middle' });
  [
    [2.65, '+25,000円', '（50%）', LTGRN, POSGR, GREEN],
    [5.0,  '+29,000円', '（58%）', LTBL,  POSGR, CORAL],
    [7.35, '▲20,000円', '（−40%）', LTRED, RED,   RED],
  ].forEach(([x, v, pct, bg, vc, sc]) => {
    card(s, x, ry, 2.28, 0.72, bg, LGREY);
    txt(s, v, x, ry + 0.04, 2.28, 0.4,
      { fontSize: 16, bold: true, color: vc, align: 'center' });
    txt(s, pct, x, ry + 0.42, 2.28, 0.24,
      { fontSize: 11, color: sc, align: 'center' });
  });

  txt(s, '※ 新品（市場購入）は2回転目以降（3〜4年目）から利益化。3回転で累計+30,000円。  ※ 配達・設置・回収の人件費・車両費は10,000円/回で試算。', 0.4, ry + 0.78, 9.2, 0.32,
    { fontSize: 9, color: MTXT });
}

// ══════════════════════════════════════════════════════════
// SLIDE 8: 営業戦略
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '営業戦略：チャンネル設計', 6);

  const channels = [
    { n: 1, pri: '最重点', col: GREEN, bg: LTGRN,
      title: '小規模生協への直接訪問と資料郵送',
      why: '競合未提携の小規模校が多く、留学生ニーズが高い。まず訪問または紙資料で接触',
      how: ['留学生比率大・小規模生協をリストアップ',
            '直接訪問または紙ベース資料を郵送して提案',
            '在日留学生コミュニティへの口コミも活用'],
      timing: '10〜12月仕込み、3〜4月獲得' },
    { n: 2, pri: '補完', col: AMBER, bg: LTAM,
      title: 'チラシ設置 + KB 2,000円/件',
      why: '入居タイミングで接触。KB設計で紹介元の動機を確保',
      how: ['仲介・入寮施設にチラシ設置（KB 2,000円/件）',
            '外国人支援窓口・留学生ハウスにも配布',
            '管理会社にQRコード付きカードを設置'],
      timing: '通年（1〜3月を重点化）' },
    { n: 3, pri: '低優先', col: MTXT, bg: OFFWH,
      title: '既存顧客経由（廃棄物回収先250社）',
      why: '既存250社との信頼関係を活用。受注コストほぼゼロだが優先度は3番',
      how: ['既存法人担当者へDM・TELで紹介依頼',
            '担当者のお子様・知人の新入生にアプローチ',
            '成約時に紹介者へ謝礼（商品券・割引等）'],
      timing: '随時（優先度3番）' },
  ];

  channels.forEach((ch, i) => {
    const y = 0.86 + i * 1.54;
    card(s, 0.4, y, 9.2, 1.38, ch.bg, LGREY);
    card(s, 0.4, y, 0.06, 1.38, ch.col);
    numCircle(s, ch.n, 0.55, y + 0.43, ch.col);
    slab(s, ch.pri, 1.14, y + 0.08, 0.78, 0.28, ch.col);
    txt(s, ch.title, 2.02, y + 0.06, 4.0, 0.36,
      { fontSize: 14, bold: true, color: DTXT });
    txt(s, ch.why, 2.02, y + 0.44, 4.0, 0.32,
      { fontSize: 10.5, color: MTXT, lineSpacingMultiple: 1.3 });
    ch.how.forEach((h, j) => {
      dot(s, 2.02, y + 0.82 + j * 0.26 + 0.02, ch.col);
      txt(s, h, 2.22, y + 0.79 + j * 0.26, 3.82, 0.25,
        { fontSize: 10, color: DTXT });
    });
    card(s, 6.18, y + 0.2, 3.22, 0.38, ch.col);
    txt(s, ch.timing, 6.18, y + 0.2, 3.22, 0.38,
      { fontSize: 10, color: WHITE, align: 'center', valign: 'middle' });
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 9 (追加): 営業仮戦略候補
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '営業仮戦略候補（追加検討）');

  // Sub-header
  card(s, 0.4, 0.76, 9.2, 0.36, OFFWH, LGREY);
  txt(s, '以下は現在検討中の追加チャンネル案です。優先度・実現性を確認しながら順次着手する想定です。',
    0.55, 0.76, 8.9, 0.36,
    { fontSize: 10, color: MTXT, valign: 'middle' });

  const extras = [
    { n: 4, title: 'リスト化＋紙DM郵送',
      body: '新入生・在学留学生のリストを整備し、案内状を郵送。大学から入手可能な情報を活用',
      timing: '1〜3月重点' },
    { n: 5, title: 'SNS広告（Instagram・LINE）',
      body: '学生層へターゲティング広告。入学シーズンに集中出稿してリーチを確保',
      timing: '2〜4月重点' },
    { n: 6, title: '大学イベントでのビラ配り',
      body: '入学式・新歓・オープンキャンパスで直接配布。QRコードで申込につなげる',
      timing: '3〜4月・9月' },
    { n: 7, title: '留学生情報サイト掲載',
      body: 'JASSO・留学生センター等の情報サイトへの掲載。在日留学生への認知獲得',
      timing: '随時' },
    { n: 8, title: '不動産仲介・管理会社提携',
      body: '新入生の入居時に紹介。KB 2,000円/件で動機づけ。既存訪問先250社と重複可',
      timing: '1〜3月重点' },
    { n: 9, title: '国際交流協会・市区役所',
      body: '外国人支援窓口・国際交流協会経由で留学生に案内。公的信頼性も付加できる',
      timing: '随時' },
  ];

  extras.forEach((ex, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.92;
    const y = 1.22 + row * 1.44;
    card(s, x, y, 4.68, 1.30, OFFWH, LGREY);
    card(s, x, y, 0.06, 1.30, CORAL);
    numCircle(s, ex.n, x + 0.12, y + 0.40, CORAL);
    // Row 1: badges
    slab(s, '仮戦略', x + 0.72, y + 0.06, 0.72, 0.24, MUTED);
    card(s, x + 1.50, y + 0.06, 1.0, 0.24, LGREY);
    txt(s, ex.timing, x + 1.50, y + 0.06, 1.0, 0.24,
      { fontSize: 8.5, color: DTXT, align: 'center', valign: 'middle' });
    // Row 2: title (full width, allows wrap)
    txt(s, ex.title, x + 0.72, y + 0.34, 3.84, 0.36,
      { fontSize: 11.5, bold: true, color: DTXT });
    // Row 3: body
    txt(s, ex.body, x + 0.72, y + 0.74, 3.84, 0.50,
      { fontSize: 9.5, color: MTXT, lineSpacingMultiple: 1.3 });
  });

  s.addNotes('追加営業仮戦略候補6件。リスト化紙DM・SNS広告・大学イベントビラ・留学生サイト掲載・不動産仲介提携・国際交流協会経由。優先度と実現性を検討しながら順次実施。');
}

// ══════════════════════════════════════════════════════════
// SLIDE 10 (旧9): 生協営業フロー（4フェーズ）
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, OFFWH);
  pgTitle(s, '生協営業フロー（4フェーズ）');

  const phases = [
    { n: 1, period: '10〜11月', title: 'アポ獲得',
      actions: ['生協・大学事務へのTEL・メール', 'ターゲットリスト50〜80校', 'アポ目標：5〜8件/月'],
      kpi: 'アポ80〜100件' },
    { n: 2, period: '11〜12月', title: '提案・交渉',
      actions: ['料金比較表・整備証明を提示', 'レベニューシェア設計を提案', '目標提携締結率：30〜40%'],
      kpi: '提携13〜15校' },
    { n: 3, period: '1〜2月', title: '告知・予約受付',
      actions: ['生協HPと掲示板に掲載', '合格発表日（2月）に集中告知', '先着制・台数上限で締め切り'],
      kpi: '予約200件以上' },
    { n: 4, period: '3〜4月', title: '配送・設置',
      actions: ['2週間に集中して完了させる', 'スケジュールを事前確定', '丁寧な設置で口コミ獲得'],
      kpi: '稼働400件達成' },
  ];

  phases.forEach((ph, i) => {
    const x = 0.3 + i * 2.38;
    card(s, x, 0.86, 2.18, 4.5, WHITE, LGREY);
    card(s, x, 0.86, 2.18, 0.06, CORAL);
    numCircle(s, ph.n, x + 0.82, 0.96, CORAL);
    txt(s, ph.period, x, 1.56, 2.18, 0.3,
      { fontSize: 10, color: MTXT, align: 'center' });
    txt(s, ph.title, x, 1.88, 2.18, 0.42,
      { fontSize: 15, bold: true, color: CORAL, align: 'center' });
    hrule(s, x + 0.15, 2.36, 1.88, CORAL);
    ph.actions.forEach((a, j) => {
      dot(s, x + 0.2, 2.51 + j * 0.44 + 0.02, CORAL);
      txt(s, a, x + 0.38, 2.48 + j * 0.44, 1.72, 0.4,
        { fontSize: 10.5, color: DTXT, lineSpacingMultiple: 1.25 });
    });
    card(s, x + 0.12, 4.64, 1.94, 0.48, CORAL);
    txt(s, ph.kpi, x + 0.12, 4.64, 1.94, 0.48,
      { fontSize: 11, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    if (i < 3) {
      txt(s, '▶', x + 2.18, 2.86, 0.2, 0.4,
        { fontSize: 14, color: CORAL, align: 'center' });
    }
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 10: KPI連鎖・週次行動量
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, 'KPI連鎖・週次行動量', 7);

  const funnel = [
    { label: 'KGI', val: '売上1,000万円/年', col: CORAL, w: 4.5 },
    { label: '稼働数', val: '400件（月次）', col: GREEN, w: 4.1 },
    { label: '提携校数', val: '13〜15校', col: DKGRN, w: 3.6 },
    { label: '商談数', val: '40〜50回', col: AMBER, w: 3.1 },
    { label: 'アポ数', val: '80〜100件', col: '9E4500', w: 2.6 },
    { label: '接触数', val: '1,500〜2,000件', col: DTXT, w: 2.1 },
  ];
  funnel.forEach((f, i) => {
    const bx = 0.3 + (4.5 - f.w) / 2;
    const y = 0.86 + i * 0.74;
    card(s, bx, y, f.w, 0.62, f.col);
    txt(s, f.label, bx + 0.12, y + 0.04, 1.0, 0.28,
      { fontSize: 9, color: WHITE });
    txt(s, f.val, bx + 0.12, y + 0.3, f.w - 0.24, 0.28,
      { fontSize: 13, bold: true, color: WHITE });
  });

  slab(s, '週次管理指標（行動量の担保）', 5.0, 0.86, 4.6, 0.32, CORAL);
  const weeks = [
    { metric: 'アプローチ数（TEL・メール）', target: '30〜40件/週' },
    { metric: '提案面談数', target: '2〜3件/週' },
    { metric: '提携締結累計（12月末）', target: '5〜7校' },
    { metric: '予約件数（2月末）', target: '200件以上' },
    { metric: '稼働開始件数（4月末）', target: '400件' },
  ];
  weeks.forEach((w, i) => {
    const y = 1.26 + i * 0.68;
    card(s, 5.0, y, 4.6, 0.6, i % 2 === 0 ? OFFWH : WHITE, LGREY);
    txt(s, w.metric, 5.12, y + 0.1, 2.98, 0.4,
      { fontSize: 11.5, color: DTXT, valign: 'middle' });
    card(s, 8.22, y + 0.1, 1.28, 0.4, CORAL);
    txt(s, w.target, 8.22, y + 0.1, 1.28, 0.4,
      { fontSize: 11, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  });

  txt(s, '週次でアプローチ数・商談数・締結数をレビューし、行動量の下振れを早期検知する', 5.0, 4.68, 4.6, 0.32,
    { fontSize: 9.5, color: MTXT });
}

// ══════════════════════════════════════════════════════════
// SLIDE 11: 商品収集戦略
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '商品収集戦略', 8);

  card(s, 0.4, 0.84, 9.2, 0.52, CORAL);
  txt(s, '初年度目標：400セット確保（冷蔵庫・洗濯機・電子レンジ 各400台）', 0.4, 0.84, 9.2, 0.52,
    { fontSize: 13, bold: true, color: WHITE, align: 'center', valign: 'middle' });

  const routes = [
    { n: 1, col: GREEN,  bg: LTGRN, title: 'サキュレ廃棄回収品の選別',
      vol: '200〜300台/年', cost: 'ほぼ0円',
      body: '既存の廃棄物回収ルートから家電を選別。調達コストが発生しない最優先ルート。選別率50〜60%を目標。' },
    { n: 2, col: CORAL,  bg: OFFWH, title: '不動産・管理会社の退去残置物',
      vol: '50〜100台', cost: '引取料のみ',
      body: '賃貸退去・原状回復時に発生する残置家電を引き取り。管理会社へ営業することで安定供給を確保。' },
    { n: 3, col: AMBER, bg: LTAM, title: 'ジモティー・業者オークション',
      vol: '100台', cost: '3〜5千円/台',
      body: 'サキュレルートで不足する台数を補完。Aucnet等の業者オークションで状態の良いものを選んで仕入れ。' },
    { n: 4, col: MTXT, bg: 'F5F5F5', title: '大学・学生寮の入れ替え品',
      vol: '30〜50台', cost: '引取料のみ',
      body: '大学の設備更新や学生寮の退去シーズン（3〜4月）に発生する廃棄予定品を先取りする。' },
  ];

  routes.forEach((r, i) => {
    const x = i < 2 ? 0.4 : 5.15;
    const y = i < 2 ? 1.48 + i * 1.72 : 1.48 + (i - 2) * 1.72;
    card(s, x, y, 4.55, 1.56, r.bg, LGREY);
    card(s, x, y, 0.06, 1.56, r.col);
    numCircle(s, r.n, x + 0.12, y + 0.52, r.col);
    txt(s, r.title, x + 0.72, y + 0.1, 3.7, 0.38,
      { fontSize: 13, bold: true, color: DTXT });
    card(s, x + 0.72, y + 0.52, 0.96, 0.28, r.col);
    txt(s, r.vol, x + 0.72, y + 0.52, 0.96, 0.28,
      { fontSize: 9, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    card(s, x + 1.74, y + 0.52, 1.2, 0.28, 'E0E0E0');
    txt(s, r.cost, x + 1.74, y + 0.52, 1.2, 0.28,
      { fontSize: 9, color: DTXT, align: 'center', valign: 'middle' });
    txt(s, r.body, x + 0.72, y + 0.88, 3.7, 0.58,
      { fontSize: 10, color: MTXT, lineSpacingMultiple: 1.3 });
  });

  card(s, 0.4, 4.9, 9.2, 0.5, LTGRN, LGREY);
  txt(s, '整備基準：動作確認全項目・クリーニング・製造から10年以内・整備済みシール貼付・チェックシート保管', 0.55, 4.9, 9.0, 0.5,
    { fontSize: 11, color: GREEN, valign: 'middle' });
}

// ══════════════════════════════════════════════════════════
// SLIDE 12: 財務シミュレーション（5カ年）
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '財務シミュレーション（5カ年）', 9);

  slab(s, '5カ年 損益シミュレーション（リユース主力・KB込み）', 0.3, 0.82, 9.4, 0.3, GREEN);

  const LX = 0.3, LW = 2.38;
  const YX = [2.72, 4.12, 5.52, 6.92, 8.32];
  const YW = 1.34;
  const RH = 0.50;
  const SPECIAL = [2, 4]; // Y3, Y5

  // Column headers
  card(s, LX, 1.16, LW, 0.5, DTXT);
  txt(s, '項目', LX, 1.16, LW, 0.5,
    { fontSize: 10, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  const yLabels = ['1年目', '2年目', '3年目', '4年目', '5年目'];
  const yBadges = ['', '', '★黒字化', '', '★累損一掃'];
  YX.forEach((x, i) => {
    const sp = SPECIAL.includes(i);
    card(s, x, 1.16, YW, 0.5, sp ? GREEN : DTXT);
    if (sp) {
      txt(s, yLabels[i], x, 1.18, YW, 0.24,
        { fontSize: 10, bold: true, color: WHITE, align: 'center' });
      txt(s, yBadges[i], x, 1.40, YW, 0.22,
        { fontSize: 9, bold: true, color: WHITE, align: 'center' });
    } else {
      txt(s, yLabels[i], x, 1.16, YW, 0.5,
        { fontSize: 10, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    }
  });

  // Data rows
  const plRows = [
    { label: '稼働件数',        vals: ['400件',  '600件',  '1,000件', '1,200件', '1,500件'], kind: 'n' },
    { label: '年間売上',        vals: ['1,000万', '1,500万', '2,500万', '3,000万', '3,750万'], kind: 'pos' },
    { label: '粗利（40%）',     vals: ['400万',  '600万',  '1,000万', '1,200万', '1,500万'], kind: 'pos' },
    { label: '固定費※',        vals: ['▲750万', '▲560万', '▲680万', '▲720万', '▲780万'], kind: 'neg' },
    { label: 'KB（2千円/件）', vals: ['▲80万',  '▲120万', '▲200万', '▲240万', '▲300万'], kind: 'neg' },
    { label: '営業利益（単年）', vals: ['▲430万', '▲80万',  '+120万', '+240万', '+420万'], kind: 'profit' },
    { label: '累積損益',        vals: ['▲430万', '▲510万', '▲390万', '▲150万', '+270万'], kind: 'cumul' },
  ];

  plRows.forEach((r, ri) => {
    const y = 1.70 + ri * RH;
    const isProfit = r.kind === 'profit';
    const isCumul  = r.kind === 'cumul';
    const rowBg = isCumul ? CORAL : (isProfit ? LTGRN : (ri % 2 === 0 ? OFFWH : WHITE));
    card(s, LX, y, LW, RH, rowBg);
    if (isProfit) card(s, LX, y, 0.06, RH, GREEN);
    if (isCumul)  card(s, LX, y, 0.06, RH, DKGRN);
    txt(s, r.label, LX + 0.12, y, LW - 0.14, RH,
      { fontSize: 10.5, bold: isProfit || isCumul,
        color: isCumul ? WHITE : DTXT, valign: 'middle' });

    r.vals.forEach((v, vi) => {
      const x = YX[vi];
      const sp = SPECIAL.includes(vi);
      const isNeg = v.startsWith('▲');
      const isPos = v.startsWith('+');
      const cellBg = isCumul
        ? (isNeg ? LTRED : LTGRN)
        : (isProfit ? (isNeg ? LTRED : LTGRN) : rowBg);
      card(s, x, y, YW, RH, cellBg);
      const valColor = (isProfit || isCumul)
        ? (isNeg ? RED : POSGR)
        : (r.kind === 'neg' ? RED : (r.kind === 'pos' ? POSGR : DTXT));
      txt(s, v, x, y, YW, RH,
        { fontSize: isProfit || isCumul ? 12 : 11,
          bold: isProfit || isCumul || sp,
          color: valColor, align: 'center', valign: 'middle' });
    });
    hrule(s, LX, y + RH, LW + YX.length * YW + 0.05);
  });

  txt(s, '※1年目固定費は初期費用（設備・HP・営業ツール等）300万円を含む　KB=キックバック（2,000円/成約件）',
    0.3, 5.26, 9.4, 0.24, { fontSize: 8.5, color: MUTED });
}

// ══════════════════════════════════════════════════════════
// SLIDE 13: 7月アクションプラン
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, OFFWH);
  pgTitle(s, '7月アクションプラン（今すぐやること）', 10);

  const actions = [
    { cat: '営業準備', col: GREEN, items: [
      'ターゲット生協リスト作成（50〜80校：大学名・担当部署・TEL）',
      '提案資料の作成（料金比較表・整備証明・配送SLA一覧）',
      'レベニューシェア条件の設計（成約2,000円/件）',
    ]},
    { cat: '商品・整備', col: CORAL, items: [
      'サキュレ回収品から使用可能な家電台数の棚卸し（目標：200〜300台）',
      '整備チェックシートの作成（動作確認・クリーニング・年式確認）',
      '整備済みシールのデザイン・発注',
    ]},
    { cat: '法令・管理', col: AMBER, items: [
      '古物商許可の取得状況確認（未取得の場合は申請着手）',
      'PSE・電安法対応チェック（対象品目の確認）',
      'レンタル契約書のひな形作成（弁護士・行政書士と確認）',
    ]},
    { cat: '財務', col: RED, items: [
      '初期投資額の試算（整備費・消耗品・営業費）',
      'キャッシュフロー表の作成（仕入れ先行 → 3月回収の流れ）',
    ]},
  ];

  actions.forEach((ac, i) => {
    const x = i < 2 ? 0.4 : 5.15;
    const y = i < 2 ? 0.84 + i * 2.35 : 0.84 + (i - 2) * 2.35;
    const h = ac.items.length === 2 ? 2.08 : 2.28;
    card(s, x, y, 4.55, h, WHITE, LGREY);
    card(s, x, y, 4.55, 0.32, ac.col);
    slab(s, ac.cat, x, y, 4.55, 0.32, ac.col);
    ac.items.forEach((item, j) => {
      const iy = y + 0.42 + j * 0.6;
      card(s, x + 0.15, iy, 0.28, 0.28, ac.col);
      txt(s, '✓', x + 0.15, iy, 0.28, 0.28,
        { fontSize: 11, bold: true, color: WHITE, align: 'center', valign: 'middle' });
      txt(s, item, x + 0.52, iy - 0.02, 3.92, 0.4,
        { fontSize: 11, color: DTXT, lineSpacingMultiple: 1.3 });
    });
  });

  card(s, 0.4, 5.18, 9.2, 0.3, CORAL);
  txt(s, '7月末時点でのチェック：リスト作成完了・棚卸し完了・提案資料初稿完成', 0.4, 5.18, 9.2, 0.3,
    { fontSize: 11, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}

// ─── Output ───────────────────────────────────────────────
const OUT = '/tmp/claude-0/-home-user-sacure/1db07409-6d1e-5d98-92a7-465fca194ddc/scratchpad/事業計画書_営業戦略プラン_v2.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('Done:', OUT));
