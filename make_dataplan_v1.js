const pptx = require('pptxgenjs');
const pres = new pptx();
pres.layout = 'LAYOUT_16x9'; // 10" × 5.625"

// ─── Palette（元FMT準拠） ──────────────────────────────────
const CORAL = 'DD7E6B';
const DKCRL = 'AA5848';
const GREEN = '156818';
const DKGRN = '115000';
const LTGRN = 'D9EAD3';
const LTRED = 'F4CCCC';
const RED   = 'DD001B';
const AMBER = 'E65100';
const LTAM  = 'FFF3E0';
const LTBL  = 'C9DAF8';
const NAVY  = '1E3461';
const TEAL  = '0D7480';
const LTTL  = 'D0EFEF';
const WHITE = 'FFFFFF';
const OFFWH = 'F3F3F3';
const DTXT  = '333333';
const MTXT  = '595959';
const MUTED = '999999';
const LGREY = 'CCCCCC';
const POSGR = '188038';

// ─── Helpers ──────────────────────────────────────────────
function card(s, x, y, w, h, fill, lineColor) {
  s.addShape(pres.ShapeType.rect, { x, y, w, h,
    fill: { color: fill || WHITE },
    line: lineColor ? { color: lineColor, width: 0.75 } : { type: 'none' } });
}
function txt(s, text, x, y, w, h, opts) {
  s.addText(text, { x, y, w, h, fontFace: 'Calibri', ...opts });
}
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
  card(s, 0, 0, 0.08, 5.625, CORAL);
  card(s, 0.08, 0, 9.92, 0.9, CORAL);

  txt(s, '廃棄物量データ事業　×　都市センシングデータ事業', 0.32, 0.15, 9, 0.3,
    { fontSize: 12, color: WHITE });
  txt(s, 'データ活用事業', 0.32, 0.96, 7.0, 0.72,
    { fontSize: 36, bold: true, color: DTXT });
  txt(s, '事業計画書', 0.32, 1.64, 6.0, 0.56,
    { fontSize: 28, bold: true, color: CORAL });

  s.addShape(pres.ShapeType.line, { x: 0.32, y: 2.34, w: 3.5, h: 0,
    line: { color: LGREY, width: 1.0 } });

  txt(s, '株式会社サキュレ', 0.32, 2.50, 5, 0.38,
    { fontSize: 14, bold: true, color: DTXT });
  txt(s, '2026年7月　作成', 0.32, 2.92, 5, 0.3,
    { fontSize: 11, color: MTXT });

  const nums = [
    { val: '200台以上', sub: '廃棄物収集トラック（データ収集基盤）' },
    { val: '2事業', sub: '廃棄物データ × 走行センサーデータ' },
    { val: '3,000万円', sub: 'KGI：3年目年間売上目標' },
  ];
  nums.forEach((n, i) => {
    const y = 1.05 + i * 1.18;
    card(s, 6.8, y, 2.7, 1.0, OFFWH, LGREY);
    card(s, 6.8, y, 0.06, 1.0, CORAL);
    txt(s, n.val, 6.86, y + 0.08, 2.64, 0.52,
      { fontSize: 22, bold: true, color: CORAL, align: 'center' });
    txt(s, n.sub, 6.86, y + 0.6, 2.64, 0.34,
      { fontSize: 9.5, color: MTXT, align: 'center' });
  });

  card(s, 0.08, 4.85, 9.92, 0.775, OFFWH);
  txt(s, '廃棄物収集の副産物から「データ」という第二の収益柱を構築する', 0.08, 4.85, 9.92, 0.775,
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
    { n: 1, text: '事業コンセプト（なぜ今・なぜサキュレか）' },
    { n: 2, text: '事業①　廃棄物量データビジネス' },
    { n: 3, text: '事業②　トラック走行・都市センシングデータ' },
    { n: 4, text: '競合優位性・参入障壁' },
    { n: 5, text: '収益モデル・料金設計' },
    { n: 6, text: 'KGI / KPI・マイルストーン' },
    { n: 7, text: 'ロードマップ（3フェーズ）' },
    { n: 8, text: '財務シミュレーション（3カ年）' },
    { n: 9, text: '次のアクション（7〜9月）' },
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
// SLIDE 3: 事業コンセプト
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '事業コンセプト　なぜ今・なぜサキュレか', 1);

  // Central message
  card(s, 0.4, 0.82, 9.2, 0.56, CORAL);
  txt(s, '200台のトラックが毎日走るだけで、価値あるデータが生まれる', 0.4, 0.82, 9.2, 0.56,
    { fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle' });

  // Left: waste data
  card(s, 0.4, 1.5, 4.42, 2.9, LTGRN, LGREY);
  card(s, 0.4, 1.5, 4.42, 0.36, GREEN);
  txt(s, '事業①　廃棄物量データ', 0.4, 1.5, 4.42, 0.36,
    { fontSize: 12, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  const wastePoints = [
    '回収先250施設の廃棄物量を毎月記録・蓄積',
    '施設の稼働状況・消費トレンドの代替指標',
    'ヘッジファンド／調査会社／ESG部門が需要',
    'ラクハイで蓄積済み → 追加コストほぼゼロ',
  ];
  wastePoints.forEach((p, i) => {
    dot(s, 0.6, 1.99 + i * 0.56 + 0.02, GREEN);
    txt(s, p, 0.82, 1.96 + i * 0.56, 3.88, 0.5,
      { fontSize: 11.5, color: DTXT, lineSpacingMultiple: 1.3 });
  });

  // Right: truck sensor data
  card(s, 5.18, 1.5, 4.42, 2.9, LTTL, LGREY);
  card(s, 5.18, 1.5, 4.42, 0.36, TEAL);
  txt(s, '事業②　走行センサーデータ', 5.18, 1.5, 4.42, 0.36,
    { fontSize: 12, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  const sensorPoints = [
    'カメラ＋エッジAIをトラックに搭載',
    '街路樹の損傷・落下物・道路損傷を自動検知',
    '自治体・インフラ管理会社・保険会社が需要',
    '走行中に自動収集 → 現場調査コストが不要',
  ];
  sensorPoints.forEach((p, i) => {
    dot(s, 5.38, 1.99 + i * 0.56 + 0.02, TEAL);
    txt(s, p, 5.6, 1.96 + i * 0.56, 3.88, 0.5,
      { fontSize: 11.5, color: DTXT, lineSpacingMultiple: 1.3 });
  });

  // Bottom insight
  card(s, 0.4, 4.55, 9.2, 0.82, OFFWH, LGREY);
  card(s, 0.4, 4.55, 0.06, 0.82, CORAL);
  txt(s, '廃棄物収集は「許可が必要な参入障壁の高い事業」。\n競合他社には施設立入権限・収集許可・行政との信頼関係がない。\nこのデータは、サキュレにしか集められない。', 0.56, 4.58, 8.9, 0.76,
    { fontSize: 11.5, bold: true, color: DTXT, lineSpacingMultiple: 1.5 });
}

// ══════════════════════════════════════════════════════════
// SLIDE 4: 事業①　廃棄物データビジネス
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '事業①　廃棄物量データビジネス', 2);

  // Left: market & data
  card(s, 0.4, 0.82, 4.42, 1.18, LTGRN, LGREY);
  slab(s, 'なぜ廃棄物データに価値があるのか', 0.4, 0.82, 4.42, 0.3, GREEN);
  const reasons = [
    '廃棄物量は「見えない経済指標」として機能',
    '空港ゴミ量→来場者数、ホテルゴミ→稼働率など',
    '日本のESGデータ開示義務が2030年方向で拡大中',
  ];
  reasons.forEach((r, i) => {
    dot(s, 0.56, 1.22 + i * 0.26 + 0.02, GREEN);
    txt(s, r, 0.78, 1.19 + i * 0.26, 3.9, 0.24,
      { fontSize: 10.5, color: DTXT });
  });

  card(s, 0.4, 2.1, 4.42, 1.28, LTGRN, LGREY);
  slab(s, 'サキュレが持つデータ', 0.4, 2.1, 4.42, 0.3, DKGRN);
  const dataItems = [
    '回収先250施設 × 月次収集量（品目別）',
    '過去累積データ（長期トレンド分析に活用可）',
    '曜日・季節・天候との相関分析が可能',
    '施設種別（飲食・医療・工場・商業）で分類済み',
  ];
  dataItems.forEach((d, i) => {
    dot(s, 0.56, 2.5 + i * 0.24 + 0.02, DKGRN);
    txt(s, d, 0.78, 2.47 + i * 0.24, 3.9, 0.22,
      { fontSize: 10, color: DTXT });
  });

  // Right: customers table
  slab(s, '顧客セグメント × 活用用途 × 想定単価', 5.02, 0.82, 4.58, 0.3, CORAL);

  const customers = [
    { seg: 'ヘッジファンド',    use: '景気先行指標として投資判断に活用',      price: '月50〜100万円', bg: LTGRN },
    { seg: '調査会社',          use: '消費・産業動向レポートへ組み込み販売',  price: '月10〜30万円',  bg: WHITE },
    { seg: '大企業ESG部門',     use: '廃棄物KPI開示データとして購入',         price: '月5〜15万円',   bg: LTGRN },
    { seg: '自治体・行政',      use: '廃棄物量予測・政策立案・資源配分最適化', price: '月5〜15万円',   bg: WHITE },
    { seg: '不動産・REIT',      use: '施設稼働率の代替指標として評価に活用',  price: '月3〜10万円',   bg: LTGRN },
  ];

  const colW = [1.4, 2.18, 1.0];
  const colX = [5.02, 6.42, 8.6];
  const colH = ['顧客', '活用方法', '単価'];
  colH.forEach((h, ci) => {
    card(s, colX[ci], 1.16, colW[ci], 0.28, DTXT);
    txt(s, h, colX[ci], 1.16, colW[ci], 0.28,
      { fontSize: 9.5, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  });

  customers.forEach((c, i) => {
    const y = 1.48 + i * 0.52;
    [c.seg, c.use, c.price].forEach((v, ci) => {
      card(s, colX[ci], y, colW[ci], 0.48, c.bg, LGREY);
      txt(s, v, colX[ci] + 0.06, y + 0.02, colW[ci] - 0.12, 0.44,
        { fontSize: ci === 2 ? 10.5 : 10, bold: ci === 2, color: ci === 2 ? GREEN : DTXT,
          valign: 'middle', lineSpacingMultiple: 1.2 });
    });
  });

  card(s, 0.4, 3.48, 9.2, 0.56, OFFWH, LGREY);
  txt(s, '初期戦略：まず調査会社1〜2社に無償サンプル提供 → ニーズ確認後に有償化。ヘッジファンドは紹介経由でのアプローチを優先。', 0.56, 3.5, 8.9, 0.52,
    { fontSize: 11, color: DTXT, lineSpacingMultiple: 1.4 });

  // Bottom chart: data flow
  slab(s, 'データ収集フロー', 0.4, 4.14, 9.2, 0.28, GREEN);
  const flow = ['施設で廃棄物収集', 'ラクハイに記録\n（品目・重量・施設）', 'DB蓄積・クレンジング', 'API / レポート販売'];
  flow.forEach((f, i) => {
    const x = 0.4 + i * 2.3;
    card(s, x, 4.3, 2.16, 1.1, i % 2 === 0 ? LTGRN : 'E8F5E9', LGREY);
    numCircle(s, i + 1, x + 0.84, 4.34, GREEN);
    txt(s, f, x + 0.06, 4.86, 2.04, 0.44,
      { fontSize: 10, color: DTXT, align: 'center', lineSpacingMultiple: 1.2 });
    if (i < 3) txt(s, '▶', x + 2.16, 4.68, 0.14, 0.36, { fontSize: 12, color: GREEN, align: 'center' });
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 5: 事業②　トラック走行・都市センシングデータ
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '事業②　トラック走行・都市センシングデータ', 3);

  // Left: sensor config
  card(s, 0.4, 0.82, 4.42, 2.2, LTTL, LGREY);
  slab(s, 'センサー構成（搭載機器）', 0.4, 0.82, 4.42, 0.3, TEAL);
  const sensors = [
    { icon: '録', label: 'フロントカメラ（AI画像解析）', desc: '街路樹・落下物・損傷を自動検出' },
    { icon: '震', label: '振動センサー', desc: '路面の凹凸・ポットホールを検知' },
    { icon: '位', label: 'GPS（既存）', desc: '走行ルート・速度・停車時間を記録' },
    { icon: '通', label: 'LTE通信モジュール', desc: '検知イベントをリアルタイム送信' },
  ];
  sensors.forEach((se, i) => {
    const y = 1.22 + i * 0.44;
    s.addShape(pres.ShapeType.ellipse, { x: 0.56, y: y + 0.04, w: 0.34, h: 0.34,
      fill: { color: TEAL }, line: { type: 'none' } });
    txt(s, se.icon, 0.56, y + 0.04, 0.34, 0.34,
      { fontSize: 10, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    txt(s, se.label, 1.0, y + 0.02, 3.6, 0.2,
      { fontSize: 10.5, bold: true, color: TEAL });
    txt(s, se.desc, 1.0, y + 0.22, 3.6, 0.2,
      { fontSize: 9.5, color: MTXT });
  });

  // Right: detection targets
  slab(s, '検知対象データ × 活用用途', 5.02, 0.82, 4.58, 0.3, TEAL);
  const detections = [
    { target: '街路樹の傾き・損傷', use: '自治体の樹木管理業務に代替', col: TEAL },
    { target: '落下物・不法投棄', use: '即時通報システムとして提供', col: CORAL },
    { target: '道路損傷・ひび割れ', use: '補修優先度の客観的評価に活用', col: GREEN },
    { target: '信号・標識の劣化', use: '老朽化インフラの定期点検補完', col: AMBER },
    { target: '交通量・渋滞パターン', use: '都市計画・交通政策の基礎データ', col: NAVY },
  ];
  detections.forEach((d, i) => {
    const y = 1.18 + i * 0.43;
    card(s, 5.02, y, 4.58, 0.4, i % 2 === 0 ? OFFWH : WHITE, LGREY);
    card(s, 5.02, y, 0.06, 0.4, d.col);
    txt(s, d.target, 5.14, y + 0.02, 2.0, 0.18,
      { fontSize: 10.5, bold: true, color: DTXT });
    txt(s, d.use, 5.14, y + 0.21, 4.36, 0.18,
      { fontSize: 10, color: MTXT });
  });

  // Customer section
  card(s, 0.4, 3.1, 4.42, 1.76, OFFWH, LGREY);
  slab(s, '顧客セグメント', 0.4, 3.1, 4.42, 0.3, TEAL);
  const tgtCust = [
    { c: '自治体（川崎市・横浜市等）', d: 'インフラ管理部門。定額月次契約。' },
    { c: '道路維持管理会社', d: 'データを活用した受注提案のツールとして。' },
    { c: '損害保険会社', d: '道路リスクマップの構築・料率算定に活用。' },
  ];
  tgtCust.forEach((c, i) => {
    dot(s, 0.56, 3.5 + i * 0.44 + 0.02, TEAL);
    txt(s, c.c, 0.78, 3.47 + i * 0.44, 4.0, 0.2,
      { fontSize: 11, bold: true, color: TEAL });
    txt(s, c.d, 0.78, 3.67 + i * 0.44, 4.0, 0.2,
      { fontSize: 10, color: MTXT });
  });

  card(s, 5.02, 3.4, 4.58, 1.5, LTTL, LGREY);
  slab(s, 'コスト試算（センサー搭載）', 5.02, 3.4, 4.58, 0.3, TEAL);
  const costs = [
    { item: 'AIカメラ+センサー（1台）', val: '5〜8万円' },
    { item: '初期搭載（3台 PoC）', val: '15〜25万円' },
    { item: '全台搭載（200台）', val: '〜1,600万円' },
    { item: 'クラウド解析費（月次）', val: '10〜30万円/月' },
  ];
  costs.forEach((c, i) => {
    const y = 3.78 + i * 0.29;
    dot(s, 5.18, y + 0.02, TEAL);
    txt(s, c.item, 5.4, y, 2.2, 0.27, { fontSize: 10, color: DTXT });
    txt(s, c.val, 7.66, y, 1.88, 0.27,
      { fontSize: 10, bold: true, color: TEAL, align: 'right' });
  });

  card(s, 0.4, 4.96, 9.2, 0.46, CORAL);
  txt(s, 'まず3〜5台で精度・費用対効果を検証（PoC）→ 自治体との実証契約を経て全台展開', 0.4, 4.96, 9.2, 0.46,
    { fontSize: 12, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}

// ══════════════════════════════════════════════════════════
// SLIDE 6: 競合優位性・参入障壁
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '競合優位性・参入障壁', 4);

  slab(s, 'サキュレにしかできない理由', 0.4, 0.82, 9.2, 0.32, CORAL);

  const advantages = [
    { n: 1, col: GREEN, title: '収集コスト = ほぼゼロ',
      body: '廃棄物収集業務の副産物としてデータが発生する。専用要員・専用車両は不要。既存業務を行うだけでデータが蓄積される。' },
    { n: 2, col: TEAL,  title: '許認可と施設立入権限',
      body: '廃棄物収集許可を持つサキュレだけが施設内に定期的に立ち入れる。競合他社は許可なく施設内データを取得できない。' },
    { n: 3, col: CORAL, title: '200台・都市圏の密なカバレッジ',
      body: '東京・神奈川エリアで200台以上が毎日走行。センサー密度が高いほど検知精度が上がり、単台数の競合が追いつけない。' },
    { n: 4, col: AMBER, title: '蓄積データの歴史的価値',
      body: '数年分の過去データを保有。トレンド分析・季節調整に必要な歴史的データは後発他社が短期間で揃えることができない。' },
  ];

  advantages.forEach((a, i) => {
    const x = i < 2 ? 0.4 : 5.15;
    const y = 1.26 + (i % 2) * 1.86;
    card(s, x, y, 4.55, 1.72, WHITE, LGREY);
    card(s, x, y, 0.06, 1.72, a.col);
    numCircle(s, a.n, x + 0.18, y + 0.58, a.col);
    txt(s, a.title, x + 0.78, y + 0.14, 3.66, 0.38,
      { fontSize: 13, bold: true, color: DTXT });
    hrule(s, x + 0.78, y + 0.58, 3.4, LGREY);
    txt(s, a.body, x + 0.78, y + 0.68, 3.66, 0.96,
      { fontSize: 11, color: MTXT, lineSpacingMultiple: 1.4 });
  });

  card(s, 0.4, 5.0, 9.2, 0.44, LTGRN, LGREY);
  card(s, 0.4, 5.0, 0.06, 0.44, GREEN);
  txt(s, '参入障壁のまとめ：許認可 × 既存施設アクセス × 過去蓄積データ × 台数規模の4条件が揃うのはサキュレのみ', 0.56, 5.0, 9.0, 0.44,
    { fontSize: 11, bold: true, color: GREEN, valign: 'middle' });
}

// ══════════════════════════════════════════════════════════
// SLIDE 7: 収益モデル・料金設計
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '収益モデル・料金設計', 5);

  slab(s, '事業①　廃棄物データ', 0.4, 0.82, 4.42, 0.3, GREEN);
  const wasteRev = [
    { type: 'データAPI 月額契約', price: '5〜100万円/月', note: 'ヘッジファンド・調査会社', col: LTGRN },
    { type: 'ESGデータレポート', price: '30〜100万円/件', note: '大企業・REIT向け年次開示', col: WHITE },
    { type: '自治体向け分析契約', price: '5〜20万円/月', note: '廃棄物量予測・政策立案用', col: LTGRN },
    { type: 'カスタム調査レポート', price: '50〜200万円/件', note: '特定施設・エリアの詳細分析', col: WHITE },
  ];
  wasteRev.forEach((r, i) => {
    const y = 1.18 + i * 0.52;
    card(s, 0.4, y, 4.42, 0.48, r.col, LGREY);
    card(s, 0.4, y, 0.06, 0.48, GREEN);
    txt(s, r.type, 0.56, y + 0.04, 2.0, 0.2, { fontSize: 10.5, bold: true, color: DTXT });
    txt(s, r.note, 0.56, y + 0.26, 2.0, 0.18, { fontSize: 9.5, color: MTXT });
    txt(s, r.price, 2.62, y + 0.08, 2.1, 0.32,
      { fontSize: 11.5, bold: true, color: GREEN, align: 'center', valign: 'middle' });
  });

  card(s, 0.4, 3.28, 4.42, 0.5, LTGRN, LGREY);
  txt(s, '1年目目標：年間収入500万円（月均41万円）', 0.56, 3.32, 4.26, 0.42,
    { fontSize: 11, bold: true, color: GREEN, valign: 'middle' });

  slab(s, '事業②　走行センサーデータ', 5.02, 0.82, 4.58, 0.3, TEAL);
  const sensorRev = [
    { type: '自治体 定額月次契約', price: '10〜50万円/月', note: '道路・インフラ管理部門', col: LTTL },
    { type: 'データ解析レポート', price: '50〜200万円/件', note: '損傷箇所一覧・優先度マップ', col: WHITE },
    { type: '損害保険会社向け', price: '20〜80万円/月', note: '道路リスクマップ提供', col: LTTL },
    { type: '道路管理会社向け', price: '5〜20万円/月', note: '点検代替・受注補完ツール', col: WHITE },
  ];
  sensorRev.forEach((r, i) => {
    const y = 1.18 + i * 0.52;
    card(s, 5.02, y, 4.58, 0.48, r.col, LGREY);
    card(s, 5.02, y, 0.06, 0.48, TEAL);
    txt(s, r.type, 5.18, y + 0.04, 2.08, 0.2, { fontSize: 10.5, bold: true, color: DTXT });
    txt(s, r.note, 5.18, y + 0.26, 2.08, 0.18, { fontSize: 9.5, color: MTXT });
    txt(s, r.price, 7.32, y + 0.08, 2.18, 0.32,
      { fontSize: 11.5, bold: true, color: TEAL, align: 'center', valign: 'middle' });
  });

  card(s, 5.02, 3.28, 4.58, 0.5, LTTL, LGREY);
  txt(s, '1年目目標：年間収入300万円（PoC後の本格化）', 5.18, 3.32, 4.42, 0.42,
    { fontSize: 11, bold: true, color: TEAL, valign: 'middle' });

  // Bottom: combined
  card(s, 0.4, 3.9, 9.2, 0.52, CORAL);
  txt(s, '2事業合計　1年目：800万円　2年目：1,500万円　3年目：3,000万円', 0.4, 3.9, 9.2, 0.52,
    { fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle' });

  txt(s, '※ 初期は無償PoC → 有償化の順で進める。契約更新率の維持が安定収益の鍵。', 0.4, 4.5, 9.2, 0.3,
    { fontSize: 9.5, color: MUTED });

  // Pricing notes
  card(s, 0.4, 4.82, 4.42, 0.64, LTGRN, LGREY);
  txt(s, '廃棄物データ初年度見込み\n顧客5社 × 平均8万円/月 → 年間480万円', 0.56, 4.84, 4.22, 0.6,
    { fontSize: 10.5, color: GREEN, lineSpacingMultiple: 1.5 });
  card(s, 5.02, 4.82, 4.58, 0.64, LTTL, LGREY);
  txt(s, '走行データ初年度見込み\nPoC2〜3件 → 有償移行で年間240〜360万円', 5.18, 4.84, 4.38, 0.6,
    { fontSize: 10.5, color: TEAL, lineSpacingMultiple: 1.5 });
}

// ══════════════════════════════════════════════════════════
// SLIDE 8: KGI / KPI・マイルストーン
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, 'KGI / KPI・マイルストーン', 6);

  const milestones = [
    { yr: '1年目', kpi: '年間売上800万円', sub: '廃棄物データ5社・走行PoC3件達成', col: CORAL },
    { yr: '2年目', kpi: '年間売上1,500万円', sub: 'センサー50台・データ顧客15社', col: GREEN },
    { yr: '3年目', kpi: '年間売上3,000万円', sub: '単年度黒字・第二収益柱確立', col: DKGRN },
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

  // KPI table
  slab(s, 'KPI管理指標', 0.4, 1.76, 9.2, 0.3, CORAL);

  const kpiRows = [
    { kpi: '廃棄物データ　有償顧客数', y1: '5社', y2: '15社', y3: '30社以上', type: 'pos' },
    { kpi: '廃棄物データ　月間API件数', y1: '試験運用', y2: '50件/月', y3: '200件/月', type: 'n' },
    { kpi: 'センサー搭載台数', y1: '3〜5台（PoC）', y2: '50台', y3: '200台（全台）', type: 'n' },
    { kpi: '走行データ　顧客数（自治体等）', y1: '2〜3件', y2: '8〜10件', y3: '20件以上', type: 'pos' },
    { kpi: '顧客継続率（年次更新）', y1: '—', y2: '85%以上', y3: '90%以上', type: 'pos' },
    { kpi: '2事業合計年間売上', y1: '800万円', y2: '1,500万円', y3: '3,000万円', type: 'profit' },
  ];

  const colsX = [0.4, 4.82, 6.5, 8.18];
  const colsW = [4.36, 1.62, 1.62, 1.62];
  const hdrs  = ['KPI', '1年目', '2年目', '3年目'];
  hdrs.forEach((h, ci) => {
    card(s, colsX[ci], 2.1, colsW[ci], 0.32, DTXT);
    txt(s, h, colsX[ci], 2.1, colsW[ci], 0.32,
      { fontSize: 10, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  });

  kpiRows.forEach((r, ri) => {
    const y = 2.46 + ri * 0.48;
    const bg = ri % 2 === 0 ? OFFWH : WHITE;
    const isP = r.type === 'profit';
    card(s, colsX[0], y, colsW[0], 0.44, isP ? LTGRN : bg);
    if (isP) card(s, colsX[0], y, 0.06, 0.44, GREEN);
    txt(s, r.kpi, colsX[0] + 0.14, y, colsW[0] - 0.2, 0.44,
      { fontSize: 11, bold: isP, color: isP ? GREEN : DTXT, valign: 'middle' });
    [r.y1, r.y2, r.y3].forEach((v, vi) => {
      card(s, colsX[vi + 1], y, colsW[vi + 1], 0.44, isP ? LTGRN : bg, LGREY);
      txt(s, v, colsX[vi + 1], y, colsW[vi + 1], 0.44,
        { fontSize: 11, bold: isP, color: isP ? POSGR : DTXT, align: 'center', valign: 'middle' });
    });
    hrule(s, colsX[0], y + 0.44, 9.4);
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 9: ロードマップ（3フェーズ）
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, OFFWH);
  pgTitle(s, 'ロードマップ（3フェーズ）', 7);

  const phases = [
    { n: 1, period: '〜6ヶ月\n（7〜12月）', title: 'データ整備・PoC',
      col: CORAL, bg: WHITE,
      actions: [
        'ラクハイデータの品目別・施設別集計整備',
        'センサー3〜5台搭載・精度検証（社内）',
        '調査会社1〜2社に無償サンプル提供',
        '自治体1市にPoC提案・接触開始',
      ],
      goal: '有償化候補 3件以上確保' },
    { n: 2, period: '7〜18ヶ月\n（2027年1〜6月）', title: 'パイロット販売',
      col: GREEN, bg: WHITE,
      actions: [
        '廃棄物データ 有償契約2〜3社を締結',
        'センサー搭載台数を50台に拡大',
        '自治体1〜2市と実証事業契約',
        'APIインフラ・ダッシュボード構築',
      ],
      goal: '月間売上50万円以上（安定化）' },
    { n: 3, period: '19ヶ月〜\n（2027年7月〜）', title: 'スケールアップ',
      col: TEAL, bg: WHITE,
      actions: [
        'データ顧客を30社以上へ拡大',
        'センサー全台搭載（200台）',
        'データ販売の標準プランを整備・自動化',
        '他社廃棄物事業者へのライセンス販売も検討',
      ],
      goal: '年間売上3,000万円・第二収益柱確立' },
  ];

  phases.forEach((ph, i) => {
    const x = 0.3 + i * 3.23;
    card(s, x, 0.86, 3.08, 4.5, ph.bg, LGREY);
    card(s, x, 0.86, 3.08, 0.06, ph.col);
    numCircle(s, ph.n, x + 1.28, 0.96, ph.col);
    txt(s, ph.period, x, 1.56, 3.08, 0.4,
      { fontSize: 9.5, color: MTXT, align: 'center', lineSpacingMultiple: 1.3 });
    txt(s, ph.title, x, 2.0, 3.08, 0.42,
      { fontSize: 15, bold: true, color: ph.col, align: 'center' });
    hrule(s, x + 0.15, 2.48, 2.78, ph.col);
    ph.actions.forEach((a, j) => {
      dot(s, x + 0.2, 2.6 + j * 0.44 + 0.02, ph.col);
      txt(s, a, x + 0.38, 2.57 + j * 0.44, 2.6, 0.4,
        { fontSize: 10.5, color: DTXT, lineSpacingMultiple: 1.25 });
    });
    card(s, x + 0.12, 4.64, 2.84, 0.48, ph.col);
    txt(s, ph.goal, x + 0.12, 4.64, 2.84, 0.48,
      { fontSize: 10, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    if (i < 2) {
      txt(s, '▶', x + 3.08, 2.86, 0.15, 0.4,
        { fontSize: 14, color: CORAL, align: 'center' });
    }
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 10: 財務シミュレーション（3カ年）
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '財務シミュレーション（3カ年）', 8);

  slab(s, '3カ年 損益シミュレーション（2事業合算）', 0.3, 0.82, 9.4, 0.3, GREEN);

  const LX = 0.3, LW = 3.22;
  const YX = [3.56, 5.56, 7.56];
  const YW = 1.94;
  const RH = 0.52;

  card(s, LX, 1.16, LW, 0.5, DTXT);
  txt(s, '項目', LX, 1.16, LW, 0.5,
    { fontSize: 10, bold: true, color: WHITE, align: 'center', valign: 'middle' });

  const yLabels = ['1年目', '2年目', '3年目'];
  const yBadges = ['PoC・試験販売', '本格展開', '★黒字化'];
  YX.forEach((x, i) => {
    const sp = i === 2;
    card(s, x, 1.16, YW, 0.5, sp ? GREEN : DTXT);
    if (sp) {
      txt(s, yLabels[i], x, 1.18, YW, 0.24,
        { fontSize: 10, bold: true, color: WHITE, align: 'center' });
      txt(s, yBadges[i], x, 1.40, YW, 0.22,
        { fontSize: 9, bold: true, color: WHITE, align: 'center' });
    } else {
      txt(s, yLabels[i], x, 1.18, YW, 0.24,
        { fontSize: 10, bold: true, color: WHITE, align: 'center' });
      txt(s, yBadges[i], x, 1.40, YW, 0.22,
        { fontSize: 9, color: LGREY, align: 'center' });
    }
  });

  const plRows = [
    { label: '廃棄物データ収入',    vals: ['300万円',  '900万円',  '1,500万円'], kind: 'pos' },
    { label: '走行センサー収入',     vals: ['200万円',  '600万円',  '1,500万円'], kind: 'pos' },
    { label: '売上合計',             vals: ['500万円',  '1,500万円','3,000万円'], kind: 'total' },
    { label: '初期投資（センサー等）',vals: ['▲300万円','▲200万円', '▲400万円'], kind: 'neg' },
    { label: 'データ整備・API運用費', vals: ['▲120万円','▲180万円', '▲240万円'], kind: 'neg' },
    { label: '営業・人件費',          vals: ['▲180万円','▲200万円', '▲220万円'], kind: 'neg' },
    { label: '営業利益（単年）',       vals: ['▲100万円','+920万円', '+2,140万円'], kind: 'profit' },
  ];

  plRows.forEach((r, ri) => {
    const y = 1.70 + ri * RH;
    const isProfit = r.kind === 'profit';
    const isTotal  = r.kind === 'total';
    const rowBg = isProfit ? LTGRN : (isTotal ? LTBL : (ri % 2 === 0 ? OFFWH : WHITE));
    card(s, LX, y, LW, RH, rowBg);
    if (isProfit) card(s, LX, y, 0.06, RH, GREEN);
    if (isTotal)  card(s, LX, y, 0.06, RH, CORAL);
    txt(s, r.label, LX + 0.14, y, LW - 0.2, RH,
      { fontSize: 11, bold: isProfit || isTotal, color: isProfit ? GREEN : DTXT, valign: 'middle' });

    r.vals.forEach((v, vi) => {
      const x = YX[vi];
      const isNeg = v.startsWith('▲');
      const isPos = v.startsWith('+') || (!isNeg && (r.kind === 'pos' || r.kind === 'total'));
      const cellBg = isProfit
        ? (isNeg ? LTRED : LTGRN)
        : (isTotal ? LTBL : rowBg);
      card(s, x, y, YW, RH, cellBg, LGREY);
      const valColor = isProfit
        ? (isNeg ? RED : POSGR)
        : (r.kind === 'neg' ? RED : (r.kind === 'pos' || r.kind === 'total' ? POSGR : DTXT));
      txt(s, v, x, y, YW, RH,
        { fontSize: isProfit || isTotal ? 13 : 11.5,
          bold: isProfit || isTotal,
          color: valColor, align: 'center', valign: 'middle' });
    });
    hrule(s, LX, y + RH, LW + YX.length * YW + 0.1);
  });

  txt(s, '※ 初期投資は廃棄物データ整備・APIシステム・センサー機器を含む。2年目以降センサー展開で投資継続。',
    0.3, 5.32, 9.4, 0.24, { fontSize: 8.5, color: MUTED });
}

// ══════════════════════════════════════════════════════════
// SLIDE 11: 次のアクション
// ══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, OFFWH);
  pgTitle(s, '次のアクション（7〜9月）', 9);

  const actions = [
    { cat: '廃棄物データ整備', col: GREEN, items: [
      'ラクハイデータの品目別・施設別・月次クレンジング（7月）',
      '営業用サンプルレポート1本作成（8月）',
      '調査会社・ヘッジファンドへのアプローチリスト作成（7月）',
    ]},
    { cat: '走行センサーPoC', col: TEAL, items: [
      'AIカメラ・振動センサーの見積もり取得（3〜5社比較）（7月）',
      'センサー3台搭載・社内精度テスト実施（8〜9月）',
      '川崎市・横浜市のインフラ管理部門への打診（8月）',
    ]},
    { cat: '顧客開拓', col: CORAL, items: [
      '廃棄物データ：無償PoC提案 3件以上（8〜9月）',
      '走行データ：自治体担当者ヒアリング 2件以上（9月）',
      '既存250社取引先への廃棄物データ活用提案（9月）',
    ]},
    { cat: '体制・法務', col: AMBER, items: [
      'データ販売に必要な個人情報・秘密情報の取扱規程確認',
      'データ提供契約書ひな形の作成（弁護士と確認）',
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
  txt(s, '9月末チェック：サンプルレポート完成・PoC提案3件以上実施・センサーPoC開始', 0.4, 5.18, 9.2, 0.3,
    { fontSize: 11, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}

// ─── Output ───────────────────────────────────────────────
const OUT = '/tmp/claude-0/-home-user-sacure/1db07409-6d1e-5d98-92a7-465fca194ddc/scratchpad/データ事業計画書_v1.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('Done:', OUT));
