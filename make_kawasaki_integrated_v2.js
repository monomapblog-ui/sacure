const pptx = require('pptxgenjs');
const pres = new pptx();
pres.layout = 'LAYOUT_16x9'; // 10" × 5.625"

// ─── Palette（事業計画書FMT準拠） ─────────────────────────
const CORAL = 'DD7E6B';
const DKCRL = 'AA5848';
const GREEN = '156818';
const TEAL  = '0D7480';  // 走行センサーデータ識別色
const LTGRN = 'D9EAD3';
const LTTL  = 'D6EEF0';
const WHITE = 'FFFFFF';
const OFFWH = 'F3F3F3';
const DTXT  = '333333';
const MTXT  = '595959';
const MUTED = '999999';
const LGREY = 'CCCCCC';

// ─── Helpers ──────────────────────────────────────────────
function card(s, x, y, w, h, fill, lineColor) {
  s.addShape(pres.ShapeType.rect, { x, y, w, h,
    fill: { color: fill || WHITE },
    line: lineColor ? { color: lineColor, width: 0.75 } : { type: 'none' } });
}
function txt(s, text, x, y, w, h, opts) {
  s.addText(text, { x, y, w, h, fontFace: 'Calibri', ...opts });
}
// 全幅コーラルヘッダーバー + 白丸セクション番号
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
    fill: { color: col || CORAL }, line: { type: 'none' } });
}
function numCircle(s, n, x, y, col) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: 0.48, h: 0.48,
    fill: { color: col || CORAL }, line: { type: 'none' } });
  txt(s, String(n), x, y, 0.48, 0.48,
    { fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}
function kanjiCircle(s, kanji, x, y, col, sz) {
  const d = sz || 0.6;
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d,
    fill: { color: col || CORAL }, line: { type: 'none' } });
  txt(s, kanji, x, y, d, d,
    { fontSize: sz ? Math.round(sz * 28) : 17, bold: true, color: WHITE, align: 'center', valign: 'middle' });
}
function hrule(s, x, y, w, col) {
  s.addShape(pres.ShapeType.line, { x, y, w, h: 0,
    line: { color: col || LGREY, width: 0.5 } });
}
function accentBar(s, x, y, h, col) {
  card(s, x, y, 0.06, h, col || CORAL);
}

// ═══════════════════════════════════════════════════════════
// SLIDE 1: 表紙
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  card(s, 0, 0, 0.08, 5.625, CORAL);
  card(s, 0.08, 0, 9.92, 0.9, CORAL);

  txt(s, 'Kawasaki Future Co-Lab　応募提案書', 0.32, 0.14, 9, 0.3,
    { fontSize: 11, color: WHITE });
  txt(s, '廃棄物データ ×\n都市センシング構想', 0.32, 0.98, 6.8, 1.1,
    { fontSize: 28, bold: true, color: DTXT, lineSpacingMultiple: 1.2 });
  txt(s, '廃棄物量データ収集 × 走行センサーデータ収集', 0.32, 2.12, 6.8, 0.4,
    { fontSize: 15, color: CORAL });
  txt(s, '2つのデータで川崎市の都市管理・環境政策を支える', 0.32, 2.54, 6.8, 0.36,
    { fontSize: 13, color: MTXT });

  hrule(s, 0.32, 3.02, 3.5);

  txt(s, '株式会社サキュレ', 0.32, 3.16, 5, 0.38,
    { fontSize: 14, bold: true, color: DTXT });
  txt(s, '担当：村田　捷樹', 0.32, 3.56, 5, 0.32,
    { fontSize: 11, color: MTXT });
  txt(s, '2026年7月', 0.32, 3.9, 5, 0.28,
    { fontSize: 11, color: MUTED });

  // 右側: 3つのキー数字
  const nums = [
    { val: '200台以上', sub: '保有トラック（東京・神奈川全域）', col: CORAL },
    { val: '250社超',  sub: '廃棄物収集施設（東京・神奈川全域）', col: GREEN },
    { val: '2事業',    sub: '川崎市への統合データ提案',          col: TEAL },
  ];
  nums.forEach((n, i) => {
    const y = 1.05 + i * 1.18;
    card(s, 7.5, y, 2.2, 1.0, OFFWH, LGREY);
    card(s, 7.5, y, 0.06, 1.0, n.col);
    txt(s, n.val, 7.58, y + 0.07, 2.1, 0.52,
      { fontSize: 20, bold: true, color: n.col, align: 'center' });
    txt(s, n.sub, 7.58, y + 0.6, 2.1, 0.34,
      { fontSize: 9.5, color: MTXT, align: 'center' });
  });

  card(s, 0.08, 4.85, 9.92, 0.775, OFFWH);
  txt(s, '廃棄物量データ　×　道路センシングデータ　×　川崎市との共創', 0.08, 4.85, 9.92, 0.775,
    { fontSize: 12, bold: true, color: CORAL, align: 'center', valign: 'middle' });

  s.addNotes('表紙。廃棄物量データ収集と走行センサーデータ収集の2事業を統合した川崎市向け提案。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 2: 会社概要
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '株式会社サキュレが提供できる価値');

  // Left: CORAL company panel
  card(s, 0.4, 0.85, 2.8, 4.55, CORAL);
  txt(s, '株式会社サキュレ', 0.55, 0.98, 2.5, 0.44,
    { fontSize: 13, bold: true, color: WHITE });
  txt(s, '廃棄物収集・運搬・中間処理事業\n運送事業', 0.55, 1.44, 2.5, 0.62,
    { fontSize: 11, color: 'F8DDCC', lineSpacingMultiple: 1.35 });
  hrule(s, 0.55, 2.1, 2.5, 'F0C4B0');

  const stats = [
    { val: '200台以上', label: '保有車両（東京・神奈川全域）' },
    { val: '250社超',   label: '法人取引先（東京・神奈川全域）' },
    { val: '東京・神奈川', label: '主要営業エリア', sm: true },
  ];
  stats.forEach((st, i) => {
    const y = 2.24 + i * 0.96;
    txt(s, st.val, 0.55, y, 2.5, 0.46,
      { fontSize: st.sm ? 15 : 20, bold: true, color: WHITE });
    txt(s, st.label, 0.55, y + (st.sm ? 0.26 : 0.34), 2.5, 0.28,
      { fontSize: 10, color: 'F8DDCC' });
  });

  // Right: 3 pillar cards
  const pillars = [
    { kanji: '網', col: CORAL,
      title: '地域密着の配送ネットワーク',
      body: '東京・神奈川を日常的にカバーする200台以上の車両網。既存インフラをそのままデータ収集基盤へ転用できます。' },
    { kanji: '循', col: GREEN,
      title: '廃棄物収集許可業者ならではの独自取得権',
      body: '廃棄物収集許可を持つ業者のみが施設内のデータを合法的に記録可能。当社ならではの独自ネットワークで継続的に取得できます。' },
    { kanji: '信', col: TEAL,
      title: '行政・企業との協業実績',
      body: '廃棄物収集業務を通じた自治体・企業との継続的な信頼関係。コンプライアンスを重視した誠実な事業運営を行っています。' },
  ];
  pillars.forEach((p, i) => {
    const y = 0.85 + i * 1.52;
    card(s, 3.38, y, 6.22, 1.42, OFFWH, LGREY);
    card(s, 3.38, y, 6.22, 0.06, p.col);
    kanjiCircle(s, p.kanji, 3.52, y + 0.38, p.col, 0.58);
    txt(s, p.title, 4.24, y + 0.1, 5.25, 0.42,
      { fontSize: 13, bold: true, color: DTXT });
    txt(s, p.body, 4.24, y + 0.56, 5.25, 0.78,
      { fontSize: 10.5, color: MTXT, lineSpacingMultiple: 1.35 });
  });

  s.addNotes('会社紹介。3つの強み：①地域密着の配送網（200台以上）、②廃棄物収集許可業者として独占取得、③行政・企業との協業実績。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 3: 2事業の提案概要
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '2つのデータ収集事業を川崎市へご提案します', 1);

  // Shared asset banner
  card(s, 0.4, 0.82, 9.2, 0.52, OFFWH, LGREY);
  accentBar(s, 0.4, 0.82, 0.52, CORAL);
  txt(s, '共通資産：200台以上の配送トラック × 250社超の廃棄物収集ネットワーク', 0.62, 0.82, 8.98, 0.52,
    { fontSize: 12.5, bold: true, color: DTXT, valign: 'middle' });

  // Left card: 廃棄物量データ
  card(s, 0.4, 1.46, 4.45, 3.94, WHITE, LGREY);
  card(s, 0.4, 1.46, 4.45, 0.06, GREEN);
  slab(s, '廃棄物量データ収集事業', 0.4, 1.52, 4.45, 0.34, GREEN);
  kanjiCircle(s, '廃', 0.56, 1.98, GREEN, 0.56);
  txt(s, '廃棄物量データ収集事業', 1.26, 2.04, 3.45, 0.4,
    { fontSize: 14, bold: true, color: GREEN });

  const wastePoints = [
    '250施設の廃棄物量を定期記録',
    '施設別・品目別・時系列データ',
    '環境KPI・廃棄物管理施策に活用',
    '追加人員・設備不要（既存業務から記録）',
    '廃棄物収集許可業者のみが取得可能なデータ',
  ];
  wastePoints.forEach((pt, i) => {
    dot(s, 0.56, 2.58 + i * 0.48 + 0.04, GREEN);
    txt(s, pt, 0.82, 2.58 + i * 0.48, 3.88, 0.44,
      { fontSize: 11.5, color: DTXT });
  });

  // Right card: 走行センサーデータ
  card(s, 5.15, 1.46, 4.45, 3.94, WHITE, LGREY);
  card(s, 5.15, 1.46, 4.45, 0.06, TEAL);
  slab(s, '走行センサーデータ収集事業', 5.15, 1.52, 4.45, 0.34, TEAL);
  kanjiCircle(s, '路', 5.3, 1.98, TEAL, 0.56);
  txt(s, '走行センサーデータ収集事業', 6.0, 2.04, 3.45, 0.4,
    { fontSize: 14, bold: true, color: TEAL });

  const sensorPoints = [
    '200台のトラックにカメラ＋エッジAI搭載',
    '道路損傷・街路樹・交通状況を自動収集',
    '道路維持管理・交通施策に活用可能',
    '専用調査車両の追加購入不要',
    '個人情報はリアルタイムマスキング処理',
  ];
  sensorPoints.forEach((pt, i) => {
    dot(s, 5.3, 2.58 + i * 0.48 + 0.04, TEAL);
    txt(s, pt, 5.56, 2.58 + i * 0.48, 3.88, 0.44,
      { fontSize: 11.5, color: DTXT });
  });

  s.addNotes('2事業の概要。左：廃棄物量データ（250施設の収集記録）、右：走行センサーデータ（200台にカメラ搭載）。どちらも既存オペレーションからデータを生み出す。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 4: 川崎市が直面する課題
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '川崎市が直面する3つのデータ課題', 2);

  const issues = [
    { n: 1, col: CORAL,
      title: '市道約2,400kmを少数体制でパトロール——カバー率の限界',
      body: '川崎市が管理する市道は約2,400kmに及びます。道路パトロール要員の不足が続くなか、全路線を高頻度でカバーすることは困難で、損傷の見落としや対応遅延が常在するリスクとなっています。' },
    { n: 2, col: CORAL,
      title: '市民通報頼りの「事後対応」——長寿命化計画が掲げる予防保全の実現が急務',
      body: '損傷の多くが市民通報を起点に発覚するため、発生から対応完了まで時間を要します。川崎市が推進する「道路インフラ長寿命化計画」の目標達成には、予防的なリアルタイム損傷検知の仕組みが不可欠です。' },
    { n: 3, col: GREEN,
      title: '環境施策・廃棄物管理のデータ不足',
      body: '廃棄物排出量の変動や施設別の廃棄物動向を定量的に把握する手段が乏しく、環境KPIの達成状況の検証や廃棄物管理施策の根拠となるデータの整備が求められています。' },
  ];

  issues.forEach((iss, i) => {
    const y = 0.88 + i * 1.54;
    card(s, 0.4, y, 9.2, 1.38, WHITE, LGREY);
    card(s, 0.4, y, 9.2, 0.06, iss.col);
    numCircle(s, iss.n, 0.54, y + 0.43, iss.col);
    txt(s, iss.title, 1.16, y + 0.14, 8.25, 0.42,
      { fontSize: 14, bold: true, color: DTXT });
    txt(s, iss.body, 1.16, y + 0.6, 8.25, 0.68,
      { fontSize: 11.5, color: MTXT, lineSpacingMultiple: 1.35 });
  });

  s.addNotes('3つの課題：①市道約2,400kmのカバー率限界（川崎市固有数値）、②道路インフラ長寿命化計画が掲げる予防保全への転換、③環境施策・廃棄物管理のデータ不足。サキュレの2事業で解決を提案。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 5: 統合データフロー
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '2つの事業が生み出すデータフロー', 3);

  // Shared source
  card(s, 0.4, 0.82, 9.2, 0.52, OFFWH, LGREY);
  accentBar(s, 0.4, 0.82, 0.52, CORAL);
  txt(s, '株式会社サキュレ　200台以上の配送トラック × 250社超の廃棄物収集ネットワーク', 0.62, 0.82, 8.98, 0.52,
    { fontSize: 12, bold: true, color: DTXT, valign: 'middle' });

  // Left: 廃棄物量データ flow
  card(s, 0.4, 1.46, 4.45, 3.94, WHITE, LGREY);
  card(s, 0.4, 1.46, 4.45, 0.06, GREEN);
  slab(s, '廃棄物量データ', 0.4, 1.52, 4.45, 0.32, GREEN);

  const wasteSteps = [
    { n: 1, label: '廃棄物の収集',   sub: '施設別に量・品目を記録' },
    { n: 2, label: 'データ分析・整備', sub: '時系列・カテゴリ別に集計' },
    { n: 3, label: '川崎市へ提供',   sub: '廃棄物管理部門・環境政策に活用' },
  ];
  wasteSteps.forEach((st, i) => {
    const y = 1.96 + i * 1.08;
    numCircle(s, st.n, 0.54, y, GREEN);
    if (i < 2) {
      s.addShape(pres.ShapeType.line, { x: 0.72, y: y + 0.52, w: 0, h: 0.52,
        line: { color: LGREY, width: 1.2 } });
    }
    txt(s, st.label, 1.16, y + 0.04, 3.55, 0.36,
      { fontSize: 13, bold: true, color: DTXT });
    txt(s, st.sub, 1.16, y + 0.42, 3.55, 0.28,
      { fontSize: 10.5, color: MTXT });
    if (i < 2) hrule(s, 0.54, y + 1.0, 4.12);
  });

  // Right: 走行センサーデータ flow
  card(s, 5.15, 1.46, 4.45, 3.94, WHITE, LGREY);
  card(s, 5.15, 1.46, 4.45, 0.06, TEAL);
  slab(s, '走行センサーデータ', 5.15, 1.52, 4.45, 0.32, TEAL);

  const sensorSteps = [
    { n: 1, label: '配送トラック走行',   sub: '搭載カメラで撮影・記録' },
    { n: 2, label: 'エッジAI処理',       sub: '個人情報を車内でマスキング' },
    { n: 3, label: 'クラウドへ送信',     sub: '処理済みデータのみ転送' },
    { n: 4, label: '行政ダッシュボード', sub: '地図上で可視化・アラート通知' },
  ];
  sensorSteps.forEach((st, i) => {
    const y = 1.96 + i * 0.82;
    numCircle(s, st.n, 5.28, y, TEAL);
    if (i < 3) {
      s.addShape(pres.ShapeType.line, { x: 5.46, y: y + 0.52, w: 0, h: 0.26,
        line: { color: LGREY, width: 1.2 } });
    }
    txt(s, st.label, 5.9, y + 0.04, 3.55, 0.34,
      { fontSize: 12, bold: true, color: DTXT });
    txt(s, st.sub, 5.9, y + 0.4, 3.55, 0.26,
      { fontSize: 10.5, color: MTXT });
    if (i < 3) hrule(s, 5.28, y + 0.76, 4.12);
  });

  s.addNotes('2事業のデータフロー。左：廃棄物量データ（3ステップ）、右：走行センサーデータ（4ステップ）。どちらも既存オペレーションから収集。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 6: 走行センサーで集まるデータ（3分野）
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, 'トラックが走るだけで3分野のデータが集まる', 4);

  // Left: truck schematic
  card(s, 0.4, 0.85, 3.0, 4.55, OFFWH, LGREY);
  card(s, 0.4, 0.85, 3.0, 0.06, TEAL);
  // Truck body
  card(s, 0.65, 1.52, 2.5, 1.4, TEAL);
  card(s, 0.65, 2.62, 2.5, 0.38, DKCRL);
  s.addShape(pres.ShapeType.ellipse, { x: 0.82, y: 2.94, w: 0.42, h: 0.42,
    fill: { color: DTXT }, line: { type: 'none' } });
  s.addShape(pres.ShapeType.ellipse, { x: 2.54, y: 2.94, w: 0.42, h: 0.42,
    fill: { color: DTXT }, line: { type: 'none' } });
  card(s, 1.45, 1.32, 0.88, 0.26, CORAL);
  txt(s, 'カメラ搭載', 1.45, 1.32, 0.88, 0.26,
    { fontSize: 8, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  txt(s, '配送トラック', 0.4, 3.46, 3.0, 0.3,
    { fontSize: 13, bold: true, color: DTXT, align: 'center' });
  txt(s, '走行しながら自動収集', 0.4, 3.78, 3.0, 0.28,
    { fontSize: 10.5, color: TEAL, align: 'center', charSpacing: 0.5 });

  // Arrows
  [1.72, 2.86, 4.0].forEach(ay => {
    s.addShape(pres.ShapeType.line, { x: 3.42, y: ay, w: 0.48, h: 0,
      line: { color: CORAL, width: 1.5 } });
    card(s, 3.88, ay - 0.12, 0.2, 0.2, CORAL);
    txt(s, '▶', 3.88, ay - 0.12, 0.2, 0.2,
      { fontSize: 9, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  });

  // Right: 3 data categories
  const cats = [
    { kanji: '路', col: CORAL, name: '道路インフラ',
      kw: '路面陥没・ひび割れ・落下物・補修優先度マップ' },
    { kanji: '緑', col: GREEN, name: '都市景観・緑地',
      kw: '街路樹の状態・看板視認性・メンテナンス要否' },
    { kanji: '交', col: TEAL, name: '交通・モビリティ',
      kw: '渋滞区間・路上駐車・交通施策の効果測定' },
  ];
  cats.forEach((cat, i) => {
    const ry = 0.98 + i * 1.48;
    if (i > 0) hrule(s, 4.14, ry - 0.1, 5.5);
    kanjiCircle(s, cat.kanji, 4.14, ry, cat.col, 0.62);
    txt(s, cat.name, 4.9, ry + 0.04, 4.72, 0.42,
      { fontSize: 17, bold: true, color: cat.col });
    txt(s, cat.kw, 4.9, ry + 0.52, 4.72, 0.6,
      { fontSize: 12.5, color: MTXT });
  });

  s.addNotes('走行センサーデータの3分野：路（道路インフラ）・緑（都市景観）・交（交通モビリティ）。既存トラックの日常走行から自動収集。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 7: 廃棄物量データが生み出す価値
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '廃棄物収集業務が生み出す都市環境データ', 5);

  // Left explanation
  txt(s, '廃棄物量を\n都市の環境指標に変える', 0.4, 0.86, 4.2, 0.85,
    { fontSize: 18, bold: true, color: DTXT, lineSpacingMultiple: 1.25 });
  txt(s, '廃棄物収集許可業者として、\n施設内の廃棄物量・品目・頻度を\n合法的に記録・蓄積できます。\n自治体では入手できない施設別の\nリアルタイム廃棄物動向データを\n定期的に提供します。', 0.4, 1.76, 4.2, 1.5,
    { fontSize: 12, color: MTXT, lineSpacingMultiple: 1.42 });

  // Key benefit box
  card(s, 0.4, 3.38, 4.2, 1.98, GREEN);
  card(s, 0.4, 3.38, 0.06, 1.98, CORAL);
  txt(s, '追加センサー・人員・設備は不要\n収集業務の「ついで」に\n都市環境データを自動生成', 0.64, 3.38, 3.96, 1.98,
    { fontSize: 13, bold: true, color: WHITE, valign: 'middle', lineSpacingMultiple: 1.5 });

  // Right: 3 data value items
  const dataVals = [
    { kanji: '量', col: GREEN, name: '廃棄物排出量トレンド',
      kw: '施設別・品目別の月次推移。環境KPI測定・廃棄物削減施策の効果検証に活用' },
    { kanji: '比', col: CORAL, name: '施設間比較・ベンチマーク',
      kw: '同業種・同規模施設との廃棄物量比較。省資源化の優先施設を特定可能' },
    { kanji: '域', col: TEAL, name: '地域別廃棄物動向',
      kw: '川崎市内エリアごとの廃棄物量分布。施策立案・収集ルート最適化に貢献' },
  ];
  dataVals.forEach((d, i) => {
    const ry = 0.86 + i * 1.56;
    if (i > 0) hrule(s, 4.82, ry - 0.1, 4.82);
    kanjiCircle(s, d.kanji, 4.82, ry, d.col, 0.6);
    txt(s, d.name, 5.56, ry + 0.04, 4.08, 0.4,
      { fontSize: 15, bold: true, color: d.col });
    txt(s, d.kw, 5.56, ry + 0.52, 4.08, 0.68,
      { fontSize: 11.5, color: MTXT, lineSpacingMultiple: 1.35 });
  });

  s.addNotes('廃棄物量データの3つの価値：量（排出量トレンド）・比（施設間比較）・域（地域別動向）。廃棄物収集許可業者として合法的・独自ネットワークで取得できるデータ。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 8: 200台ネットワークの優位性
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '200台ネットワークがもたらす3つの優位性', 6);

  const strengths = [
    { n: 1, col: CORAL, title: '網羅性と継続性',
      points: ['高頻度・日常的なエリアカバー', '体系的・定期的なデータ収集', '長期的な時系列データ蓄積'] },
    { n: 2, col: GREEN,  title: '追加投資不要の経済性',
      points: ['専用車両の追加購入不要', '既存業務に付随してデータ収集', '運用コストを大幅に抑制'] },
    { n: 3, col: TEAL,  title: '現場知と独自の取得権',
      points: ['廃棄物収集許可業者のみ合法取得可', '現地の実態を熟知した人材', '行政・企業との長期信頼関係'] },
  ];

  strengths.forEach((st, i) => {
    const x = 0.4 + i * 3.2;
    card(s, x, 0.82, 2.95, 4.58, OFFWH, LGREY);
    card(s, x, 0.82, 2.95, 0.06, st.col);
    txt(s, String(st.n), x, 0.92, 2.95, 1.04,
      { fontSize: 56, bold: true, color: st.col, align: 'center', transparency: 12 });
    txt(s, st.title, x + 0.15, 2.0, 2.65, 0.52,
      { fontSize: 14, bold: true, color: st.col, align: 'center' });
    hrule(s, x + 0.2, 2.56, 2.55);
    st.points.forEach((pt, j) => {
      dot(s, x + 0.2, 2.72 + j * 0.6 + 0.04, st.col);
      txt(s, pt, x + 0.4, 2.69 + j * 0.6, 2.4, 0.52,
        { fontSize: 11, color: DTXT });
    });
  });

  s.addNotes('3つの優位性：①網羅性と継続性、②追加投資不要の経済性、③現場知と独自の取得権（廃棄物収集許可業者）。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 9: 個人情報・法令対応の方針
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '個人情報・プライバシーへの対応方針', 7);

  // Left: policy panel
  card(s, 0.4, 0.82, 3.6, 4.58, CORAL);
  card(s, 0.4, 0.82, 0.06, 4.58, DKCRL);
  txt(s, '個人情報を\n取得しない設計', 0.64, 1.02, 3.2, 0.82,
    { fontSize: 17, bold: true, color: WHITE, lineSpacingMultiple: 1.3 });
  txt(s, '走行センサーシステムは道路インフラ管理に特化した設計のもと、歩行者・車両の個人識別情報を一切保持しない運用体制を採ります。廃棄物量データは施設単位の集計値のみを扱います。', 0.64, 1.9, 3.2, 1.12,
    { fontSize: 11, color: 'F8DDCC', lineSpacingMultiple: 1.4 });

  txt(s, '準拠する主な法令', 0.64, 3.1, 3.1, 0.3,
    { fontSize: 9.5, color: 'F8DDCC', charSpacing: 0.5, bold: true });
  const laws = [
    '個人情報保護法 第17条（適正な取得）',
    '個人情報保護法 第27条（第三者提供制限）',
    '川崎市個人情報保護条例 への準拠',
  ];
  laws.forEach((law, i) => {
    dot(s, 0.64, 3.48 + i * 0.36 + 0.04, 'F8DDCC');
    txt(s, law, 0.9, 3.44 + i * 0.36, 3.0, 0.34,
      { fontSize: 9.5, color: 'F8DDCC' });
  });

  // Right: 3 compliance items
  const items = [
    { n: 1, col: CORAL, title: 'エッジAIによる自動マスキング',
      body: '歩行者の顔・車両ナンバープレートをカメラ搭載のエッジAIがリアルタイムで自動処理。個人を識別できる情報は収集・保存しません。' },
    { n: 2, col: GREEN,  title: 'データの暗号化と厳格なアクセス管理',
      body: '収集データは暗号化のうえクラウド管理。アクセス権限は最小限に設定し、定期的なセキュリティ監査を実施します。' },
    { n: 3, col: TEAL,  title: '行政との協議による運用体制の整備',
      body: '川崎市情報セキュリティ担当部署と協議のうえ、データ管理規程・利用目的・保存期間・消去プロセスを明文化します。' },
  ];
  items.forEach((item, i) => {
    const y = 0.82 + i * 1.52;
    card(s, 4.18, y, 5.42, 1.42, WHITE, LGREY);
    card(s, 4.18, y, 5.42, 0.06, item.col);
    numCircle(s, item.n, 4.34, y + 0.46, item.col);
    txt(s, item.title, 4.96, y + 0.1, 4.5, 0.42,
      { fontSize: 13, bold: true, color: DTXT });
    txt(s, item.body, 4.96, y + 0.56, 4.5, 0.78,
      { fontSize: 11, color: MTXT, lineSpacingMultiple: 1.35 });
  });

  s.addNotes('プライバシー・法令対応。3つの対策：①エッジAIによる自動マスキング、②暗号化・アクセス管理、③行政との協議による運用体制整備。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 10: 実証実験（PoC）の提案
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '小さく始め、効果を定量的に検証する', 8);

  // 4 parameter boxes
  const pocs = [
    { label: '対象エリア',     val: '川崎区・中原区\n幹線道路を優先指定', col: CORAL },
    { label: '実施期間',       val: '第1期  3ヶ月\n第2期  3〜6ヶ月',     col: CORAL },
    { label: 'センサー車両',   val: '10〜20台\n段階的に拡大',            col: TEAL },
    { label: '廃棄物収集施設', val: '10〜30施設\n市内取引先から選定', col: GREEN },
  ];
  pocs.forEach((p, i) => {
    const x = 0.4 + i * 2.3;
    card(s, x, 0.82, 2.15, 2.1, WHITE, LGREY);
    card(s, x, 0.82, 2.15, 0.06, p.col);
    slab(s, p.label, x, 0.88, 2.15, 0.3, p.col);
    txt(s, p.val, x + 0.1, 1.24, 1.95, 0.95,
      { fontSize: 11.5, bold: true, color: DTXT, align: 'center', lineSpacingMultiple: 1.35 });
  });

  txt(s, '実証実験において定量的に検証するKPI', 0.4, 3.08, 8, 0.36,
    { fontSize: 12.5, bold: true, color: DTXT });

  const kpis = [
    { col: TEAL,  n: 'KPI 1', text: '道路損傷の異常検知精度（検知率・誤検知率・平均検知所要時間）' },
    { col: CORAL, n: 'KPI 2', text: '道路管理工数の変化（巡回・点検・対応処理 前年度比）' },
    { col: GREEN, n: 'KPI 3', text: '廃棄物排出量データの精度・カバレッジ（施設数・品目網羅率）' },
  ];
  kpis.forEach((kpi, i) => {
    const y = 3.54 + i * 0.58;
    card(s, 0.4, y + 0.04, 0.62, 0.34, kpi.col);
    txt(s, kpi.n, 0.4, y + 0.04, 0.62, 0.34,
      { fontSize: 9, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    txt(s, kpi.text, 1.16, y + 0.05, 8.44, 0.34,
      { fontSize: 12, color: DTXT });
    if (i < 2) hrule(s, 0.4, y + 0.48, 9.2);
  });

  s.addNotes('PoC設計。4つのパラメータ：対象エリア・実施期間・センサー車両・廃棄物収集施設。3つのKPI：走行センサー（検知精度・工数）、廃棄物データ（精度・カバレッジ）。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 11: 想定される効果
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  pgTitle(s, '川崎市の都市管理に想定される変化（PoC検証予定）', 9);

  const outcomes = [
    { n: 1, col: CORAL, title: '道路インフラの予防保全',
      bef: '市民通報が主な情報源',
      aft: '異常の予兆を早期に自動検知' },
    { n: 2, col: GREEN,  title: '巡回・監視コストの削減',
      bef: '人的パトロール中心の運用',
      aft: '自動収集が補完し\n職員は対応業務に集中' },
    { n: 3, col: TEAL,  title: '廃棄物・環境KPIの高度化',
      bef: '廃棄物量の把握が難しく\n施策効果を定量化困難',
      aft: '施設別データで\n環境KPIを定量的に検証' },
  ];

  outcomes.forEach((oc, i) => {
    const x = 0.4 + i * 3.2;
    card(s, x, 0.82, 2.95, 4.58, OFFWH, LGREY);
    card(s, x, 0.82, 2.95, 0.06, oc.col);
    numCircle(s, oc.n, x + 1.18, 0.96, oc.col);
    txt(s, oc.title, x + 0.15, 1.56, 2.65, 0.46,
      { fontSize: 13, bold: true, color: oc.col, align: 'center' });
    hrule(s, x + 0.2, 2.06, 2.55);

    card(s, x + 0.15, 2.14, 2.65, 0.24, LGREY);
    txt(s, '現状', x + 0.15, 2.14, 2.65, 0.24,
      { fontSize: 8.5, color: MTXT, align: 'center', valign: 'middle', bold: true, charSpacing: 1 });
    txt(s, oc.bef, x + 0.2, 2.42, 2.55, 0.62,
      { fontSize: 12, color: MTXT, align: 'center', lineSpacingMultiple: 1.3 });

    s.addShape(pres.ShapeType.line, { x: x + 1.18, y: 3.1, w: 0, h: 0.24,
      line: { color: oc.col, width: 1.5 } });
    txt(s, '▼', x + 1.08, 3.28, 0.78, 0.26,
      { fontSize: 12, bold: true, color: oc.col, align: 'center' });

    card(s, x + 0.15, 3.58, 2.65, 0.24, oc.col);
    txt(s, '想定効果', x + 0.15, 3.58, 2.65, 0.24,
      { fontSize: 8.5, color: WHITE, align: 'center', valign: 'middle', bold: true, charSpacing: 1 });
    txt(s, oc.aft, x + 0.2, 3.86, 2.55, 0.62,
      { fontSize: 12, bold: true, color: oc.col, align: 'center', lineSpacingMultiple: 1.3 });
  });

  hrule(s, 0.4, 5.28, 9.2);
  txt(s, '※ 上記は想定効果であり、実際の効果はPoC実施後の検証データをもとに確認します。', 0.4, 5.32, 9.2, 0.28,
    { fontSize: 10, color: MUTED });

  s.addNotes('想定される効果（PoC検証予定）。3つの変化：①道路インフラ予防保全、②巡回コスト削減、③廃棄物・環境KPIの高度化。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 12: 今後の進め方（クロージング）
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);
  card(s, 0, 0, 0.08, 5.625, CORAL);
  card(s, 0.08, 0, 9.92, 0.9, CORAL);
  card(s, 0.08, 4.85, 9.92, 0.775, OFFWH);

  txt(s, '今後の進め方', 0.32, 0.14, 9, 0.3,
    { fontSize: 11, color: WHITE });
  txt(s, 'まずは、ご担当者様とお話しさせてください。', 0.32, 0.98, 9, 0.72,
    { fontSize: 22, bold: true, color: DTXT });

  const steps = [
    { step: '第1段階', title: '課題ヒアリング',
      body: '廃棄物管理・道路インフラそれぞれの具体的な課題・優先エリア・予算感・庁内調整の要件をお聞かせください。実証設計に反映します。' },
    { step: '第2段階', title: 'PoC設計・合意形成',
      body: '対象エリア・実施期間・評価KPI・費用分担の枠組みを共同で設定。法令対応・情報管理の方針も確認します。' },
    { step: '第3段階', title: 'PoC実施・効果検証',
      body: '実証実験を開始し、定期報告を実施します。検証データをもとに本格展開の可否を判断いただきます。' },
  ];
  steps.forEach((st, i) => {
    const x = 0.4 + i * 3.2;
    card(s, x, 1.86, 2.95, 2.88, WHITE, LGREY);
    card(s, x, 1.86, 2.95, 0.06, CORAL);
    slab(s, st.step, x, 1.92, 2.95, 0.34, CORAL);
    numCircle(s, i + 1, x + 0.18, 2.38, CORAL);
    txt(s, st.title, x + 0.8, 2.42, 2.0, 0.44,
      { fontSize: 13, bold: true, color: DTXT });
    txt(s, st.body, x + 0.15, 2.98, 2.65, 1.62,
      { fontSize: 11, color: MTXT, lineSpacingMultiple: 1.4 });
    if (i < 2) {
      txt(s, '→', x + 2.95, 3.12, 0.25, 0.38,
        { fontSize: 16, bold: true, color: CORAL, align: 'center' });
    }
  });

  txt(s, '株式会社サキュレ　　担当：村田　捷樹', 0.32, 4.86, 9.2, 0.44,
    { fontSize: 13, bold: true, color: CORAL, valign: 'middle' });
  txt(s, 'ご連絡をお待ちしております', 0.32, 4.86, 9.5, 0.44,
    { fontSize: 12, color: MTXT, align: 'right', valign: 'middle' });

  s.addNotes('今後の進め方。3段階：①課題ヒアリング→②PoC設計・合意形成→③PoC実施・効果検証。廃棄物管理・道路インフラ両方の課題を統合的にヒアリング。');
}

// ─── Output ───────────────────────────────────────────────
const OUT = '/home/user/sacure/川崎市_統合提案資料_v2.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('Done:', OUT));
