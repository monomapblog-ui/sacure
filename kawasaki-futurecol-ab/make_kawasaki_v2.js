const pptx = require('pptxgenjs');
const pres = new pptx();
pres.layout = 'LAYOUT_16x9'; // 10" × 5.625"

// ─── Palette ──────────────────────────────────────────────
const NAVY  = '1E3461';
const TEAL  = '00939D';
const TEAL2 = '006B74'; // darker teal for text
const STEEL = '4A6FA5';
const WHITE = 'FFFFFF';
const OFFWH = 'F0F4F8';
const LTBL  = 'E8F0F7';
const DTXT  = '1A2332';
const MTXT  = '546E7A';
const MUTED = 'B0C4CE';
const LGREY = 'D8E4EE';

// ─── Helpers ──────────────────────────────────────────────
function sectionLabel(s, text, x, y) {
  // Japanese section label in a small teal box
  card(s, x, y, text.length * 0.175 + 0.3, 0.26, TEAL);
  s.addText(text, { x: x+0.12, y: y+0.01, w: text.length*0.175+0.08, h: 0.24,
    fontSize: 9, bold: true, color: WHITE, fontFace: 'Calibri', charSpacing: 1 });
}
function title(s, text, x, y, w) {
  s.addText(text, { x, y, w: w||9.2, h: 0.8, fontSize: 24, bold: true,
    color: NAVY, fontFace: 'Calibri' });
}
function hrule(s, x, y, w) {
  s.addShape(pres.ShapeType.line, { x, y, w, h: 0,
    line: { color: MUTED, width: 0.5 } });
}
function card(s, x, y, w, h, fill, lineColor) {
  s.addShape(pres.ShapeType.rect, { x, y, w, h,
    fill: { color: fill||WHITE },
    line: lineColor ? { color: lineColor, width: 0.75 } : { type: 'none' } });
}
function txt(s, text, x, y, w, h, opts) {
  s.addText(text, { x, y, w, h, fontFace: 'Calibri', ...opts });
}
function numCircle(s, n, x, y, col) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: 0.52, h: 0.52,
    fill:{ color: col||TEAL }, line:{ type:'none' } });
  txt(s, String(n), x, y, 0.52, 0.52,
    { fontSize: 15, bold:true, color: WHITE, align:'center', valign:'middle' });
}
function dot(s, x, y, col) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: 0.14, h: 0.14,
    fill:{ color: col||TEAL }, line:{ type:'none' } });
}
function accentBar(s, x, y, h, col) {
  card(s, x, y, 0.06, h, col||TEAL);
}

// ═══════════════════════════════════════════════════════════
// SLIDE 1: 表紙
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, NAVY);
  // Bottom accent
  card(s, 0, 5.08, 10, 0.545, TEAL);
  // Subtle geometric decorations top-right
  [3.0, 2.2, 1.5, 0.9].forEach((r, i) => {
    s.addShape(pres.ShapeType.ellipse, {
      x: 9.6 - r/2, y: -r/2, w: r, h: r,
      fill:{ color: TEAL, transparency: 90 - i*5 },
      line:{ type:'none' }
    });
  });
  // Left accent stripe
  card(s, 0, 0, 0.06, 5.08, TEAL);

  txt(s, 'Kawasaki Future Co-Lab　応募提案書', 0.5, 0.52, 9, 0.3,
    { fontSize: 11, color: TEAL, charSpacing: 0.5 });

  txt(s, '動く都市センサー構想', 0.5, 0.88, 9.2, 1.1,
    { fontSize: 44, bold:true, color: WHITE });

  txt(s, '川崎市の配送ネットワークを活用した\n道路インフラ維持管理・交通データ収集の提案', 0.5, 2.08, 9, 0.95,
    { fontSize: 16, color: 'B8CCE0', lineSpacingMultiple: 1.4 });

  s.addShape(pres.ShapeType.line, { x:0.5, y:3.18, w:3.2, h:0,
    line:{ color: TEAL, width: 1.2 } });

  txt(s, '株式会社サキュレ', 0.5, 3.35, 5, 0.42,
    { fontSize: 15, bold:true, color: WHITE });
  txt(s, '担当：村田　捷樹', 0.5, 3.77, 5, 0.35,
    { fontSize: 13, color: 'A0BACE' });
  txt(s, '2026年7月', 0.5, 4.12, 3, 0.3,
    { fontSize: 12, color: '7090A8' });

  txt(s, '道路インフラ維持管理　×　データ活用　×　川崎市との共創', 0, 5.08, 10, 0.545,
    { fontSize: 11.5, bold:true, color: WHITE, align:'center', valign:'middle' });

  s.addNotes('表紙。「動く都市センサー構想」—既存の配送トラック200台を活用した川崎市向けスマートシティソリューション。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 2: 株式会社サキュレについて
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, WHITE);
  sectionLabel(s, '会社概要', 0.5, 0.22);
  title(s, '株式会社サキュレが提供できる価値', 0.5, 0.5, 9.2);
  hrule(s, 0.5, 1.22, 9.2);

  // Left: company panel
  card(s, 0.5, 1.38, 3.1, 4.0, NAVY);
  txt(s, '株式会社サキュレ', 0.65, 1.52, 2.8, 0.5,
    { fontSize: 15, bold:true, color: WHITE });
  txt(s, '廃棄物収集・運搬・中間処理事業\n運送事業', 0.65, 2.02, 2.8, 0.65,
    { fontSize: 11.5, color: TEAL, lineSpacingMultiple: 1.35 });
  hrule(s, 0.65, 2.72, 2.65);

  const stats = [
    { val:'200台以上', label:'保有車両台数' },
    { val:'250社超',  label:'法人取引実績' },
    { val:'東京・神奈川', label:'主要営業エリア', small: true },
  ];
  stats.forEach((st, i) => {
    const y = 2.88 + i * 0.82;
    txt(s, st.val, 0.65, y, 2.8, 0.46,
      { fontSize: st.small ? 14 : 18, bold:true, color: TEAL });
    txt(s, st.label, 0.65, y + (st.small ? 0.26 : 0.32), 2.8, 0.28,
      { fontSize: 10.5, color: 'A0BDD0' });
  });

  // Right: 3 value strips (no emoji, colored shapes)
  const pillars = [
    { kanji:'網', col: NAVY,
      title:'地域密着の配送ネットワーク',
      body:'東京・神奈川を日常的にカバーする200台以上の車両網。既存インフラをそのままセンサーネットワークへ転用できます。' },
    { kanji:'循', col: TEAL,
      title:'循環型経済の担い手',
      body:'廃棄物収集・リファービッシュ・再流通のサイクルを一貫運営。持続可能な都市管理との親和性が高い事業基盤です。' },
    { kanji:'信', col: STEEL,
      title:'行政・企業との協業実績',
      body:'廃棄物収集業務を通じた自治体・企業との継続的な信頼関係。コンプライアンスを重視した誠実な事業運営を行っています。' },
  ];
  pillars.forEach((p, i) => {
    const x = 3.78;
    const y = 1.38 + i * 1.38;
    card(s, x, y, 5.72, 1.28, OFFWH, LGREY);
    card(s, x, y, 5.72, 0.06, p.col);
    // Icon circle with kanji
    s.addShape(pres.ShapeType.ellipse, { x: x+0.15, y: y+0.32, w: 0.6, h: 0.6,
      fill:{ color: p.col }, line:{ type:'none' } });
    txt(s, p.kanji, x+0.15, y+0.32, 0.6, 0.6,
      { fontSize: 16, bold:true, color: WHITE, align:'center', valign:'middle' });
    txt(s, p.title, x+0.88, y+0.1, 4.7, 0.4,
      { fontSize: 13, bold:true, color: NAVY });
    txt(s, p.body, x+0.88, y+0.52, 4.7, 0.7,
      { fontSize: 11, color: MTXT });
  });

  s.addNotes('会社紹介。3つの強み：①地域密着の配送網（200台以上）、②循環型経済の担い手、③行政・企業との協業実績。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 3: 提案の要点
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, OFFWH);
  sectionLabel(s, '提案概要', 0.5, 0.22);
  title(s, '既存の配送網を都市インフラの目に変える', 0.5, 0.5, 9.2);
  hrule(s, 0.5, 1.22, 9.2);

  // 3 stat tiles
  const stats = [
    { val:'200台\n以上', sub:'の配送トラックが\n市内を日常的に走行', col: NAVY },
    { val:'追加設備\n投資不要', sub:'既存車両に搭載するだけ\n新たな調査車両は不要', col: TEAL },
    { val:'3分野', sub:'のデータを\nリアルタイムで収集', col: STEEL },
  ];
  stats.forEach((st, i) => {
    const x = 0.45 + i * 3.2;
    card(s, x, 1.38, 2.95, 2.68, WHITE, LGREY);
    card(s, x, 1.38, 2.95, 0.06, st.col);
    txt(s, st.val, x+0.15, 1.52, 2.65, 0.95,
      { fontSize: 28, bold:true, color: st.col,
        align:'center', valign:'middle', lineSpacingMultiple: 1.15 });
    txt(s, st.sub, x+0.15, 2.52, 2.65, 0.95,
      { fontSize: 12, color: MTXT, align:'center', lineSpacingMultiple: 1.35 });
  });

  // Summary block
  card(s, 0.5, 4.2, 9.2, 1.18, NAVY);
  accentBar(s, 0.5, 4.2, 1.18, TEAL);
  txt(s, '株式会社サキュレが保有する200台以上の配送トラックにカメラとエッジAIを搭載し、日常の配送業務と同時に道路陥没・街路樹の状態・渋滞状況のデータを自動収集します。専用調査車両を追加することなく、川崎市の道路維持管理業務の効率化と予防保全体制の構築に貢献することを目指します。', 0.75, 4.2, 8.9, 1.18,
    { fontSize: 12, color: WHITE, valign:'middle', lineSpacingMultiple: 1.4 });

  s.addNotes('提案の要点。3つのポイント：200台以上の車両・追加設備投資不要・3分野のデータ収集。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 4: 川崎市が直面する課題
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, WHITE);
  sectionLabel(s, '課題の認識', 0.5, 0.22);
  title(s, '道路・インフラ管理における現状の課題', 0.5, 0.5, 9.2);
  hrule(s, 0.5, 1.22, 9.2);

  const issues = [
    { n:1, col: NAVY,
      title:'パトロール体制の維持が困難',
      body:'自治体職員の減少と高齢化が進むなか、広大なエリアを高頻度でカバーする人的パトロールの継続は、人員・予算の両面から困難な状況にあります。' },
    { n:2, col: TEAL,
      title:'市民通報を主とする「事後対応」からの脱却',
      body:'道路陥没・落下物等のインフラ損傷は、市民からの通報を受けて初めて対応するケースが多く、損傷の早期発見・予防的な管理体制の構築が喫緊の課題です。' },
    { n:3, col: STEEL,
      title:'施策立案に必要なデータの不足',
      body:'街路樹の状態・路上駐車の実態・渋滞の発生状況など、都市の現状を定量的に把握するためのデータが乏しく、優先度付けや根拠に基づく施策立案が難しい状況です。' },
  ];

  issues.forEach((iss, i) => {
    const y = 1.38 + i * 1.38;
    accentBar(s, 0.5, y+0.05, 1.1, iss.col);
    numCircle(s, iss.n, 0.7, y+0.05, iss.col);
    txt(s, iss.title, 1.35, y+0.06, 8.25, 0.42,
      { fontSize: 15, bold:true, color: NAVY });
    txt(s, iss.body, 1.35, y+0.52, 8.25, 0.72,
      { fontSize: 12, color: MTXT, lineSpacingMultiple: 1.35 });
  });

  s.addNotes('3つの課題：①パトロール体制の困難、②事後対応型管理、③データ不足。これらを解決するのが動く都市センサー構想。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 5: ソリューションの仕組み
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, OFFWH);
  sectionLabel(s, '仕組み', 0.5, 0.22);
  title(s, '「動く都市センサー」のデータフロー', 0.5, 0.5, 9.2);
  hrule(s, 0.5, 1.22, 9.2);

  // Left explanation
  txt(s, '配送を\n都市観測に変える', 0.5, 1.35, 4.0, 0.88,
    { fontSize: 20, bold:true, color: NAVY, lineSpacingMultiple: 1.25 });
  txt(s, '日常の配送ルートと並行して、\n搭載カメラとエッジAIが\n自動的に市内の状況を記録。\nマスキング処理済みのデータのみを\nクラウド経由でダッシュボードへ提供します。', 0.5, 2.28, 4.0, 1.5,
    { fontSize: 12.5, color: DTXT, lineSpacingMultiple: 1.45 });

  // Key benefit box
  card(s, 0.5, 3.88, 4.0, 1.52, NAVY);
  accentBar(s, 0.5, 3.88, 1.52, TEAL);
  txt(s, '追加の車両・ルート・人員は不要\n配送業務の「ついで」に都市データを収集', 0.75, 3.88, 3.75, 1.52,
    { fontSize: 13, bold:true, color: WHITE, valign:'middle', lineSpacingMultiple: 1.5 });

  // Flow steps (right side) - no emoji, use numbered circles
  const steps = [
    { n:1, label:'配送トラック走行',  sub:'搭載カメラで撮影・記録' },
    { n:2, label:'エッジAI処理',      sub:'個人情報を車内でマスキング' },
    { n:3, label:'クラウドへ送信',    sub:'処理済みデータのみ転送' },
    { n:4, label:'行政ダッシュボード', sub:'地図上で可視化・アラート通知' },
  ];
  steps.forEach((st, i) => {
    const x = 4.78;
    const y = 1.35 + i * 1.05;
    numCircle(s, st.n, x, y+0.05, TEAL);
    if (i < 3) {
      s.addShape(pres.ShapeType.line, { x: x+0.18, y: y+0.6, w: 0, h: 0.42,
        line:{ color: MUTED, width: 1.2 } });
    }
    txt(s, st.label, x+0.65, y+0.05, 2.5, 0.38,
      { fontSize: 14, bold:true, color: NAVY });
    txt(s, st.sub, x+0.65, y+0.45, 2.5, 0.3,
      { fontSize: 11, color: MTXT });
    hrule(s, x, y+0.95, 4.9);
  });

  // Right caption panel
  card(s, 8.0, 1.35, 1.68, 4.05, NAVY);
  txt(s, '配\n送\nと\nデ\nー\nタ\n収\n集\nを\n同\n時\n実\n現', 8.0, 1.55, 1.68, 3.7,
    { fontSize: 12, bold:true, color: WHITE, align:'center', valign:'middle', charSpacing: 1 });

  s.addNotes('データフロー：①走行→②エッジAI処理→③クラウド送信→④ダッシュボード。既存オペレーションを活用。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 6: 3つの優位性
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, WHITE);
  sectionLabel(s, '優位性', 0.5, 0.22);
  title(s, '200台ネットワークがもたらす3つの優位性', 0.5, 0.5, 9.2);
  hrule(s, 0.5, 1.22, 9.2);

  const strengths = [
    { n:1, col: NAVY, title:'網羅性と継続性',
      points:['日常的・高頻度でカバー','体系的な巡回収集','長期的なデータ蓄積'] },
    { n:2, col: TEAL, title:'追加投資不要の経済性',
      points:['専用車両の追加購入不要','配送業務に付随して収集','運用コストを大幅に抑制'] },
    { n:3, col: STEEL, title:'現場知と即時対応力',
      points:['損傷を担当部署へ即通報','配送エリアを熟知した人材','現地の迅速・的確な把握'] },
  ];

  strengths.forEach((st, i) => {
    const x = 0.45 + i * 3.2;
    card(s, x, 1.38, 2.95, 4.0, OFFWH, LGREY);
    card(s, x, 1.38, 2.95, 0.06, st.col);
    // Large number
    txt(s, String(st.n), x, 1.5, 2.95, 0.95,
      { fontSize: 52, bold:true, color: st.col, align:'center', transparency: 12 });
    txt(s, st.title, x+0.15, 2.46, 2.65, 0.52,
      { fontSize: 15, bold:true, color: st.col, align:'center' });
    hrule(s, x+0.2, 3.02, 2.55);
    st.points.forEach((pt, j) => {
      dot(s, x+0.18, 3.17 + j*0.58 + 0.02, st.col);
      txt(s, pt, x+0.38, 3.14 + j*0.58, 2.45, 0.5,
        { fontSize: 12, color: DTXT });
    });
  });

  s.addNotes('3つの優位性：①網羅性と継続性、②追加投資不要の経済性、③現場知と即時対応力。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 7: 収集するデータと活用イメージ
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, WHITE);
  sectionLabel(s, 'データ内容', 0.5, 0.22);
  title(s, 'トラックが走るだけで3つのデータが集まる', 0.5, 0.5, 9.2);
  hrule(s, 0.5, 1.22, 9.2);

  // Left: Truck source panel
  card(s, 0.5, 1.38, 3.1, 4.0, NAVY);
  // Abstract truck shape: large rectangle body
  card(s, 0.72, 2.05, 2.65, 1.5, TEAL);
  card(s, 0.72, 3.1, 2.65, 0.38, '006B74');
  // Wheels
  s.addShape(pres.ShapeType.ellipse, { x: 0.88, y: 3.42, w: 0.45, h: 0.45,
    fill:{ color: DTXT }, line:{ type:'none' } });
  s.addShape(pres.ShapeType.ellipse, { x: 2.62, y: 3.42, w: 0.45, h: 0.45,
    fill:{ color: DTXT }, line:{ type:'none' } });
  // Camera indicator on truck
  card(s, 1.6, 1.82, 0.88, 0.28, '00B4D8');
  txt(s, 'カメラ搭載', 1.6, 1.82, 0.88, 0.28,
    { fontSize: 8, bold:true, color: WHITE, align:'center', valign:'middle' });
  // Label
  txt(s, '配送トラック', 0.5, 3.95, 3.1, 0.32,
    { fontSize: 13, bold:true, color: WHITE, align:'center' });
  txt(s, '走行しながら自動収集', 0.5, 4.27, 3.1, 0.28,
    { fontSize: 10.5, color: TEAL, align:'center', charSpacing: 0.5 });

  // Connector arrows
  const arrowYs = [1.78, 3.0, 4.22];
  arrowYs.forEach(ay => {
    s.addShape(pres.ShapeType.line, { x: 3.62, y: ay + 0.26, w: 0.55, h: 0,
      line:{ color: TEAL, width: 1.5 } });
    card(s, 4.14, ay + 0.14, 0.22, 0.22, TEAL);
    txt(s, '▶', 4.14, ay + 0.14, 0.22, 0.22,
      { fontSize: 9, bold:true, color: WHITE, align:'center', valign:'middle' });
  });

  // Right: 3 data categories (no emoji - use kanji circle)
  const cats = [
    { kanji:'路', col: NAVY,
      name:'道路インフラ',
      kw:'路面陥没・ひび割れ・落下物・補修優先度マップ' },
    { kanji:'緑', col: TEAL,
      name:'都市景観・緑地',
      kw:'街路樹の状態・看板視認性・メンテナンス要否' },
    { kanji:'交', col: STEEL,
      name:'交通・モビリティ',
      kw:'渋滞区間・路上駐車・交通施策の効果測定' },
  ];
  cats.forEach((cat, i) => {
    const ry = 1.52 + i * 1.35;
    if (i > 0) hrule(s, 4.42, ry - 0.12, 5.28);
    // Kanji circle
    s.addShape(pres.ShapeType.ellipse, { x: 4.42, y: ry, w: 0.65, h: 0.65,
      fill:{ color: cat.col }, line:{ type:'none' } });
    txt(s, cat.kanji, 4.42, ry, 0.65, 0.65,
      { fontSize: 20, bold:true, color: WHITE, align:'center', valign:'middle' });
    txt(s, cat.name, 5.18, ry + 0.03, 4.5, 0.42,
      { fontSize: 17, bold:true, color: cat.col });
    txt(s, cat.kw, 5.18, ry + 0.5, 4.5, 0.6,
      { fontSize: 12.5, color: MTXT });
  });

  s.addNotes('3分野の収集データ：道路インフラ・都市景観緑地・交通モビリティ。トラックが走るだけで自動収集。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 8: 個人情報・法令対応の方針
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, WHITE);
  sectionLabel(s, '法令・コンプライアンス', 0.5, 0.22);
  title(s, '個人情報・プライバシーへの対応方針', 0.5, 0.5, 9.2);
  hrule(s, 0.5, 1.22, 9.2);

  // Left: policy statement panel
  card(s, 0.5, 1.38, 3.6, 4.05, NAVY);
  accentBar(s, 0.5, 1.38, 4.05, TEAL);
  txt(s, '個人情報を\n取得しない設計', 0.75, 1.58, 3.25, 0.82,
    { fontSize: 18, bold:true, color: WHITE, lineSpacingMultiple: 1.3 });
  txt(s, '本システムは道路インフラ管理に特化した設計のもと、歩行者・車両の個人識別情報を一切保持しない運用体制を採ります。', 0.75, 2.45, 3.25, 1.0,
    { fontSize: 11.5, color: 'C8DCE8', lineSpacingMultiple: 1.4 });

  // Law references
  const laws = [
    '個人情報保護法 第17条（適正な取得）',
    '個人情報保護法 第27条（第三者提供の制限）',
    '川崎市個人情報保護条例 への準拠',
  ];
  txt(s, '準拠する主な法令', 0.75, 3.55, 3.2, 0.32,
    { fontSize: 10, color: TEAL, charSpacing: 0.5, bold:true });
  laws.forEach((law, i) => {
    dot(s, 0.75, 3.94 + i * 0.36 + 0.05, TEAL);
    txt(s, law, 1.0, 3.9 + i * 0.36, 3.05, 0.34,
      { fontSize: 10, color: MUTED });
  });

  // Right: 3 compliance items
  const items = [
    { n:1, col: NAVY, title:'エッジAIによる自動マスキング',
      body:'歩行者の顔・車両ナンバープレートをカメラ搭載のエッジAIがリアルタイムで自動処理。個人を識別できる情報は収集・保存しません。' },
    { n:2, col: TEAL, title:'データの暗号化と厳格なアクセス管理',
      body:'収集データは暗号化のうえクラウド管理。アクセス権限は最小限に設定し、定期的なセキュリティ監査を実施します。' },
    { n:3, col: STEEL, title:'行政との協議による運用体制の整備',
      body:'川崎市情報セキュリティ担当部署と協議のうえ、データ管理規程・利用目的・保存期間・消去プロセスを明文化します。' },
  ];
  items.forEach((item, i) => {
    const y = 1.38 + i * 1.34;
    numCircle(s, item.n, 4.38, y, item.col);
    txt(s, item.title, 5.0, y+0.05, 4.65, 0.4,
      { fontSize: 14, bold:true, color: NAVY });
    txt(s, item.body, 5.0, y+0.48, 4.65, 0.78,
      { fontSize: 11.5, color: MTXT, lineSpacingMultiple: 1.35 });
    if (i < 2) hrule(s, 4.38, y+1.27, 5.28);
  });

  s.addNotes('プライバシー・法令対応。3つの対策：①エッジAIによる自動マスキング、②暗号化・アクセス管理、③行政との協議による運用体制整備。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 9: 実証実験（PoC）の提案
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, OFFWH);
  sectionLabel(s, '実証実験の設計', 0.5, 0.22);
  title(s, '小さく始め、効果を定量的に検証する', 0.5, 0.5, 9.2);
  hrule(s, 0.5, 1.22, 9.2);

  // 3 parameter boxes
  const pocs = [
    { label:'対象エリア', val:'川崎区・中原区\n幹線道路を優先指定' },
    { label:'実施期間',   val:'第1フェーズ　3ヶ月\n第2フェーズ　3〜6ヶ月' },
    { label:'参加車両',   val:'10〜20台\n段階的に拡大' },
  ];
  pocs.forEach((p, i) => {
    const x = 0.45 + i * 3.2;
    card(s, x, 1.38, 2.95, 1.95, WHITE, LGREY);
    card(s, x, 1.38, 2.95, 0.06, TEAL);
    // Label box
    card(s, x+0.15, 1.52, 2.65, 0.28, LTBL);
    txt(s, p.label, x+0.15, 1.52, 2.65, 0.28,
      { fontSize: 10, color: TEAL2, align:'center', bold:true, valign:'middle', charSpacing:1 });
    txt(s, p.val, x+0.15, 1.85, 2.65, 0.88,
      { fontSize: 13, bold:true, color: NAVY, align:'center', lineSpacingMultiple: 1.4 });
  });

  // KPI section
  txt(s, '実証実験において定量的に検証するKPI', 0.5, 3.48, 8, 0.38,
    { fontSize: 13, bold:true, color: TEAL2 });

  const kpis = [
    { n:'KPI 1', text:'道路管理コストの削減額・削減率（対前年度比）' },
    { n:'KPI 2', text:'担当職員の業務工数の変化（巡回・点検・対応処理）' },
    { n:'KPI 3', text:'異常検知の精度（検知率・誤検知率・平均検知所要時間）' },
  ];
  kpis.forEach((kpi, i) => {
    const y = 3.98 + i * 0.52;
    card(s, 0.5, y+0.04, 0.62, 0.34, TEAL);
    txt(s, kpi.n, 0.5, y+0.04, 0.62, 0.34,
      { fontSize: 9, bold:true, color: WHITE, align:'center', valign:'middle' });
    txt(s, kpi.text, 1.25, y+0.05, 8.35, 0.36,
      { fontSize: 12.5, color: DTXT });
    if (i < 2) hrule(s, 0.5, y+0.44, 9.2);
  });

  s.addNotes('PoC設計。対象エリア：川崎区・中原区の幹線道路。2フェーズ制。3つのKPI：コスト削減・工数変化・検知精度。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 10: 導入によって想定される効果
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, WHITE);
  sectionLabel(s, '想定効果', 0.5, 0.22);
  title(s, '川崎市の道路管理に想定される変化（PoC検証予定）', 0.5, 0.5, 9.2);
  hrule(s, 0.5, 1.22, 9.2);

  const outcomes = [
    { n:1, col: NAVY, title:'巡回コストの削減',
      bef:'人的パトロール中心の運用',
      aft:'自動収集が補完\n職員は対応業務に集中' },
    { n:2, col: TEAL, title:'予防保全体制への移行',
      bef:'市民通報が主な情報源',
      aft:'異常の予兆を早期に自動検知' },
    { n:3, col: STEEL, title:'根拠に基づく施策立案',
      bef:'経験・慣例で優先度を判断',
      aft:'データで優先度を可視化' },
  ];

  outcomes.forEach((oc, i) => {
    const x = 0.45 + i * 3.2;
    card(s, x, 1.38, 2.95, 3.82, OFFWH, LGREY);
    card(s, x, 1.38, 2.95, 0.06, oc.col);
    numCircle(s, oc.n, x+1.18, 1.5, oc.col);
    txt(s, oc.title, x+0.15, 2.1, 2.65, 0.46,
      { fontSize: 13, bold:true, color: oc.col, align:'center' });
    hrule(s, x+0.2, 2.6, 2.55);

    // Current state (muted)
    card(s, x+0.15, 2.68, 2.65, 0.22, LTBL);
    txt(s, '現状', x+0.15, 2.68, 2.65, 0.22,
      { fontSize: 8.5, color: MTXT, align:'center', valign:'middle', bold:true, charSpacing:1 });
    txt(s, oc.bef, x+0.2, 2.94, 2.55, 0.62,
      { fontSize: 12.5, color: MTXT, align:'center', lineSpacingMultiple: 1.35 });

    // Arrow
    s.addShape(pres.ShapeType.line, { x: x+1.18, y: 3.6, w: 0, h: 0.24,
      line:{ color: oc.col, width: 1.5 } });
    txt(s, '▼', x+1.08, 3.78, 0.78, 0.28,
      { fontSize: 12, bold:true, color: oc.col, align:'center' });

    // After state (colored)
    card(s, x+0.15, 4.1, 2.65, 0.22, oc.col);
    txt(s, '想定効果', x+0.15, 4.1, 2.65, 0.22,
      { fontSize: 8.5, color: WHITE, align:'center', valign:'middle', bold:true, charSpacing:1 });
    txt(s, oc.aft, x+0.2, 4.36, 2.55, 0.62,
      { fontSize: 12.5, bold:true, color: oc.col, align:'center', lineSpacingMultiple: 1.35 });
  });

  // Bottom note (replaces SDGs bar)
  hrule(s, 0.5, 5.28, 9.2);
  txt(s, '※ 上記は想定効果であり、実際の効果はPoC実施後の検証データをもとに確認します。', 0.5, 5.32, 9.2, 0.28,
    { fontSize: 10, color: MTXT });

  s.addNotes('想定される効果（PoC検証予定）。3つの変化：①巡回コスト削減、②予防保全への移行、③データ根拠の施策立案。SDGsではなく実務的な注記で締める。');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 11: 今後の進め方
// ═══════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  card(s, 0,0,10,5.625, OFFWH);
  card(s, 0, 4.9, 10, 0.725, TEAL);

  sectionLabel(s, '今後の進め方', 0.5, 0.22);
  title(s, 'まずは、ご担当者様とお話しさせてください。', 0.5, 0.5, 9.2);
  hrule(s, 0.5, 1.22, 9.2);

  const steps = [
    { step:'第1段階', title:'課題ヒアリング',
      body:'貴庁の具体的な課題・優先エリア・予算感・庁内調整の要件をお聞かせください。実証設計に反映します。' },
    { step:'第2段階', title:'PoC設計・合意形成',
      body:'対象エリア・実施期間・評価KPI・費用分担の枠組みを共同で設定。法令対応・情報管理の方針も確認します。' },
    { step:'第3段階', title:'PoC実施・効果検証',
      body:'実証実験を開始し、定期報告を実施します。検証データをもとに本格展開の可否を判断いただきます。' },
  ];
  steps.forEach((st, i) => {
    const x = 0.5 + i * 3.15;
    card(s, x, 1.38, 2.9, 3.35, WHITE, LGREY);
    card(s, x, 1.38, 2.9, 0.38, TEAL);
    txt(s, st.step, x, 1.38, 2.9, 0.38,
      { fontSize: 12, bold:true, color: WHITE, align:'center', valign:'middle' });
    numCircle(s, i+1, x+0.18, 1.88, TEAL);
    txt(s, st.title, x+0.82, 1.92, 1.95, 0.46,
      { fontSize: 13, bold:true, color: NAVY });
    txt(s, st.body, x+0.15, 2.52, 2.62, 1.72,
      { fontSize: 11.5, color: MTXT, lineSpacingMultiple: 1.4 });
    if (i < 2) {
      txt(s, '→', x+2.9, 2.78, 0.25, 0.4,
        { fontSize: 16, bold:true, color: TEAL, align:'center' });
    }
  });

  txt(s, '株式会社サキュレ　　担当：村田　捷樹', 0.5, 4.92, 9.2, 0.45,
    { fontSize: 13.5, bold:true, color: WHITE, valign:'middle' });
  txt(s, 'ご連絡をお待ちしております', 0.5, 4.92, 9.2, 0.45,
    { fontSize: 12, color: WHITE, align:'right', valign:'middle' });

  s.addNotes('今後の進め方。3段階：①課題ヒアリング→②PoC設計・合意形成→③PoC実施・効果検証。「第1段階」という日本語表記でAI感を排除。');
}

// ─── Output ───────────────────────────────────────────────
const OUT = '/tmp/claude-0/-home-user-sacure/1db07409-6d1e-5d98-92a7-465fca194ddc/scratchpad/川崎FutureCo-Lab_提案資料_v2.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('Done:', OUT));
