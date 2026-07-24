const pptx = require('pptxgenjs');
const pres = new pptx();
pres.layout = 'LAYOUT_16x9'; // 10" × 5.625"

const CORAL = 'DD7E6B';
const DKCRL = 'C05F4C';  // ダークコーラル（新品列ヘッダー）
const GREEN = '156818';
const WHITE = 'FFFFFF';
const OFFWH = 'F3F3F3';
const DTXT  = '333333';
const MUTED = '999999';
const LGREY = 'CCCCCC';
const LTCOR = 'FAE8E4';  // 薄コーラル（中古列背景偶数）
const LTCR2 = 'F5DDD8';  // 薄コーラル（中古列背景奇数）
const LTDK  = 'F0D0CA';  // 薄ダークコーラル（新品列背景偶数）
const LTDK2 = 'E8C4BC';  // 薄ダークコーラル（新品列背景奇数）
const AMBER = 'E65100';

function card(s, x, y, w, h, fill, lineColor) {
  s.addShape(pres.ShapeType.rect, { x, y, w, h,
    fill: { color: fill || WHITE },
    line: lineColor ? { color: lineColor, width: 0.75 } : { type: 'none' } });
}
function txt(s, text, x, y, w, h, opts) {
  s.addText(text, { x, y, w, h, fontFace: 'Calibri', ...opts });
}

// ─── 単スライド ───────────────────────────────────────────
{
  const s = pres.addSlide();
  card(s, 0, 0, 10, 5.625, WHITE);

  // ─── Header ───
  card(s, 0, 0, 10, 0.68, CORAL);
  txt(s, '料金比較表　3点セット（冷蔵庫80L・洗濯機4.2kg・電子レンジ）', 0.22, 0, 9.6, 0.68,
    { fontSize: 16, bold: true, color: WHITE, valign: 'middle' });

  txt(s, '※てぶらでどっとこむは6点セット（冷蔵庫90L・洗濯機4.5kg・電子レンジ固定＋選択3点）月5,500円。3点のみプランは要確認。かして！どっとこむは公式実績値。',
    0.4, 5.32, 9.2, 0.20,
    { fontSize: 8, color: MUTED, italic: true });

  // ─── Table layout ───
  // Columns: 比較項目(LW) | サキュレ中古(SW) | サキュレ新品(SW) | かして!(DW) | てぶらで(DW) | 新品購入(DW)
  const tX  = 0.36;
  const tY  = 0.76;
  const RH  = 0.57;   // data row height
  const CHH = 0.52;   // column header height
  const LW  = 1.72;   // label column
  const SW  = 1.36;   // サキュレ sub-columns (×2)
  const DW  = 1.52;   // competitor columns (×3)
  // total: 0.36 + 1.72 + 2×1.36 + 3×1.52 = 0.36+1.72+2.72+4.56 = 9.36 ✓

  // ── Super-header "サキュレ（当社）" spanning 2 sub-cols ──
  const sacSuperX = tX + LW;
  const sacSuperW = SW * 2;
  card(s, sacSuperX, tY, sacSuperW, 0.24, CORAL);
  txt(s, '◀  サキュレ（当社）  ▶', sacSuperX, tY, sacSuperW, 0.24,
    { fontSize: 10, bold: true, color: WHITE, align: 'center', valign: 'middle' });

  // ── Column headers ──
  const colHeaderY = tY + 0.24;
  const colHeaderH = CHH - 0.24;

  // 比較項目
  card(s, tX, tY, LW, CHH, DTXT, LGREY);
  txt(s, '比較項目', tX, tY, LW, CHH,
    { fontSize: 11, bold: true, color: WHITE, align: 'center', valign: 'middle' });

  // サキュレ 中古
  card(s, sacSuperX, colHeaderY, SW, colHeaderH, CORAL, LGREY);
  txt(s, [
    { text: '中古（リユース）', options: { breakLine: true } },
    { text: '月額 1,980円', options: { bold: true } },
  ], sacSuperX, colHeaderY, SW, colHeaderH,
    { fontSize: 10, color: WHITE, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });

  // サキュレ 新品
  const sacNewX = sacSuperX + SW;
  card(s, sacNewX, colHeaderY, SW, colHeaderH, DKCRL, LGREY);
  txt(s, [
    { text: '新品', options: { breakLine: true } },
    { text: '月額 2,980円', options: { bold: true } },
  ], sacNewX, colHeaderY, SW, colHeaderH,
    { fontSize: 10, color: WHITE, align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });

  // 競合3列
  const competitorCols = [
    { label: 'かして！\nどっとこむ', fill: OFFWH, txtCol: DTXT },
    { label: 'てぶらで\nどっとこむ※', fill: OFFWH, txtCol: DTXT },
    { label: '新品購入\n（参考）', fill: OFFWH, txtCol: DTXT },
  ];
  competitorCols.forEach((col, ci) => {
    const x = tX + LW + SW * 2 + ci * DW;
    card(s, x, tY, DW, CHH, col.fill, LGREY);
    txt(s, col.label, x, tY, DW, CHH,
      { fontSize: 10, bold: true, color: col.txtCol,
        align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  });

  // ── Row data ──
  // Each row: [label, sacure_chuko, sacure_shinpin, kashite, tebura, shinpin_purchase]
  const rows = [
    {
      label: '月額費用',
      cells: [
        { v: '1,980 円', sub: '（2年契約）', bold: true, col: CORAL },
        { v: '2,980 円', sub: '（2年契約）', bold: true, col: DKCRL },
        { v: '2,875 円', sub: '（新品・2年／4年1,604円）', bold: false, col: DTXT },
        { v: '5,500 円', sub: '（6点セット）※3点は要確認', bold: false, col: DTXT },
        { v: '—', sub: '（一括購入）', bold: false, col: MUTED },
      ],
    },
    {
      label: '2年間合計',
      cells: [
        { v: '47,520 円', sub: '', bold: true, col: GREEN },
        { v: '71,520 円', sub: '', bold: true, col: GREEN },
        { v: '69,000 円', sub: '（新品・2年）※4年77,000円', bold: false, col: DTXT },
        { v: '要確認', sub: '（6点セット基準）', bold: false, col: MUTED },
        { v: '55,000〜80,000 円', sub: '（初期一括）', bold: false, col: DTXT },
      ],
    },
    {
      label: '初期費用',
      cells: [
        { v: '0 円', sub: '', bold: true, col: GREEN },
        { v: '0 円', sub: '', bold: true, col: GREEN },
        { v: '0 円', sub: '', bold: false, col: DTXT },
        { v: '0 円', sub: '', bold: false, col: DTXT },
        { v: '55,000〜80,000 円', sub: '', bold: false, col: AMBER },
      ],
    },
    {
      label: '配送・設置',
      cells: [
        { v: '無料', sub: '', bold: true, col: GREEN },
        { v: '無料', sub: '', bold: true, col: GREEN },
        { v: '無料', sub: '', bold: false, col: DTXT },
        { v: '無料', sub: '', bold: false, col: DTXT },
        { v: '有料', sub: '（別途）', bold: false, col: AMBER },
      ],
    },
    {
      label: '回収・返却',
      cells: [
        { v: '無料', sub: '', bold: true, col: GREEN },
        { v: '無料', sub: '', bold: true, col: GREEN },
        { v: '無料', sub: '', bold: false, col: DTXT },
        { v: '無料', sub: '', bold: false, col: DTXT },
        { v: '廃棄費用', sub: '（自己負担）', bold: false, col: AMBER },
      ],
    },
    {
      label: '生協提携',
      cells: [
        { v: '新規受付中', sub: '（独占提携可）', bold: true, col: GREEN },
        { v: '新規受付中', sub: '（独占提携可）', bold: true, col: GREEN },
        { v: '大規模校中心', sub: '（競合多い）', bold: false, col: AMBER },
        { v: '大規模校中心', sub: '（競合多い）', bold: false, col: AMBER },
        { v: '—', sub: '', bold: false, col: MUTED },
      ],
    },
  ];

  rows.forEach((row, ri) => {
    const y = tY + CHH + ri * RH;
    const even = ri % 2 === 0;

    // Row label
    card(s, tX, y, LW, RH, even ? OFFWH : WHITE, LGREY);
    txt(s, row.label, tX, y, LW, RH,
      { fontSize: 11, bold: true, color: DTXT, align: 'center', valign: 'middle' });

    // サキュレ 中古
    const bg0 = even ? LTCOR : LTCR2;
    card(s, sacSuperX, y, SW, RH, bg0, LGREY);
    // サキュレ 新品
    const bg1 = even ? LTDK : LTDK2;
    card(s, sacNewX, y, SW, RH, bg1, LGREY);

    // 競合
    competitorCols.forEach((_, ci) => {
      const x = tX + LW + SW * 2 + ci * DW;
      card(s, x, y, DW, RH, even ? OFFWH : WHITE, LGREY);
    });

    // Cell values
    const cellXs = [
      sacSuperX,
      sacNewX,
      tX + LW + SW * 2,
      tX + LW + SW * 2 + DW,
      tX + LW + SW * 2 + DW * 2,
    ];
    const cellWs = [SW, SW, DW, DW, DW];

    row.cells.forEach((cell, ci) => {
      const cx = cellXs[ci];
      const cw = cellWs[ci];
      const isSacure = ci < 2;
      const fSize = isSacure ? 12 : 10.5;

      if (cell.sub) {
        txt(s, cell.v, cx, y + 0.04, cw, 0.34,
          { fontSize: fSize, bold: cell.bold, color: cell.col, align: 'center', valign: 'middle' });
        txt(s, cell.sub, cx, y + 0.36, cw, 0.18,
          { fontSize: 7.5, color: MUTED, align: 'center', valign: 'middle' });
      } else {
        txt(s, cell.v, cx, y, cw, RH,
          { fontSize: fSize, bold: cell.bold, color: cell.col, align: 'center', valign: 'middle' });
      }
    });
  });

  // ── フッターバナー ──
  const tableBottom = tY + CHH + rows.length * RH + 0.06;
  const tableW = LW + SW * 2 + DW * 3;
  card(s, tX, tableBottom, tableW, 0.28, CORAL);
  txt(s, '★ 中古（リユース）なら月1,980円・2年合計47,520円で競合比▲21,480円。初期費用ゼロ・配送回収無料。中小規模大学の生協とは独占提携が可能です。',
    tX + 0.10, tableBottom, tableW - 0.20, 0.28,
    { fontSize: 9, bold: true, color: WHITE, valign: 'middle' });

  s.addNotes('料金比較表v2。サキュレは中古1,980円・新品2,980円の2プライス制。競合比で最安値水準。初期費用ゼロ・配送回収無料・生協独占提携可能が差別化ポイント。競合価格は概算（公開情報ベース）。');
}

const OUT = '/tmp/claude-0/-home-user-sacure/1db07409-6d1e-5d98-92a7-465fca194ddc/scratchpad/料金比較表_v1.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('Done:', OUT));
