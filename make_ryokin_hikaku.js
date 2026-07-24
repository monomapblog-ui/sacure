const pptx = require('pptxgenjs');
const pres = new pptx();
pres.layout = 'LAYOUT_16x9'; // 10" × 5.625"

const CORAL = 'DD7E6B';
const GREEN = '156818';
const LTGRN = 'D9EAD3';
const WHITE = 'FFFFFF';
const OFFWH = 'F3F3F3';
const DTXT  = '333333';
const MTXT  = '595959';
const MUTED = '999999';
const LGREY = 'CCCCCC';
const LTCOR = 'FAE8E4';  // 薄コーラル（サキュレ列背景）
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
  txt(s, '料金比較表　3点セット（冷蔵庫・洗濯機・電子レンジ）', 0.22, 0, 9.6, 0.68,
    { fontSize: 16, bold: true, color: WHITE, valign: 'middle' });

  // ─── 凡例（右上） ───
  txt(s, '※競合価格は公開情報をもとにした概算。正確な金額は各社サイトでご確認ください。',
    0.4, 5.30, 9.2, 0.22,
    { fontSize: 8, color: MUTED, italic: true });

  // ─── Table setup ───
  const tX = 0.38;       // table left edge
  const tY = 0.76;       // table top
  const RH = 0.60;       // row height (data rows)
  const CHH = 0.54;      // column header height
  const LW = 2.10;       // label column width
  const DW = 1.86;       // data column width (×4)
  // total: 2.10 + 4×1.86 = 9.54"... adjust
  // with tX=0.38: right edge = 0.38+2.10+4×1.86 = 0.38+2.10+7.44 = 9.92 ✓

  const cols = [
    { label: 'サキュレ\n（当社）', fill: CORAL, txtCol: WHITE, isSacure: true },
    { label: 'かして！どっとこむ', fill: OFFWH, txtCol: DTXT, isSacure: false },
    { label: 'てぶらで\nどっとこむ', fill: OFFWH, txtCol: DTXT, isSacure: false },
    { label: '新品購入\n（参考）', fill: OFFWH, txtCol: DTXT, isSacure: false },
  ];

  const rows = [
    {
      label: '月額費用',
      vals: [
        { v: '2,083 円', sub: '（2年契約）', bold: true, col: CORAL },
        { v: '3,200〜4,500 円', sub: '（概算）', bold: false, col: DTXT },
        { v: '2,900〜4,200 円', sub: '（概算）', bold: false, col: DTXT },
        { v: '— 円', sub: '（一括購入）', bold: false, col: MUTED },
      ],
    },
    {
      label: '2年間合計',
      vals: [
        { v: '50,000 円', sub: '', bold: true, col: GREEN },
        { v: '76,800〜108,000 円', sub: '', bold: false, col: DTXT },
        { v: '69,600〜100,800 円', sub: '', bold: false, col: DTXT },
        { v: '55,000〜80,000 円', sub: '（初期一括）', bold: false, col: DTXT },
      ],
    },
    {
      label: '初期費用',
      vals: [
        { v: '0 円', sub: '', bold: true, col: GREEN },
        { v: '0 円', sub: '', bold: false, col: DTXT },
        { v: '0 円', sub: '', bold: false, col: DTXT },
        { v: '55,000〜80,000 円', sub: '', bold: false, col: AMBER },
      ],
    },
    {
      label: '配送・設置',
      vals: [
        { v: '無料', sub: '', bold: true, col: GREEN },
        { v: '無料', sub: '', bold: false, col: DTXT },
        { v: '無料', sub: '', bold: false, col: DTXT },
        { v: '有料', sub: '（別途）', bold: false, col: AMBER },
      ],
    },
    {
      label: '回収・返却',
      vals: [
        { v: '無料', sub: '', bold: true, col: GREEN },
        { v: '無料', sub: '', bold: false, col: DTXT },
        { v: '無料', sub: '', bold: false, col: DTXT },
        { v: '廃棄費用', sub: '（自己負担）', bold: false, col: AMBER },
      ],
    },
    {
      label: '生協提携',
      vals: [
        { v: '新規受付中', sub: '（独占提携可）', bold: true, col: GREEN },
        { v: '大規模校中心', sub: '（競合多い）', bold: false, col: AMBER },
        { v: '大規模校中心', sub: '（競合多い）', bold: false, col: AMBER },
        { v: '—', sub: '', bold: false, col: MUTED },
      ],
    },
  ];

  // ── Column headers ──
  cols.forEach((col, ci) => {
    const x = tX + LW + ci * DW;
    card(s, x, tY, DW, CHH, col.fill, LGREY);
    txt(s, col.label, x, tY, DW, CHH,
      { fontSize: col.isSacure ? 12 : 10, bold: true, color: col.txtCol,
        align: 'center', valign: 'middle', lineSpacingMultiple: 1.2 });
  });
  // label column header (empty)
  card(s, tX, tY, LW, CHH, DTXT, LGREY);
  txt(s, '比較項目', tX, tY, LW, CHH,
    { fontSize: 11, bold: true, color: WHITE, align: 'center', valign: 'middle' });

  // ── Data rows ──
  rows.forEach((row, ri) => {
    const y = tY + CHH + ri * RH;
    const isEven = ri % 2 === 0;

    // Row label
    card(s, tX, y, LW, RH, isEven ? OFFWH : WHITE, LGREY);
    txt(s, row.label, tX, y, LW, RH,
      { fontSize: 11.5, bold: true, color: DTXT, align: 'center', valign: 'middle' });

    // Data cells
    row.vals.forEach((val, ci) => {
      const x = tX + LW + ci * DW;
      const isSacure = ci === 0;
      const bgFill = isSacure ? (isEven ? LTCOR : 'F5DDD8') : (isEven ? OFFWH : WHITE);
      card(s, x, y, DW, RH, bgFill, LGREY);

      if (val.sub) {
        // main value + sub note
        txt(s, val.v, x, y + 0.04, DW, 0.36,
          { fontSize: isSacure ? 13 : 11, bold: val.bold, color: val.col,
            align: 'center', valign: 'middle' });
        txt(s, val.sub, x, y + 0.37, DW, 0.20,
          { fontSize: 8, color: MUTED, align: 'center', valign: 'middle' });
      } else {
        txt(s, val.v, x, y, DW, RH,
          { fontSize: isSacure ? 13 : 11, bold: val.bold, color: val.col,
            align: 'center', valign: 'middle' });
      }
    });
  });

  // ── 強調フッター（全幅） ──
  const tableBottom = tY + CHH + rows.length * RH + 0.08;
  card(s, tX, tableBottom, LW + DW * 4, 0.28, CORAL);
  txt(s, '★ サキュレは業界最安値水準（月額2,083円）・初期費用ゼロ・生協独占提携可で、中小規模大学への先行参入に最適です。',
    tX + 0.10, tableBottom, LW + DW * 4 - 0.20, 0.28,
    { fontSize: 9.5, bold: true, color: WHITE, valign: 'middle' });

  s.addNotes('料金比較表。サキュレの3点セット月額2,083円（2年50,000円）は競合比較で最安値水準。初期費用ゼロ・配送回収無料・生協独占提携可能が差別化ポイント。競合価格は概算（公開情報ベース）。');
}

const OUT = '/tmp/claude-0/-home-user-sacure/1db07409-6d1e-5d98-92a7-465fca194ddc/scratchpad/料金比較表_v1.pptx';
pres.writeFile({ fileName: OUT }).then(() => console.log('Done:', OUT));
