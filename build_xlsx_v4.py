"""
採算管理ブック v4 — ルート別・案件別の採算確認に特化（4シート）
"""
import openpyxl, xlrd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.chart import BarChart, Reference

SRC_MASTER = "/root/.claude/uploads/ffa395c4-1170-42eb-9d67-c76992faca3c/83a04c7f-__________________.xlsx"
SRC_2026   = "/root/.claude/uploads/ffa395c4-1170-42eb-9d67-c76992faca3c/912c64bc-___2026.xlsx"
SRC_R76    = "/root/.claude/uploads/ffa395c4-1170-42eb-9d67-c76992faca3c/8a0a54a8-__R7.6.xls"

# ── カラー ─────────────────────────────────────────────────────
NAVY  = "0D2B5E";  BLUE  = "1A5FC8";  GREEN = "2EB872"
GRAY  = "55657A";  WHITE = "FFFFFF";  LGRAY = "F2F6FC"
YEL   = "FFF3CD";  REDLT = "FFE0E0";  GRNLT = "D4EDDA"
BLULT = "D0E8FF";  PURLT = "EDE7F6";  BDR_C = "C5D5E8"

# 案件カラー
CASE = {
    "C001": {"name":"東急グループ案件", "bg":BLULT, "hd":"1565A8"},
    "C002": {"name":"都営地下鉄案件",   "bg":PURLT, "hd":"5B2D8E"},
    "C003": {"name":"一般事業者案件",   "bg":GRNLT, "hd":"1A6B2E"},
}

# ルート定義
ROUTES = [
    ("612-1","運送（本社）",   "NAVY"),
    ("612-2","入札（公共）",   "NAVY"),
    ("613-1","車輌",           "BLUE"),
    ("613-2","KC",             "BLUE"),
    ("614-1","コンテナ",       "GREEN"),
    ("614-2","ENDO",           "GREEN"),
    ("614-3","現金",           "GREEN"),
    ("615-1","板橋",           "NAVY"),
    ("615-2","新横浜",         "NAVY"),
    ("615-3","現金2",          "BLUE"),
    ("615-4","有価物",         "GREEN"),
]
CODES = [r[0] for r in ROUTES]

# 配分設定 (C001%, C002%, C003%)
MAPPING = {
    "612-1": (0,   0, 100),
    "612-2": (0, 100,   0),
    "613-1": (20,  10,  70),
    "613-2": (0,   0,   0),
    "614-1": (0,   0, 100),
    "614-2": (0,   0, 100),
    "614-3": (0,   0, 100),
    "615-1": (60,  0,  40),
    "615-2": (0,   0, 100),
    "615-3": (0,   0, 100),
    "615-4": (0,   0, 100),
}

def fill(h): return PatternFill("solid", fgColor=h)
def f(bold=False, sz=11, color=NAVY, name="Yu Gothic UI"):
    return Font(bold=bold, size=sz, color=color, name=name)
def a(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def b():
    s = Side(style="thin", color=BDR_C)
    return Border(top=s, bottom=s, left=s, right=s)

def c(ws, row, col, val=None, bold=False, sz=11, bg=None, fc=NAVY,
      ha="left", va="center", wrap=False, fmt=None):
    cell = ws.cell(row=row, column=col, value=val)
    cell.font = f(bold=bold, sz=sz, color=fc)
    cell.alignment = a(h=ha, v=va, wrap=wrap)
    cell.border = b()
    if bg:  cell.fill = fill(bg)
    if fmt: cell.number_format = fmt
    return cell

# ────────────────────────────────────────────────────────────────
#  データ取得
# ────────────────────────────────────────────────────────────────
def reiwa(s):
    s = str(s).strip()
    if not s.startswith("R"): return None
    try:
        era, m = s[1:].split(".")
        return f"{int(era)+2018}-{int(m):02d}"
    except: return None

def load_sales():
    sales = {}

    def parse_summary(rows, header):
        for row in rows:
            if not row or not row[0] or not str(row[0]).startswith("R"): continue
            ym = reiwa(str(row[0]))
            if not ym: continue
            d = {}
            for code in CODES:
                for ci, h in enumerate(header):
                    if str(h).strip() == code and ci < len(row):
                        try: d[code] = float(row[ci]) if row[ci] else 0.0
                        except: d[code] = 0.0
                        break
            for ci, h in enumerate(header):
                if "売" in str(h) and "上" in str(h) and ci < len(row):
                    try: d["_total"] = float(row[ci]) if row[ci] else 0.0
                    except: d["_total"] = 0.0
                    break
            if any(v > 0 for v in d.values()):
                sales[ym] = d

    def parse_month_header(header):
        col_ym = {}
        cur_era = None
        for ci, h in enumerate(header):
            if not h: continue
            h = str(h).strip()
            if h.startswith("R") and "." in h:
                try:
                    era_s, m_s = h[1:].split(".")
                    cur_era = int(era_s)
                    col_ym[ci] = f"{cur_era+2018}-{int(float(m_s)):02d}"
                except: pass
            elif cur_era is not None:
                try:
                    m = int(float(h))
                    if 1 <= m <= 12:
                        col_ym[ci] = f"{cur_era+2018}-{m:02d}"
                except: pass
        return col_ym

    # R7.6.xls: R6.7〜R7.6
    wb = xlrd.open_workbook(SRC_R76)
    ws = wb.sheet_by_name("売上")
    hdr = [str(ws.cell_value(2, c)).strip() for c in range(ws.ncols)]
    rows = [[ws.cell_value(r, c) for c in range(ws.ncols)] for r in range(3, ws.nrows)]
    parse_summary(rows, hdr)

    # 2026.xlsx 売上シート: R7.7, R7.8（全11ルート確定値）
    wb2 = openpyxl.load_workbook(SRC_2026, data_only=True)
    ws2 = wb2["売上"]
    hdr2 = [str(v).strip() if v else "" for v in next(ws2.iter_rows(min_row=3, max_row=3, values_only=True))]
    rows2 = [list(row) for row in ws2.iter_rows(min_row=4, values_only=True)]
    parse_summary(rows2, hdr2)

    # 2026.xlsx 各拠点シート合計行: R7.9以降（調整ファイルから取得）
    SHEET_CODE = {
        "本社": "612-1", "公共": "612-2", "遠藤": "614-2",
        "板橋": "615-1", "新横浜": "615-2", "有価物": "615-4",
    }
    for sheet_name, code in SHEET_CODE.items():
        ws_i = wb2[sheet_name]
        all_rows = list(ws_i.iter_rows(values_only=True))

        # ヘッダー行を特定
        hdr_idx = None
        for ri, row in enumerate(all_rows):
            if any(v and str(v).strip().startswith("R") and "." in str(v) for v in row):
                hdr_idx = ri; break
        if hdr_idx is None: continue

        col_ym = parse_month_header(all_rows[hdr_idx])
        r77_col = next((ci for ci, ym in col_ym.items() if ym == "2025-07"), None)
        if r77_col is None: continue

        # R7.7の既知合計値で合計行を特定（ダブルカウント回避）
        anchor = sales.get("2025-07", {}).get(code, 0)
        if anchor == 0: continue

        total_row = None
        for row in all_rows:
            if r77_col < len(row) and row[r77_col]:
                try:
                    if abs(float(row[r77_col]) - anchor) < 10:
                        total_row = list(row); break
                except: pass
        if total_row is None: continue

        # R7.9以降の値を追加
        for ci, ym in col_ym.items():
            if ym <= "2025-08": continue  # 売上シートから取得済
            if ci >= len(total_row) or not total_row[ci]: continue
            try:
                v = float(total_row[ci])
                if v > 1000:  # ゴミ値除外
                    if ym not in sales:
                        sales[ym] = {}
                    sales[ym][code] = v
            except: pass

    # _total が未計算の月は合算
    for ym, d in sales.items():
        if "_total" not in d:
            d["_total"] = sum(v for k, v in d.items() if k != "_total")

    return {k: v for k, v in sorted(sales.items()) if any(v > 0 for v in v.values())}

# ════════════════════════════════════════════════════════════════
#  SHEET 1: 案件別採算
# ════════════════════════════════════════════════════════════════
def sheet_anken(ws, sales):
    ws.title = "案件別採算"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "C4"

    # タイトル
    ws.merge_cells("A1:T1")
    t = ws["A1"]
    t.value = "案件別 採算管理　（★黄色セル＝入力　緑セル＝自動計算）"
    t.font = f(bold=True, sz=14, color=WHITE)
    t.fill = fill(NAVY); t.alignment = a(h="center"); ws.row_dimensions[1].height = 34

    # 説明
    ws.merge_cells("A2:T2")
    note = ws["A2"]
    note.value = "【売上】⑤売上原票シートをVLOOKUPで参照（D列=集計月 YYYY-MM を入力すると自動セット）　【原価】黄色セルに毎月入力してください"
    note.font = f(sz=9, color=GRAY); note.fill = fill(LGRAY); note.alignment = a(h="left"); note.border = b()
    ws.row_dimensions[2].height = 16

    # カラム幅
    widths = [10, 22, 11,  14, 14, 16,  14, 12, 12, 12, 12, 11, 11, 16,  14, 10, 12,  12, 16, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ヘッダー（2段）
    ws.row_dimensions[3].height = 16
    ws.row_dimensions[4].height = 38
    def merge_hd(r1,c1,r2,c2,txt,bg):
        ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
        cell = ws.cell(row=r1,column=c1,value=txt)
        cell.font=f(bold=True,sz=9,color=WHITE); cell.fill=fill(bg)
        cell.alignment=a(h="center",v="center",wrap=True); cell.border=b()

    merge_hd(3,1,4,1,"案件ID",    NAVY)
    merge_hd(3,2,4,2,"案件名",    NAVY)
    merge_hd(3,3,4,3,"集計月\n(YYYY-MM)", NAVY)
    # 売上
    merge_hd(3,4,3,6,"売上（VLOOKUPで自動参照）","155A00")
    for col, txt in [(4,"運送\n612-1"),(5,"入札\n612-2"),(6,"板橋\n615-1")]:
        merge_hd(4,col,4,col,txt,"155A00")
    merge_hd(3,7,4,7,"その他\n売上合計","155A00")
    merge_hd(3,8,4,8,"売上\n合計","155A00")
    # 原価
    merge_hd(3,9,3,15,"原価　（黄色セルに入力）","7B3F00")
    for col,txt in [(9,"人件費"),(10,"燃料費"),(11,"車両\n維持費"),(12,"中間\n処理費"),
                    (13,"搬入\n運搬費"),(14,"電子M\n代行費"),(15,"管理費")]:
        merge_hd(4,col,4,col,txt,"7B3F00")
    merge_hd(3,16,4,16,"原価\n合計","7B3F00")
    # 採算
    merge_hd(3,17,4,17,"粗利益","1A6B2E")
    merge_hd(3,18,4,18,"粗利率","1A6B2E")
    merge_hd(3,19,4,19,"採算\n判定","1A6B2E")
    merge_hd(3,20,4,20,"前月比\n売上","1A6B2E")

    JPY = "#,##0"
    PCT = "0.0%"
    months = sorted(sales.keys())

    ROW = 5
    prev_row = {}   # cid -> last row index
    for cid, info in CASE.items():
        bg = info["bg"]
        hd = info["hd"]
        name = info["name"]
        pct_col = {"C001":4,"C002":5,"C003":6}[cid]  # マッピング列

        for ym in months:
            ws.row_dimensions[ROW].height = 22

            c(ws,ROW,1,cid,   bold=True,sz=11,bg=bg,ha="center")
            c(ws,ROW,2,name,  sz=11,bg=bg)

            # D列: 集計月（入力キー）
            cx = ws.cell(row=ROW,column=3,value=ym)
            cx.font=f(bold=True,sz=11,color=WHITE); cx.fill=fill(hd)
            cx.alignment=a(h="center"); cx.border=b()

            # 売上（VLOOKUP）
            # E=612-1分, F=612-2分, G=615-1分
            for col, code, src_col in [(4,"612-1",3),(5,"612-2",4),(6,"615-1",10)]:
                vl = ws.cell(row=ROW,column=col)
                vl.value = (f'=IFERROR(VLOOKUP(C{ROW},▼売上原票!$A:$P,{src_col},0)'
                            f'*VLOOKUP("{code}",▼マッピング!$A:$G,{pct_col},0)/100,0)')
                vl.number_format=JPY; vl.font=f(sz=10)
                vl.fill=fill("E8F5E9"); vl.alignment=a(h="right"); vl.border=b()

            # G: その他科目合計
            other_codes = [("613-1",5),("613-2",6),("614-1",7),("614-2",8),("614-3",9),
                           ("615-2",11),("615-3",12),("615-4",13)]
            other_expr = "+".join(
                f'IFERROR(VLOOKUP(C{ROW},▼売上原票!$A:$P,{sc_},0)*VLOOKUP("{cd}",▼マッピング!$A:$G,{pct_col},0)/100,0)'
                for cd, sc_ in other_codes
            )
            g_ = ws.cell(row=ROW,column=7)
            g_.value = f"={other_expr}"
            g_.number_format=JPY; g_.font=f(sz=10)
            g_.fill=fill("E8F5E9"); g_.alignment=a(h="right"); g_.border=b()

            # H: 売上合計
            h_ = ws.cell(row=ROW,column=8)
            h_.value = f"=D{ROW}+E{ROW}+F{ROW}+G{ROW}"
            h_.number_format=JPY; h_.font=f(bold=True,sz=12)
            h_.fill=fill("C8F7D4"); h_.alignment=a(h="right"); h_.border=b()

            # 原価（黄色）空欄
            for ci in range(9, 16):
                c(ws,ROW,ci,None,sz=10,bg=YEL,ha="right",fmt=JPY)

            # P: 原価合計
            p_ = ws.cell(row=ROW,column=16)
            p_.value = f"=I{ROW}+J{ROW}+K{ROW}+L{ROW}+M{ROW}+N{ROW}+O{ROW}"
            p_.number_format=JPY; p_.font=f(bold=True,sz=11)
            p_.fill=fill(YEL); p_.alignment=a(h="right"); p_.border=b()

            # Q: 粗利
            q_ = ws.cell(row=ROW,column=17)
            q_.value = f"=H{ROW}-P{ROW}"
            q_.number_format=JPY; q_.font=f(bold=True,sz=12)
            q_.fill=fill(GRNLT); q_.alignment=a(h="right"); q_.border=b()

            # R: 粗利率
            r_ = ws.cell(row=ROW,column=18)
            r_.value = f"=IF(H{ROW}<>0,Q{ROW}/H{ROW},0)"
            r_.number_format=PCT; r_.font=f(bold=True,sz=12)
            r_.fill=fill(GRNLT); r_.alignment=a(h="center"); r_.border=b()

            # S: 採算判定
            s_ = ws.cell(row=ROW,column=19)
            s_.value = (f'=IF(R{ROW}>=0.25,"◎ 優良",'
                        f'IF(R{ROW}>=0.15,"○ 良好",'
                        f'IF(R{ROW}>=0.05,"△ 要注意","× 赤字")))')
            s_.font=f(bold=True,sz=11); s_.fill=fill(GRNLT)
            s_.alignment=a(h="center"); s_.border=b()

            # T: 前月比
            if cid in prev_row:
                t_ = ws.cell(row=ROW,column=20)
                t_.value = f"=IF(H{prev_row[cid]}<>0,H{ROW}/H{prev_row[cid]}-1,\"\")"
                t_.number_format="+0.0%;-0.0%;0.0%"; t_.font=f(sz=10)
                t_.fill=fill(bg); t_.alignment=a(h="center"); t_.border=b()
            else:
                c(ws,ROW,20,"—",sz=10,bg=bg,ha="center")

            prev_row[cid] = ROW
            ROW += 1

        # 案件区切り線（空行）
        ws.row_dimensions[ROW].height = 8
        for col in range(1,21):
            ws.cell(row=ROW,column=col).fill = fill(NAVY)
        ROW += 1

    # 入力用空白行（将来月）
    for _ in range(15):
        ws.row_dimensions[ROW].height = 22
        bg = LGRAY if ROW%2==0 else WHITE
        c(ws,ROW,1,"",sz=11,bg=bg,ha="center")
        c(ws,ROW,2,"",sz=11,bg=bg)
        c(ws,ROW,3,"",sz=11,bg=bg,ha="center")
        for col in range(4,8):
            cc = ws.cell(row=ROW,column=col)
            cc.value=""
            cc.number_format=JPY; cc.font=f(sz=10)
            cc.fill=fill("E8F5E9"); cc.alignment=a(h="right"); cc.border=b()
        for col in range(9,16):
            c(ws,ROW,col,None,sz=10,bg=YEL,ha="right",fmt=JPY)
        for col,expr,nm in [
            (8,  f"=D{ROW}+E{ROW}+F{ROW}+G{ROW}", JPY),
            (16, f"=IF(I{ROW}+J{ROW}+K{ROW}+L{ROW}+M{ROW}+N{ROW}+O{ROW}=0,\"\",I{ROW}+J{ROW}+K{ROW}+L{ROW}+M{ROW}+N{ROW}+O{ROW})", JPY),
            (17, f"=IF(H{ROW}<>0,H{ROW}-P{ROW},\"\")", JPY),
            (18, f"=IF(H{ROW}<>0,Q{ROW}/H{ROW},\"\")", PCT),
            (19, f'=IF(R{ROW}="","",IF(R{ROW}>=0.25,"◎ 優良",IF(R{ROW}>=0.15,"○ 良好",IF(R{ROW}>=0.05,"△ 要注意","× 赤字"))))', "@"),
        ]:
            cc = ws.cell(row=ROW,column=col)
            cc.value=expr; cc.number_format=nm
            cc.font=f(bold=(col in(8,16,17,18)),sz=11 if col in(8,17,18) else 10)
            cc.fill=fill("E8F5E9" if col in(4,5,6,7) else YEL if 9<=col<=15 else GRNLT if col>=16 else bg)
            cc.alignment=a(h="center" if col==19 else "right"); cc.border=b()
        ROW += 1

    # 条件付き書式
    end = ROW-1
    for expr, fg_c, bg_c in [
        ('S5="◎ 優良"', "145A32", "C8F7D4"),
        ('S5="○ 良好"', "1A6B2E", GRNLT),
        ('S5="△ 要注意"',"856404", YEL),
        ('S5="× 赤字"', "8B0000", REDLT),
    ]:
        ws.conditional_formatting.add(f"S5:S{end}",
            FormulaRule(formula=[expr], fill=fill(bg_c), font=Font(bold=True,color=fg_c)))
    ws.conditional_formatting.add(f"R5:R{end}",
        ColorScaleRule(start_type="num",start_value=0,   start_color="FF6B6B",
                       mid_type="num",  mid_value=0.15,  mid_color="FFD93D",
                       end_type="num",  end_value=0.30,  end_color="6BCB77"))


# ════════════════════════════════════════════════════════════════
#  SHEET 2: ルート別採算
# ════════════════════════════════════════════════════════════════
def sheet_route(ws, sales):
    ws.title = "ルート別採算"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D4"

    ws.merge_cells("A1:S1")
    t = ws["A1"]
    t.value = "ルート別 採算管理　（ルート = 科目コード / 各営業担当）"
    t.font=f(bold=True,sz=14,color=WHITE); t.fill=fill(NAVY)
    t.alignment=a(h="center"); ws.row_dimensions[1].height=34

    ws.merge_cells("A2:S2")
    note=ws["A2"]
    note.value = "【売上】⑤売上原票からVLOOKUPで自動参照　【原価】黄色セルに実績を入力　各ルートの粗利率で採算の良し悪しを確認できます"
    note.font=f(sz=9,color=GRAY); note.fill=fill(LGRAY); note.alignment=a(h="left"); note.border=b()
    ws.row_dimensions[2].height=16

    # カラム幅
    widths = [11,18,11, 16, 14,12,12,12,11,11,14, 14,10,12, 12,16,16]
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width=w

    # ヘッダー
    ws.row_dimensions[3].height=16; ws.row_dimensions[4].height=36
    def mhd(r1,c1,r2,c2,txt,bg):
        ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
        cell=ws.cell(row=r1,column=c1,value=txt)
        cell.font=f(bold=True,sz=9,color=WHITE); cell.fill=fill(bg)
        cell.alignment=a(h="center",v="center",wrap=True); cell.border=b()

    mhd(3,1,4,1,"ルートコード",NAVY); mhd(3,2,4,2,"ルート名",NAVY)
    mhd(3,3,4,3,"集計月\n(YYYY-MM)",NAVY)
    mhd(3,4,4,4,"売上\n(VLOOKUP)","155A00")
    mhd(3,5,3,11,"原価　（黄色セル入力）","7B3F00")
    for col,txt in [(5,"人件費"),(6,"燃料費"),(7,"車両費"),(8,"処理費"),
                    (9,"運搬費"),(10,"その他"),(11,"合計")]:
        mhd(4,col,4,col,txt,"7B3F00")
    mhd(3,12,4,12,"粗利益","1A6B2E")
    mhd(3,13,4,13,"粗利率","1A6B2E")
    mhd(3,14,4,14,"採算判定","1A6B2E")
    mhd(3,15,4,15,"主担当\n案件","1A6B2E")
    mhd(3,16,4,16,"前月比\n売上","1A6B2E")
    mhd(3,17,4,17,"月間回収量\n(kg推算)",GRAY)

    JPY="#,##0"; PCT="0.0%"

    # ルートの色グループ
    route_bg = {
        "612-1":BLULT,"612-2":BLULT,    # 本社系
        "613-1":PURLT,"613-2":PURLT,    # 車輌系
        "614-1":GRNLT,"614-2":GRNLT,"614-3":GRNLT,  # 現場系
        "615-1":BLULT,"615-2":BLULT,    # 営業所系
        "615-3":"FFF9C4","615-4":GRNLT, # 現金・有価物
    }
    # 主担当案件（マッピング設定から）
    route_case = {}
    for code, (p1,p2,p3) in MAPPING.items():
        if p1>=p2 and p1>=p3 and p1>0: route_case[code]="C001東急"
        elif p2>=p1 and p2>=p3 and p2>0: route_case[code]="C002都営"
        elif p3>0: route_case[code]="C003一般"
        else: route_case[code]="別管理"

    # 原価の目安（月次参考値）
    cost_guide = {
        "612-1":(650000,190000,85000,0,75000,45000),
        "612-2":(1800000,520000,240000,1200000,180000,120000),
        "613-1":(180000,55000,180000,0,15000,20000),
        "613-2":(0,0,0,0,0,0),
        "614-1":(85000,25000,15000,0,20000,10000),
        "614-2":(120000,38000,18000,0,25000,12000),
        "614-3":(20000,6000,3000,0,5000,3000),
        "615-1":(850000,248000,118000,580000,95000,62000),
        "615-2":(280000,82000,38000,195000,32000,20000),
        "615-3":(15000,4000,2000,0,3000,2000),
        "615-4":(35000,10000,5000,0,8000,5000),
    }

    months = sorted(sales.keys())
    ROW = 5
    prev_row = {}

    for code, rname, _ in ROUTES:
        bg = route_bg.get(code, LGRAY)
        src_col = 3 + CODES.index(code)   # ⑤売上原票の列番号

        for ym in months:
            ws.row_dimensions[ROW].height=21

            c(ws,ROW,1,code,bold=True,sz=11,bg=bg,ha="center")
            c(ws,ROW,2,rname,sz=11,bg=bg)

            cx=ws.cell(row=ROW,column=3,value=ym)
            cx.font=f(bold=True,sz=10,color=WHITE); cx.fill=fill(NAVY)
            cx.alignment=a(h="center"); cx.border=b()

            # 売上（VLOOKUP from ▼売上原票）
            vl=ws.cell(row=ROW,column=4)
            vl.value=f"=IFERROR(VLOOKUP(C{ROW},▼売上原票!$A:$P,{src_col},0),0)"
            vl.number_format=JPY; vl.font=f(bold=True,sz=11)
            vl.fill=fill("E8F5E9"); vl.alignment=a(h="right"); vl.border=b()

            # 原価（黄色入力欄）
            costs = cost_guide.get(code,(0,0,0,0,0,0))
            for ci,v in enumerate(costs,5):
                c(ws,ROW,ci,v if v>0 else None,sz=10,bg=YEL,ha="right",fmt=JPY)

            # 原価合計
            k_=ws.cell(row=ROW,column=11)
            k_.value=f"=E{ROW}+F{ROW}+G{ROW}+H{ROW}+I{ROW}+J{ROW}"
            k_.number_format=JPY; k_.font=f(bold=True,sz=11)
            k_.fill=fill(YEL); k_.alignment=a(h="right"); k_.border=b()

            # 粗利
            l_=ws.cell(row=ROW,column=12)
            l_.value=f"=D{ROW}-K{ROW}"
            l_.number_format=JPY; l_.font=f(bold=True,sz=12)
            l_.fill=fill(GRNLT); l_.alignment=a(h="right"); l_.border=b()

            # 粗利率
            m_=ws.cell(row=ROW,column=13)
            m_.value=f"=IF(D{ROW}<>0,L{ROW}/D{ROW},0)"
            m_.number_format=PCT; m_.font=f(bold=True,sz=12)
            m_.fill=fill(GRNLT); m_.alignment=a(h="center"); m_.border=b()

            # 採算判定
            n_=ws.cell(row=ROW,column=14)
            n_.value=(f'=IF(D{ROW}=0,"—",'
                      f'IF(M{ROW}>=0.25,"◎ 優良",'
                      f'IF(M{ROW}>=0.15,"○ 良好",'
                      f'IF(M{ROW}>=0.05,"△ 要注意","× 赤字"))))')
            n_.font=f(bold=True,sz=10); n_.fill=fill(GRNLT)
            n_.alignment=a(h="center"); n_.border=b()

            # 主担当案件
            c(ws,ROW,15,route_case.get(code,"—"),sz=10,bg=bg,ha="center")

            # 前月比
            if code in prev_row:
                o_=ws.cell(row=ROW,column=16)
                o_.value=f"=IF(D{prev_row[code]}<>0,D{ROW}/D{prev_row[code]}-1,\"\")"
                o_.number_format="+0.0%;-0.0%;0.0%"; o_.font=f(sz=10)
                o_.fill=fill(bg); o_.alignment=a(h="center"); o_.border=b()
            else:
                c(ws,ROW,16,"—",sz=10,bg=bg,ha="center")
            prev_row[code]=ROW

            # kg推算（売上÷想定単価100円/kgで概算）
            p_=ws.cell(row=ROW,column=17)
            p_.value=f"=IFERROR(D{ROW}/100,0)"
            p_.number_format="#,##0"; p_.font=f(sz=10,color=GRAY)
            p_.fill=fill(bg); p_.alignment=a(h="right"); p_.border=b()

            ROW+=1

        # ルート区切り
        ws.row_dimensions[ROW].height=5
        for col in range(1,18):
            ws.cell(row=ROW,column=col).fill=fill("CCDDEE")
        ROW+=1

    # 条件付き書式
    end=ROW-1
    for expr,fg_c,bg_c in [
        ('N5="◎ 優良"',"145A32","C8F7D4"),
        ('N5="○ 良好"',"1A6B2E",GRNLT),
        ('N5="△ 要注意"',"856404",YEL),
        ('N5="× 赤字"',"8B0000",REDLT),
    ]:
        ws.conditional_formatting.add(f"N5:N{end}",
            FormulaRule(formula=[expr],fill=fill(bg_c),font=Font(bold=True,color=fg_c)))
    ws.conditional_formatting.add(f"M5:M{end}",
        ColorScaleRule(start_type="num",start_value=0,  start_color="FF6B6B",
                       mid_type="num", mid_value=0.15,  mid_color="FFD93D",
                       end_type="num", end_value=0.30,  end_color="6BCB77"))


# ════════════════════════════════════════════════════════════════
#  SHEET 3: ▼売上原票（VLOOKUP元データ）
# ════════════════════════════════════════════════════════════════
def sheet_uriagehyo(ws, sales):
    ws.title = "▼売上原票"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "C3"

    ws.merge_cells("A1:P1")
    t=ws["A1"]
    t.value="▼売上原票　月次売上実績（科目別・税抜）　※このシートは直接編集しないでください"
    t.font=f(bold=True,sz=13,color=WHITE); t.fill=fill("335577")
    t.alignment=a(h="center"); ws.row_dimensions[1].height=30

    ws.merge_cells("A2:P2")
    n2=ws["A2"]
    n2.value="【毎月の更新方法】 新しい月のデータを最終行の下に1行追加してください。A列=YYYY-MM、C〜M列=各科目売上（税抜円）"
    n2.font=f(sz=9,color=GRAY); n2.fill=fill(LGRAY); n2.alignment=a(h="left"); n2.border=b()
    ws.row_dimensions[2].height=16

    widths=[13,9]+[13]*11+[15,11,15]
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width=w

    ws.row_dimensions[3].height=36
    for col,(lbl,w) in enumerate([
        ("年月\n★VLOOKUPキー",13),("R表記",9),
        *[(f"{cd}\n{nm}",13) for cd,nm,_ in ROUTES],
        ("合計\n(税抜)",15),("消費税\n(△)",11),("合計\n(税込)",15)
    ],1):
        c(ws,3,col,lbl,bold=True,sz=9,bg=BLUE,fc=WHITE,ha="center",wrap=True)

    months = sorted(sales.keys())
    for r, ym in enumerate(months, 4):
        ws.row_dimensions[r].height=19
        d = sales[ym]
        era=int(ym[:4])-2018
        era_str=f"R{era}.{int(ym[5:7])}"
        total=d.get("_total",0)
        tax=-round(total*10/11)
        bg=LGRAY if r%2==0 else WHITE

        c(ws,r,1,ym,bold=True,sz=10,bg=bg,ha="center")
        c(ws,r,2,era_str,sz=10,bg=bg,ha="center")
        for ci,code in enumerate(CODES,3):
            v=d.get(code,0)
            c(ws,r,ci,v if v else None,sz=10,bg=bg,ha="right",fmt="#,##0")
        c(ws,r,14,total if total else None,bold=True,sz=11,bg=bg,ha="right",fmt="#,##0")
        c(ws,r,15,tax if total else None,sz=10,bg=bg,ha="right",fmt="#,##0")
        cell16=ws.cell(row=r,column=16)
        cell16.value=f"=IF(N{r}<>\"\",N{r}+O{r},\"\")"
        cell16.number_format="#,##0"; cell16.font=f(bold=True,sz=10)
        cell16.fill=fill(bg); cell16.alignment=a(h="right"); cell16.border=b()

    # 将来月入力枠
    for i in range(12):
        r=3+len(months)+1+i
        ws.row_dimensions[r].height=19
        bg=LGRAY if r%2==0 else WHITE
        c(ws,r,1,"←ここに入力",sz=9,bg="FFFDE7",ha="center",fc="856404")
        c(ws,r,2,"",sz=10,bg=bg,ha="center")
        for ci in range(3,14):
            c(ws,r,ci,None,sz=10,bg=bg,ha="right",fmt="#,##0")
        auto=f"=IF(C{r}+D{r}+E{r}+F{r}+G{r}+H{r}+I{r}+J{r}+K{r}+L{r}+M{r}=0,\"\",C{r}+D{r}+E{r}+F{r}+G{r}+H{r}+I{r}+J{r}+K{r}+L{r}+M{r})"
        c(ws,r,14,auto,bold=True,sz=10,bg=bg,ha="right",fmt="#,##0")
        cell15=ws.cell(row=r,column=15)
        cell15.value=f"=IF(N{r}=\"\",\"\",ROUND(-N{r}/11,0))"
        cell15.number_format="#,##0"; cell15.font=f(sz=10)
        cell15.fill=fill(bg); cell15.alignment=a(h="right"); cell15.border=b()
        cell16=ws.cell(row=r,column=16)
        cell16.value=f"=IF(N{r}=\"\",\"\",N{r}+O{r})"
        cell16.number_format="#,##0"; cell16.font=f(bold=True,sz=10)
        cell16.fill=fill(bg); cell16.alignment=a(h="right"); cell16.border=b()


# ════════════════════════════════════════════════════════════════
#  SHEET 4: ▼マッピング設定
# ════════════════════════════════════════════════════════════════
def sheet_mapping(ws):
    ws.title = "▼マッピング設定"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:G1")
    t=ws["A1"]
    t.value="▼マッピング設定　ルート(科目)→案件の配分率　★初回のみ確認・調整してください"
    t.font=f(bold=True,sz=13,color=WHITE); t.fill=fill("335577")
    t.alignment=a(h="center"); ws.row_dimensions[1].height=30

    ws.merge_cells("A2:G2")
    n2=ws["A2"]
    n2.value="各行の D+E+F=100 になるよう設定してください。615-1(板橋)は東急グループと一般事業者の両方を担当しているため分割設定しています。"
    n2.font=f(sz=9,color=GRAY); n2.fill=fill(LGRAY); n2.alignment=a(h="left"); n2.border=b()
    ws.row_dimensions[2].height=16

    for i,w in enumerate([12,18,30,14,14,14,14],1):
        ws.column_dimensions[get_column_letter(i)].width=w

    ws.row_dimensions[3].height=36
    for col,lbl in enumerate(["科目コード","科目名","配分根拠（メモ）",
                               "C001\n東急グループ%","C002\n都営地下鉄%",
                               "C003\n一般事業者%","合計%\n(100が正常)"],1):
        c(ws,3,col,lbl,bold=True,sz=9,bg=NAVY,fc=WHITE,ha="center",wrap=True)

    reasons = {
        "612-1":"本社が受注する民間運送 → 一般事業者に100%",
        "612-2":"行政・公共入札案件 → 都営地下鉄・区役所等に100%",
        "613-1":"車輌売上（全案件に按分）",
        "613-2":"KCサキュレ分（別管理・0%）",
        "614-1":"コンテナ → 一般事業者に100%",
        "614-2":"ENDO → 一般事業者に100%",
        "614-3":"現金取引 → 一般事業者に100%",
        "615-1":"板橋営業所：東急グループ拠点60%・一般40% ★実績に応じて調整",
        "615-2":"新横浜営業所 → 一般事業者に100%",
        "615-3":"現金2 → 一般事業者に100%",
        "615-4":"有価物 → 一般事業者に100%",
    }

    for r, (code, rname, _) in enumerate(ROUTES, 4):
        ws.row_dimensions[r].height=28
        bg=LGRAY if r%2==0 else WHITE
        p1,p2,p3=MAPPING[code]

        c(ws,r,1,code, bold=True,sz=11,bg=bg,ha="center")
        c(ws,r,2,rname,sz=11,bg=bg)
        c(ws,r,3,reasons.get(code,""),sz=9,bg=bg,wrap=True,fc=GRAY)

        # 入力セル（配分率）
        for col,val,bg_c in [(4,p1,"E8F5E9"),(5,p2,"F3E5F5"),(6,p3,"E0F2F1")]:
            cell=ws.cell(row=r,column=col,value=val)
            cell.number_format="0"; cell.font=f(bold=True,sz=13)
            cell.fill=fill(bg_c); cell.alignment=a(h="center"); cell.border=b()

        # 合計確認
        g_=ws.cell(row=r,column=7)
        g_.value=f"=D{r}+E{r}+F{r}"
        g_.number_format="0"
        g_.font=Font(bold=True,size=13,name="Yu Gothic UI")
        g_.fill=fill("C8F7D4"); g_.alignment=a(h="center"); g_.border=b()
        ws.conditional_formatting.add(f"G{r}:G{r}",
            FormulaRule(formula=[f"G{r}<>100"],
                        fill=fill(REDLT),font=Font(bold=True,color="8B0000")))

    # 凡例
    nr=4+len(ROUTES)+1
    ws.merge_cells(f"A{nr}:G{nr}")
    legend=ws.cell(row=nr,column=1,
        value="【判定基準】 ◎ 優良=粗利率25%以上　○ 良好=15〜25%　△ 要注意=5〜15%　× 赤字=5%未満")
    legend.font=f(sz=10,color=NAVY); legend.fill=fill(LGRAY)
    legend.alignment=a(h="center"); legend.border=b()
    ws.row_dimensions[nr].height=22


# ════════════════════════════════════════════════════════════════
#  BUILD
# ════════════════════════════════════════════════════════════════
def build():
    print("データ読込中...")
    sales = load_sales()
    print(f"  売上データ: {len(sales)}ヶ月")

    wb = Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet()  # 案件別採算
    ws2 = wb.create_sheet()  # ルート別採算
    ws3 = wb.create_sheet()  # ▼売上原票
    ws4 = wb.create_sheet()  # ▼マッピング設定

    sheet_anken(ws1, sales)
    sheet_route(ws2, sales)
    sheet_uriagehyo(ws3, sales)
    sheet_mapping(ws4)

    ws1.sheet_properties.tabColor = NAVY
    ws2.sheet_properties.tabColor = "1A6B2E"
    ws3.sheet_properties.tabColor = "335577"
    ws4.sheet_properties.tabColor = "335577"

    out = "/home/user/sacure/sacure_採算管理_v4_ルート別案件別.xlsx"
    wb.save(out)
    print(f"保存: {out}")
    for ws_ in wb.worksheets:
        print(f"  {ws_.title}")

if __name__ == "__main__":
    build()
