"""
경비지원 및 공제 집계 보고서 생성 모듈
입력:
  input/★26년 05월_경비내역서.xlsx     → 직배비 (J열)
  input/◈매장공제건 집계(26년 5월).xlsx → 덜받음(D), 임의사은품+대체할인(G+H), POS공제(J)
  input/26년 05월 재고실사 공제건 합계_수정.xlsx → 유통하자LOSS (I열 공제가합계)
출력:
  output/YYYY-MM/경비지원및공제_집계.xlsx
    시트1: 경비지원및공제  (이미지 형식과 동일)
    시트2: 회수대상목록    (합계 < 0 인 매장)
    시트3: 별도지급목록    (합계 > 0, 근로소득자)
"""

import sys
from pathlib import Path
from glob import glob

try:
    import openpyxl
    import pandas as pd
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "openpyxl", "pandas", "--quiet"])
    import openpyxl, pandas as pd
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

# ── 재고실사 파일 매장명 약칭 → 정식 매장명 매핑 ──────────────
INV_NAME_MAP = {
    "신세계김해":    "신세계김해점",
    "신세계파주":    "신세계아울렛파주점",
    "신세계경기":    "신세계경기점",
    "롯데울산":      "롯데울산점",
    "신세계천안아산":"신세계천안아산점",
    "신세계의정부":  "신세계의정부점",
    "현대중동":      "현대중동점",
    "갤러리아진주":  "갤러리아진주점",
    "현대대구":      "현대대구점",
    "롯데대구":      "롯데대구점",
    "롯데이천아울렛":"롯데아울렛이천점",
    "롯데김해아울렛":"롯데아울렛김해점",
}

# ── 매장코드 정렬 순서 ─────────────────────────────────────────
STORE_CODE_MAP = {
    "롯데잠실점":"00101","롯데관악점":"00102","롯데광주점":"00106",
    "롯데강남점":"00110","롯데포항점":"00111","롯데울산점":"00112",
    "롯데동래점":"00113","롯데창원점":"00114","롯데상인점":"00119",
    "롯데미아점":"00120","롯데구리점":"00125","롯데영등포점":"00126",
    "롯데부산본점":"00128","롯데평촌점":"00129","롯데안산점":"00130",
    "롯데노원점":"00131","롯데광복점":"00132","롯데수원점":"00135",
    "롯데대구점":"00137","롯데전주점":"00139","롯데인천터미널점":"00141",
    "롯데동탄점":"00142","롯데잠실캐슬프라자점":"00143",
    "현대천호점":"00201","현대울산점":"00203","현대압구정본점":"00207",
    "현대미아점":"00208","현대중동점":"00210","현대킨텍스점":"00212",
    "현대대구점":"00213","현대충청점":"00214","현대동구점":"00218",
    "현대판교점":"00219","현대신촌점":"00222",
    "신세계광주점":"00301","신세계마산점":"00303","신세계강남점":"00304",
    "신세계경기점":"00306","신세계센텀점":"00307","신세계타임스퀘어점":"00308",
    "신세계천안아산점":"00309","신세계의정부점":"00310","신세계김해점":"00311",
    "신세계스타필드하남점":"00313","신세계대구점":"00314","신세계대전점":"00315",
    "갤러리아천안점":"00503","갤러리아광교점":"00504","갤러리아진주점":"00505",
    "갤러리아타임월드점":"00506",
    "AK프라자분당점":"00601","AK프라자수원점":"00603",
    "신세계본점":"01305",
    "롯데아울렛김해점":"05002","롯데아울렛이천점":"05003",
    "롯데아울렛광명점":"05004","롯데아울렛고양점":"05005",
    "롯데아울렛광교점":"05006","롯데아울렛군산점":"05007",
    "롯데아울렛동부산점":"05009","LF스퀘어양주점":"05201",
    "신세계아울렛파주점":"05552",
    "현대아울렛대전점":"05601","현대아울렛남양주점":"05602",
}

FMT = "#,##0"

def _st(ws, r, c, v=None, bg=None, fg="000000", bold=False,
        size=9, align="center", wrap=False, fmt=None):
    cell = ws.cell(r, c)
    if v is not None: cell.value = v
    cell.font = Font(name="맑은 고딕", bold=bold, color=fg, size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center",
                                wrap_text=wrap)
    if bg: cell.fill = PatternFill("solid", fgColor=bg)
    if fmt: cell.number_format = fmt
    return cell

def _bdr(ws, r1, c1, r2, c2, outer="medium", inner="thin"):
    tk = Side(style=outer); tn = Side(style=inner)
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            ws.cell(r,c).border = Border(
                left=tk if c==c1 else tn, right=tk if c==c2 else tn,
                top=tk if r==r1 else tn, bottom=tk if r==r2 else tn)

# ══════════════════════════════════════════════════════════════
# 1. 소스 파일 읽기
# ══════════════════════════════════════════════════════════════
def _find(patterns):
    for pat in patterns:
        found = list(Path(".").glob(pat))
        if found: return found[0]
    return None

def load_expense(path: Path) -> dict:
    """경비내역서 → {매장명: 직배비합계}  (26년 5월경비 시트)"""
    if not path or not path.exists():
        return {}
    xl = pd.ExcelFile(path)
    sheet = next((s for s in xl.sheet_names if "5월" in s and "경비" in s),
                 xl.sheet_names[0])
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    data = {}
    for r in range(3, len(df)):
        shop = df.iloc[r, 3]
        val  = df.iloc[r, 11]   # L열(idx11) = 계
        if pd.notna(shop) and str(shop).strip() not in ("", "nan", " 매장명"):
            shop_s = str(shop).strip()
            if pd.notna(val) and isinstance(val, (int, float)):
                data[shop_s] = int(val)
    return data

def load_deduction_detail(path: Path) -> dict:
    """매장공제건 집계 → {매장명: {shortfall, gift, pos}}"""
    if not path or not path.exists():
        return {}
    xl = pd.ExcelFile(path)
    sheet = next((s for s in xl.sheet_names if "매장별공제금액" in s),
                 xl.sheet_names[0])
    df = pd.read_excel(path, sheet_name=sheet, header=0)
    data = {}
    for _, row in df.iterrows():
        shop = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not shop or shop in ("매장명","nan","합계"): continue
        def v(i):
            val = row.iloc[i]
            return int(val) if pd.notna(val) and val else 0
        # D=idx3 덜받음, G=idx6 상품권대체할인, H=idx7 임의사은품합계, J=idx9 POS오류
        data[shop] = {
            "shortfall": v(3),
            "gift":      v(6) + v(7),   # G+H
            "pos":       v(9),
        }
    return data

def load_inventory(path: Path) -> dict:
    """재고실사 → {매장명: 공제가합계}  (I열=idx8)"""
    if not path or not path.exists():
        return {}
    xl = pd.ExcelFile(path)
    # "Sheet1 (2)" 가 5월분
    sheet = next((s for s in xl.sheet_names if "(2)" in s or "05" in s or "5월" in s),
                 xl.sheet_names[0])
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    data = {}
    for r in range(4, len(df)):
        shop = df.iloc[r, 1]
        val  = df.iloc[r, 8]   # I열
        if pd.notna(shop) and str(shop).strip() and str(shop) != "nan":
            shop_raw = str(shop).strip()
            shop_s   = INV_NAME_MAP.get(shop_raw, shop_raw)
            if pd.notna(val) and isinstance(val, (int, float)) and val > 0:
                data[shop_s] = data.get(shop_s, 0) + int(val)
    return data

# ══════════════════════════════════════════════════════════════
# 2. 보고서 생성
# ══════════════════════════════════════════════════════════════
def make_expense_report(ym: str, base_dir: Path, master: list) -> Path:
    """
    master: 사원마스터 list [{shop, name, grade, income, pay_type, bank, account}]
    """
    INPUT  = base_dir / "input"
    OUTPUT = base_dir / "output" / ym
    OUTPUT.mkdir(parents=True, exist_ok=True)

    # ── 소스 파일 탐지 ───────────────────────────────────────
    def find_in(patterns):
        for pat in patterns:
            found = list(INPUT.glob(pat))
            if found: return found[0]
        return None

    exp_path = find_in(["★*경비내역서*.xlsx", "*경비내역서*.xlsx"])
    ded_path = find_in(["◈*매장공제건*.xlsx", "*매장공제건*.xlsx", "◈*.xlsx"])
    inv_path = find_in(["*재고실사*공제건*합계*.xlsx", "*재고실사*.xlsx"])

    print(f"     직배비 파일: {exp_path.name if exp_path else '없음 → 0처리'}")
    print(f"     공제건집계:  {ded_path.name if ded_path else '없음 → 0처리'}")
    print(f"     재고실사:    {inv_path.name if inv_path else '없음 → 0처리'}")

    exp_data = load_expense(exp_path)
    ded_data = load_deduction_detail(ded_path)
    inv_data = load_inventory(inv_path)

    # ── 사원마스터에서 매장 정보 추출 ────────────────────────
    shop_info = {}   # {shop: {mgr_name, income, pay_type, bank, account}}
    for emp in master:
        shop = emp.get("shop","")
        if not shop: continue
        pt = emp.get("pay_type","")
        gr = emp.get("grade","")
        # 대표자 우선 (매니저 or 본사-M)
        if pt == "중간관리" or gr in ("매니저","본사-M"):
            if shop not in shop_info or gr == "매니저":
                shop_info[shop] = {
                    "mgr":     emp.get("name",""),
                    "income":  emp.get("income",""),
                    "pay_type":pt,
                    "bank":    emp.get("bank",""),
                    "account": emp.get("account",""),
                    "note":    emp.get("note",""),
                }
        # 행사매장은 사원 유무 무관하게 note 등록
        if (emp.get("note","") or "").strip() == "행사매장":
            if shop not in shop_info:
                shop_info[shop] = {}
            shop_info[shop]["note"] = "행사매장"

    # ── 전체 매장 목록 (매장코드 오름차순) ───────────────────
    all_shops = sorted(STORE_CODE_MAP.keys(), key=lambda x: STORE_CODE_MAP[x])

    # 행사매장 목록 (사원마스터 비고="행사매장" → POS환급 배분 제외)
    event_shops = {
        shop for shop, inf in shop_info.items()
        if (inf.get("note","") or "").strip() == "행사매장"
    }

    # POS환급 계산: ded_data에 있고 행사매장이 아닌 POS=0 매장에 균등 배분
    pos_total    = sum(d.get("pos",0) for d in ded_data.values())
    pos_zero_cnt = sum(
        1 for shop, d in ded_data.items()
        if d.get("pos",0) == 0 and shop not in event_shops
    )
    pos_refund_unit = (int(pos_total / pos_zero_cnt / 10) * 10
                       if pos_zero_cnt > 0 else 0)
    print(f"     POS환급 단가: {pos_refund_unit:,}원 × {pos_zero_cnt}개 매장 = {pos_refund_unit*pos_zero_cnt:,}원")
    print(f"     행사매장 제외: {sorted(event_shops) if event_shops else '없음'}")

    # ── 행 데이터 계산 ────────────────────────────────────────
    rows = []
    for shop in all_shops:
        d   = ded_data.get(shop, {})
        inf = shop_info.get(shop, {})
        j_direct   = exp_data.get(shop, 0)          # J: 직배비
        l_short    = d.get("shortfall", 0)           # L: 덜받음
        n_loss     = inv_data.get(shop, 0)           # N: 유통하자LOSS
        o_gift     = d.get("gift", 0)                # O: 임의사은품+대체할인
        s_pos      = d.get("pos", 0)                 # S: POS정정공제
        # T: POS환급 — POS공제가 없고 행사매장이 아닌 경우에만 배분
        is_event_shop = shop in event_shops
        t_refund = (pos_refund_unit
                    if s_pos == 0 and not is_event_shop
                    else 0)

        # 합계 = +J -L -N -O -S +T
        total = j_direct - l_short - n_loss - o_gift - s_pos + t_refund

        # 비고 (지급방법)
        pay_type = inf.get("pay_type","")
        bigo = "중간관리" if pay_type == "중간관리" else "본사지급"

        rows.append({
            "shop":     shop,
            "mgr":      inf.get("mgr",""),
            "income":   inf.get("income",""),
            "pay_type": pay_type,
            "bank":     inf.get("bank",""),
            "account":  inf.get("account",""),
            "j_direct": j_direct,
            "l_short":  l_short,
            "n_loss":   n_loss,
            "o_gift":   o_gift,
            "s_pos":    s_pos,
            "t_refund": t_refund,
            "total":    total,
            "bigo":     bigo,
        })

    # ══════════════════════════════════════════════════════════
    # 시트1: 경비지원및공제 집계
    # ══════════════════════════════════════════════════════════
    NAVY="1F3864"; BLUE="2E5EAA"; PALE="EEF3FB"
    WHITE="FFFFFF"; GRAY="E8E8E8"; LGREEN="E2EFDA"
    RED="C00000"; REDL="FFE0E0"; GREEN="1E6B3C"
    AMBER="FFF2CC"; YELLOW="FFFF00"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "경비지원및공제"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A8"

    # 열 너비 (A~N)
    col_w = [22, 8, 12, 10, 10, 14, 12, 10, 10, 12, 10, 14, 10]
    for i,w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── 헤더 (이미지와 동일 구조) ────────────────────────────
    # 행1~2: 그룹 헤더
    ws.row_dimensions[1].height = 10
    ws.merge_cells("A2:B2")
    _st(ws,2,1,"매 장 경 비",bg=BLUE,fg=WHITE,bold=True,size=9)
    ws.merge_cells("C2:I2")
    _st(ws,2,3,"공  제",bg="C00000",fg=WHITE,bold=True,size=9)
    ws.merge_cells("J2:K2")
    _st(ws,2,10,"공제 및 환급",bg="7F7F7F",fg=WHITE,bold=True,size=9)
    ws.merge_cells("L2:M2")
    _st(ws,2,12,"소  계",bg=NAVY,fg=WHITE,bold=True,size=9)
    ws.row_dimensions[2].height = 16

    # 행3~6: 복합 헤더 (이미지 기준)
    # 행3
    ws.merge_cells("A3:A6"); _st(ws,3,1,"매  장  명",bg=GRAY,bold=True,size=9)
    ws.merge_cells("B3:B6"); _st(ws,3,2,"직배비",bg=GRAY,bold=True,size=9,wrap=True)
    ws.merge_cells("C3:C6"); _st(ws,3,3,"덜받음",bg="FCE4D6",bold=True,size=9,wrap=True)
    ws.merge_cells("D3:D6"); _st(ws,3,4,"인센티브\n공제",bg="FCE4D6",bold=True,size=9,wrap=True)
    ws.merge_cells("E3:E6"); _st(ws,3,5,"유통하자\n/재고LOSS",bg="FCE4D6",bold=True,size=9,wrap=True)
    ws.merge_cells("F3:F6"); _st(ws,3,6,"임의사은품,\n대체할인지급",bg="FCE4D6",bold=True,size=9,wrap=True)
    ws.merge_cells("G3:G6"); _st(ws,3,7,"임의\n추가할인",bg="FCE4D6",bold=True,size=9,wrap=True)
    ws.merge_cells("H3:H6"); _st(ws,3,8,"임의사은품,\n추가할인\n공제환급",bg="FCE4D6",bold=True,size=9,wrap=True)
    ws.merge_cells("I3:I6"); _st(ws,3,9,"계",bg="FCE4D6",bold=True,size=9)
    ws.merge_cells("J3:J6"); _st(ws,3,10,"POS정정\n요청공제",bg="D9D9D9",bold=True,size=9,wrap=True)
    ws.merge_cells("K3:K6"); _st(ws,3,11,"POS정정\n요청환급\n(배분)",bg="D9D9D9",bold=True,size=9,wrap=True)
    ws.merge_cells("L3:L6"); _st(ws,3,12,"계",bg=NAVY,fg=WHITE,bold=True,size=9)
    ws.merge_cells("M3:M6"); _st(ws,3,13,"비고",bg=NAVY,fg=WHITE,bold=True,size=9)
    for r in range(3,7): ws.row_dimensions[r].height = 18

    # 행7: 수식 설명
    ws.merge_cells("A7:M7")
    _st(ws,7,1,
        f"※ 계 = (+직배비) - (-덜받음) - (-유통하자LOSS) - (-임의사은품) - (-POS정정공제) + (+POS환급)  |  POS환급 단가: {pos_refund_unit:,}원 (S열=0인 {pos_zero_cnt}개 매장 배분)",
        bg=AMBER, fg="664400", size=8, align="left")
    ws.row_dimensions[7].height = 14

    # ── 데이터 행 ─────────────────────────────────────────────
    for ei, r in enumerate(rows):
        row = ei + 8
        # 배경: 매장 교대 + 합계 <0 빨강, >0 연초록
        if r["total"] < 0:
            rb = REDL
        elif r["total"] > 0 and r["bigo"] == "본사지급":
            rb = LGREEN
        else:
            rb = WHITE if ei % 2 == 0 else GRAY

        _st(ws,row,1, r["shop"], bg=rb, size=9, align="left")
        # B: 직배비
        _st(ws,row,2, r["j_direct"] or "", bg=rb, size=9,
            fmt=FMT if r["j_direct"] else None)
        # C: 덜받음 (음수 표시)
        v = -r["l_short"] if r["l_short"] else ""
        _st(ws,row,3, v, bg=rb, size=9,
            fg=RED if r["l_short"] else "000000",
            fmt=FMT if r["l_short"] else None)
        # D: 인센티브공제 (없음)
        _st(ws,row,4, "", bg=rb, size=9)
        # E: 유통하자LOSS (음수)
        v = -r["n_loss"] if r["n_loss"] else ""
        _st(ws,row,5, v, bg=rb, size=9,
            fg=RED if r["n_loss"] else "000000",
            fmt=FMT if r["n_loss"] else None)
        # F: 임의사은품 (음수)
        v = -r["o_gift"] if r["o_gift"] else ""
        _st(ws,row,6, v, bg=rb, size=9,
            fg=RED if r["o_gift"] else "000000",
            fmt=FMT if r["o_gift"] else None)
        # G: 임의추가할인 (없음)
        _st(ws,row,7, "", bg=rb, size=9)
        # H: 공제환급 (없음)
        _st(ws,row,8, "", bg=rb, size=9)
        # I: 공제계 = -(L+N+O)
        공제계 = -(r["l_short"] + r["n_loss"] + r["o_gift"])
        _st(ws,row,9, 공제계 if 공제계 else "", bg=rb, size=9,
            fg=RED if 공제계 < 0 else "000000",
            fmt=FMT if 공제계 else None)
        # J: POS정정공제
        _st(ws,row,10, -r["s_pos"] if r["s_pos"] else "", bg=rb, size=9,
            fg=RED if r["s_pos"] else "000000",
            fmt=FMT if r["s_pos"] else None)
        # K: POS환급
        _st(ws,row,11, r["t_refund"] if r["t_refund"] else "", bg=rb, size=9,
            fmt=FMT if r["t_refund"] else None)
        # L: 합계 (소계)
        tot_bg = RED if r["total"] < 0 else (GREEN if r["total"] > 0 else rb)
        tot_fg = WHITE if r["total"] != 0 else "000000"
        c = ws.cell(row, 12, r["total"] if r["total"] else 0)
        c.font = Font(name="맑은 고딕", bold=True, color=tot_fg, size=9)
        c.fill = PatternFill("solid", fgColor=tot_bg)
        c.number_format = FMT
        c.alignment = Alignment(horizontal="center", vertical="center")
        # M: 비고
        bigo_bg = AMBER if r["bigo"]=="중간관리" else PALE
        _st(ws,row,13, r["bigo"], bg=bigo_bg, size=9)
        ws.row_dimensions[row].height = 17

    # 합계행
    sr = len(rows) + 8
    ws.merge_cells(f"A{sr}:A{sr}")
    _st(ws,sr,1,"합  계",bg=NAVY,fg=WHITE,bold=True,size=10)
    sum_cols = {
        2:  sum(r["j_direct"] for r in rows),
        3:  -sum(r["l_short"] for r in rows),
        5:  -sum(r["n_loss"] for r in rows),
        6:  -sum(r["o_gift"] for r in rows),
        9:  -sum(r["l_short"]+r["n_loss"]+r["o_gift"] for r in rows),
        10: -sum(r["s_pos"] for r in rows),
        11:  sum(r["t_refund"] for r in rows),
        12:  sum(r["total"] for r in rows),
    }
    for ci, val in sum_cols.items():
        bg = GREEN if ci==12 and val>0 else (RED if ci==12 and val<0 else NAVY)
        fg = WHITE if ci in [12] or val < 0 else WHITE
        c = ws.cell(sr, ci, val if val else 0)
        c.font = Font(name="맑은 고딕", bold=True, color=WHITE, size=10)
        c.fill = PatternFill("solid", fgColor=bg)
        c.number_format = FMT
        c.alignment = Alignment(horizontal="center", vertical="center")
    for ci in [4,7,8,13]:
        ws.cell(sr,ci).fill=PatternFill("solid",fgColor=NAVY)
    ws.row_dimensions[sr].height=22
    _bdr(ws,2,1,sr,13)

    # ══════════════════════════════════════════════════════════
    # 시트2: 회수대상목록 (합계 < 0)
    # ══════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("회수대상목록")
    ws2.sheet_view.showGridLines = False
    for i,w in enumerate([5,22,10,14,12],1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    ws2.merge_cells("A1:E1")
    _st(ws2,1,1,f"⚠️  회수 대상 매장  ▶  {ym}  (합계 < 0 → 현금 회수 필요)",
        bg=RED,fg=WHITE,bold=True,size=12)
    ws2.row_dimensions[1].height=26
    ws2.merge_cells("A2:E2")
    _st(ws2,2,1,"공제 합계가 음수인 매장: 직배비보다 공제금이 많아 현금으로 회수해야 합니다.",
        bg=REDL,fg=RED,size=9,align="left")
    ws2.row_dimensions[2].height=14
    for ci,h in enumerate([("NO",NAVY),("매장명",NAVY),("매니저",NAVY),
                            ("회수금액",RED),("비고",NAVY)],1):
        _st(ws2,3,ci,h[0],bg=h[1],fg=WHITE,bold=True,size=9)
    ws2.row_dimensions[3].height=20

    rev_rows = [r for r in rows if r["total"] < 0 and r["pay_type"] != "중간관리"]
    for ei,r in enumerate(rev_rows):
        row=ei+4; rb=WHITE if ei%2==0 else REDL
        _st(ws2,row,1,ei+1,bg=rb,size=9)
        _st(ws2,row,2,r["shop"],bg=rb,size=9,align="left")
        _st(ws2,row,3,r["mgr"],bg=rb,size=9)
        c=ws2.cell(row,4,abs(r["total"]))
        c.font=Font(name="맑은 고딕",bold=True,color=WHITE,size=10)
        c.fill=PatternFill("solid",fgColor=RED); c.number_format=FMT
        c.alignment=Alignment(horizontal="center",vertical="center")
        _st(ws2,row,5,r["bigo"],bg=AMBER if r["bigo"]=="중간관리" else PALE,size=9)
        ws2.row_dimensions[row].height=18

    sr2=len(rev_rows)+4
    ws2.merge_cells(f"A{sr2}:C{sr2}")
    _st(ws2,sr2,1,f"합계  ({len(rev_rows)}개 매장)",bg=NAVY,fg=WHITE,bold=True,size=10)
    c=ws2.cell(sr2,4,sum(abs(r["total"]) for r in rev_rows))
    c.font=Font(name="맑은 고딕",bold=True,color=WHITE,size=11)
    c.fill=PatternFill("solid",fgColor=RED); c.number_format=FMT
    c.alignment=Alignment(horizontal="center",vertical="center")
    ws2.row_dimensions[sr2].height=22
    _bdr(ws2,3,1,sr2,5)

    # ══════════════════════════════════════════════════════════
    # 시트3: 별도지급목록 (합계 > 0, 근로소득자)
    # ══════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("별도지급목록")
    ws3.sheet_view.showGridLines = False
    for i,w in enumerate([5,22,10,12,22,14],1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    ws3.merge_cells("A1:F1")
    _st(ws3,1,1,f"💳  별도 지급 대상  ▶  {ym}  (합계 > 0, 근로소득자 본사지급)",
        bg=GREEN,fg=WHITE,bold=True,size=12)
    ws3.row_dimensions[1].height=26
    ws3.merge_cells("A2:F2")
    _st(ws3,2,1,"직배비가 공제금보다 많은 근로소득자: 차액을 별도로 이체 지급합니다. (중간관리자는 매출집계(분석)에서 처리)",
        bg=LGREEN,fg=GREEN,size=9,align="left")
    ws3.row_dimensions[2].height=14
    for ci,h in enumerate([("NO",NAVY),("매장명",NAVY),("매니저",NAVY),
                            ("거래은행",NAVY),("계좌번호",NAVY),("지급금액",GREEN)],1):
        _st(ws3,3,ci,h[0],bg=h[1],fg=WHITE,bold=True,size=9)
    ws3.row_dimensions[3].height=20

    pay_rows = [r for r in rows
                if r["total"] > 0 and r["income"] in ("근로소득",)
                and r["pay_type"] != "중간관리"]
    for ei,r in enumerate(pay_rows):
        row=ei+4; rb=WHITE if ei%2==0 else LGREEN
        _st(ws3,row,1,ei+1,bg=rb,size=9)
        _st(ws3,row,2,r["shop"],bg=rb,size=9,align="left")
        _st(ws3,row,3,r["mgr"],bg=rb,size=9)
        _st(ws3,row,4,r["bank"],bg=rb,size=9)
        _st(ws3,row,5,r["account"],bg=rb,size=9,align="left")
        c=ws3.cell(row,6,r["total"])
        c.font=Font(name="맑은 고딕",bold=True,color=WHITE,size=10)
        c.fill=PatternFill("solid",fgColor=GREEN); c.number_format=FMT
        c.alignment=Alignment(horizontal="center",vertical="center")
        ws3.row_dimensions[row].height=18

    sr3=len(pay_rows)+4
    ws3.merge_cells(f"A{sr3}:E{sr3}")
    _st(ws3,sr3,1,f"합계  ({len(pay_rows)}개 매장)",bg=NAVY,fg=WHITE,bold=True,size=10)
    c=ws3.cell(sr3,6,sum(r["total"] for r in pay_rows))
    c.font=Font(name="맑은 고딕",bold=True,color=WHITE,size=11)
    c.fill=PatternFill("solid",fgColor=GREEN); c.number_format=FMT
    c.alignment=Alignment(horizontal="center",vertical="center")
    ws3.row_dimensions[sr3].height=22
    _bdr(ws3,3,1,sr3,6)

    # 저장
    out_path = OUTPUT / "경비지원및공제_집계.xlsx"
    wb.save(out_path)
    return out_path, len(rev_rows), len(pay_rows)

# ══════════════════════════════════════════════════════════════
# 메인 (단독 실행 테스트용)
# ══════════════════════════════════════════════════════════════
def run(ym: str, base_dir: Path, master: list):
    print(f"  📋 경비지원및공제 집계 처리 중...")
    try:
        path, rev, pay = make_expense_report(ym, base_dir, master)
        print(f"     ✅ 경비지원및공제_집계.xlsx  "
              f"(회수 {rev}개 매장, 별도지급 {pay}개 매장)")
        return path
    except Exception as e:
        print(f"     ⚠️  경비지원및공제 처리 오류: {e}")
        import traceback; traceback.print_exc()
        return None

if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).parent))
    from run import load_master, INPUT, BASE
    master = load_master(INPUT / "사원마스터.xlsx")
    run("2026-05", BASE, master)
