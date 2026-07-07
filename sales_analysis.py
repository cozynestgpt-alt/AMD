"""
매출집계(분석) 보고서 생성 모듈 — 원본 xlsm 불필요
데이터 소스:
  영*판매*.xlsx  → 판매금액, 수금액V+, 생산원가 (소계 판매원가/생산원가)
  run.py 결과   → 매니저(인티포함), 본사-M, 본사-S1, 아르바이트
  expense_report → 경비(직배+경비-DC)
  사원마스터     → 구분(중간관리 여부)

컬럼 구조 (원본 제외 컬럼: M~P, S~U, AB~AD):
  1  매장명(코드)
  2  매장상호
  3  판매금액
  4  수금액(V+)   = 소계 판매원가
  5  점 수수료    = 판매금액 - 수금액V+
  6  비율
  7  수금액(V-)   = ROUND(수금액V+ / 1.1, 0)
  8  생산원가(V-) = 소계 생산원가 (오프+온 합산)
  9  영업이익(V-) = 수금액V- - 생산원가
  10 매니저(인티포함) = 기본수수료 + 인센티브
  11 비율
  12 본사-M 기본수수료
  13 본사-S1 기본수수료
  14 아르바이트
  15 비율
  16 인건비총액(매장) = 매니저+본사M+본사S1+아르바이트
  17 비율
  18 경비(직배+경비-DC)
  19 비율
  20 총경비 = 인건비총액 + 경비
  21 순이익 = 영업이익 - 총경비
  22 순이익율
  23 경비율
  24 구분
"""

import sys
import math
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable,"-m","pip","install","openpyxl","--quiet"])
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

from common import STORE_CODE_MAP, CODE_TO_STORE, ONLINE_CODE_MAP

# ── 색상 ──────────────────────────────────────────────────────
NAVY    = "1F3864"; BLUE    = "2E5EAA"; PALE    = "EEF3FB"
WHITE   = "FFFFFF"; GRAY    = "E8E8E8"; GRAY2   = "F2F2F2"
GREEN   = "1E6B3C"; LGREEN  = "E2EFDA"; AMBER   = "FFF2CC"
RED     = "C00000"; REDL    = "FFE0E0"; ORANGE  = "F4B942"
FMT     = "#,##0"; PCT     = "0.00%"

def _st(ws, r, c, v=None, bg=None, fg="000000", bold=False,
        size=9, align="center", wrap=False, fmt=None):
    cell = ws.cell(r, c)
    if v is not None: cell.value = v
    cell.font = Font(name="맑은 고딕", bold=bold, color=fg, size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if bg: cell.fill = PatternFill("solid", fgColor=bg)
    if fmt: cell.number_format = fmt
    return cell

def _bdr(ws, r1, c1, r2, c2):
    tk = Side(style="medium"); tn = Side(style="thin")
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            ws.cell(r,c).border = Border(
                left=tk if c==c1 else tn, right=tk if c==c2 else tn,
                top=tk if r==r1 else tn, bottom=tk if r==r2 else tn)

# ══════════════════════════════════════════════════════════════
# 1. 영판매 파일에서 수금액V+, 생산원가 추출
# ══════════════════════════════════════════════════════════════

def load_arba_subtotals(path: Path) -> dict:
    """
    아르바이트_정산서.xlsx 점포별소계 시트 → {매장명: 소계금액}
    A열 구분이 '백화점'인 구간의 소계 행(B열에 '점포명  소계')만 추출
    직영점·미입점행사 구간은 제외
    """
    if not path or not path.exists():
        return {}
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if "점포별소계" not in wb.sheetnames:
            return {}
        ws = wb["점포별소계"]
        data = {}
        current_cat = None
        for row in ws.iter_rows(min_row=2, values_only=True):
            a = str(row[0] or "").strip()
            b = str(row[1] or "").strip()
            f = row[5]  # F열 = 금액
            # A열에 구분명이 있으면 현재 구분 업데이트
            if a and a not in ("구분", ""):
                current_cat = a
            # 백화점 구간의 소계 행만 추출
            if current_cat == "백화점" and "소계" in b and isinstance(f, (int, float)):
                shop = b.replace("소계", "").strip()
                data[shop] = int(f)
        return data
    except Exception as e:
        print(f"     ⚠️  아르바이트_정산서 읽기 오류: {e}")
        return {}

def load_sales_vp_cg(path: Path) -> dict:
    """
    영*판매*.xlsx → {매장명: {pf(판매금액), vp(수금액V+), cg(생산원가)}}
    소계 열 기준 (오프라인 + 온라인 코드 행 합산):
      A열(idx0)  = 매장코드
      B열(idx1)  = 매장명
      X열(idx23) = 소계 판매금액    → 매출집계(분석) "판매금액"
      Y열(idx24) = 소계 판매원가    → 매출집계(분석) "수금액V+"
      Z열(idx25) = 소계 생산원가    → 매출집계(분석) "생산원가"
    계산:
      점수수료   = 판매금액 - 수금액V+
      수금액V-   = ROUND(수금액V+ / 1.1, 0)
      영업이익   = 수금액V- - 생산원가
    """
    if not path or not path.exists():
        return {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    data = {}  # {store: {pf, vp, cg}}
    for row in ws.iter_rows(min_row=4, values_only=True):
        code = str(row[0]).strip() if row[0] else ""
        if not code or code in ("매장명","총계"): continue
        pf  = int(row[23]) if isinstance(row[23], (int,float)) else 0  # X열
        vp  = int(row[24]) if isinstance(row[24], (int,float)) else 0  # Y열
        cg  = int(row[25]) if isinstance(row[25], (int,float)) else 0  # Z열
        # 오프라인 매장코드 또는 온라인 매장코드 → 매장명으로 변환
        store = CODE_TO_STORE.get(code) or ONLINE_CODE_MAP.get(code)
        if not store: continue
        is_online = code.startswith("1") and len(code)==5
        if store not in data:
            data[store] = {"pf":0,"vp":0,"cg":0,"off_pf":0,"on_pf":0}
        data[store]["pf"] += pf
        data[store]["vp"] += vp
        data[store]["cg"] += cg
        if is_online:
            data[store]["on_pf"]  += pf   # 온라인 판매금액
        else:
            data[store]["off_pf"] += pf   # 오프라인 판매금액
    return data

# ══════════════════════════════════════════════════════════════
# 2. 보고서 생성
# ══════════════════════════════════════════════════════════════
def make_sales_analysis(ym: str, base_dir: Path,
                        results: list,        # run.py calc() 결과
                        expense_data: dict,   # {shop: total} expense_report
                        sales_path: Path,     # 영*판매*.xlsx
                        master: list,
                        arba_path: Path = None) -> Path:  # 아르바이트_정산서.xlsx
    """
    results     : [{"emp":{...},"res":{...}}, ...]
    expense_data: {매장명: 직배비합계(경비집계 total)}
    sales_path  : 영판매 파일
    master      : 사원마스터 list
    """
    OUTPUT = base_dir / "output" / ym
    OUTPUT.mkdir(parents=True, exist_ok=True)

    # ── 소스 데이터 로드 ─────────────────────────────────────
    sales_vp   = load_sales_vp_cg(sales_path)
    arba_subs  = load_arba_subtotals(arba_path)  # 아르바이트 점포별소계

    # 사원마스터에서 구분 + 매니저 수수료율 추출
    shop_구분    = {}
    shop_mgr_rate = {}  # {shop: {rate_off, rate_on}} — 중간관리 매니저만
    for emp in master:
        shop = emp.get("shop","")
        pt   = emp.get("pay_type","")
        gr   = emp.get("grade","")
        if shop and shop not in shop_구분:
            shop_구분[shop] = "중간관리" if pt == "중간관리" else ""
        # 중간관리 매니저의 수수료율
        if pt == "중간관리" and gr == "매니저" and shop not in shop_mgr_rate:
            shop_mgr_rate[shop] = {
                "rate_off": emp.get("rate_off", 0) or 0,
                "rate_on":  emp.get("rate_on",  0) or 0,
            }

    # calc 결과에서 매장별 집계
    # 같은 매장에 여러 사원이 있으므로 합산
    shop_calc = {}  # {shop: {mgr, bosm, boss1, arba}}
    for r in results:
        emp = r["emp"]; res = r["res"]
        shop  = emp.get("shop","")
        grade = emp.get("grade","")
        if not shop: continue
        if shop not in shop_calc:
            shop_calc[shop] = {"mgr":0,"bosm":0,"boss1":0,"arba":0}
        pay = res.get("total_pay",0) - res.get("vat",0)  # 부가세 제외 지급액
        if grade in ("매니저",): # 중간관리 매니저 = 기본+인센
            shop_calc[shop]["mgr"]   += pay
        elif grade == "본사-M":
            shop_calc[shop]["bosm"]  += emp.get("base_fee",0)
        elif grade == "본사-S1":
            shop_calc[shop]["boss1"] += emp.get("base_fee",0)
        # 아르바이트는 월별입력 사용 안 함 → 아르바이트_정산서 백화점소계 또는 expense_report
        shop_calc[shop]["arba"] = res.get("arba",0)

    # ── 엑셀 생성 ──────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "매출집계(분석)"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    # 열 너비
    #          A   B   C   D   E   F   G   H   I   J   K   L   M   N   O   P   Q   R   S   T   U   V   W   X   Y   Z
    col_widths=[7, 18, 14, 14, 14,  7, 14, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 14, 14, 12, 12, 12, 14,  7,  7, 10]
    for i,w in enumerate(col_widths,1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── 행1: 제목 ──────────────────────────────────────────────
    ws.merge_cells("A1:Z1")
    _st(ws,1,1,f"📊 매출집계 분석  ▶  {ym}",
        bg=NAVY,fg=WHITE,bold=True,size=13)
    ws.row_dimensions[1].height=28

    # ── 행2: 그룹 헤더 ─────────────────────────────────────────
    groups = [
        (1,2,"",NAVY),(3,9,"판매수수료 작업시트 (AMD)",BLUE),
        (10,12,"매니저","4472C4"),(13,14,"본사","70AD47"),
        (15,16,"아르바이트",ORANGE),(17,20,"인건비","ED7D31"),
        (21,24,"손익",NAVY),(25,25,"구분",GRAY),(26,26,"",GRAY),
    ]
    for c1,c2,lbl,bg in groups:
        if c1!=c2: ws.merge_cells(f"{get_column_letter(c1)}2:{get_column_letter(c2)}2")
        _st(ws,2,c1,lbl,bg=bg,fg=WHITE,bold=True,size=8)
    ws.row_dimensions[2].height=14

    # ── 행3: 안내 ──────────────────────────────────────────────
    ws.merge_cells("A3:Z3")
    _st(ws,3,1,
        f"※ 수금액V+ = 영판매파일 소계판매원가(오프+온)  |  생산원가 = 소계생산원가(오프+온)  |  "
        f"수금액V- = ROUND(V+/1.1,0)  |  제외: 시니어1·2, 본사S2, 아르고정, 시니어보장급·원천세",
        bg=PALE,fg="444444",size=8,align="left")
    ws.row_dimensions[3].height=14

    # ── 행4: 컬럼 헤더 ─────────────────────────────────────────
    MGR_BG = "4472C4"
    headers = [
        (1,NAVY,"매장명\n(코드)"),(2,NAVY,"매장상호"),
        (3,BLUE,"판매금액"),(4,BLUE,"수금액\n(V+)"),(5,BLUE,"점 수수료"),(6,BLUE,"비율"),
        (7,BLUE,"수금액\n(V-)"),(8,BLUE,"생산원가\n(V-)"),(9,BLUE,"영업이익\n(V-)"),
        (10,MGR_BG,"오프\n수수료"),(11,MGR_BG,"온라인\n수수료"),(12,MGR_BG,"매니저\n합산"),
        (13,"70AD47","본사-M"),(14,"70AD47","본사-S1"),
        (15,ORANGE,"아르바이트"),(16,ORANGE,"비율"),
        (17,"ED7D31","인건비총액\n(매장)"),(18,"ED7D31","비율"),
        (19,"ED7D31","경비\n(직배+경비-DC)"),(20,"ED7D31","비율"),
        (21,NAVY,"총경비"),(22,NAVY,"순이익"),(23,NAVY,"순이익율"),(24,NAVY,"경비율"),
        (25,"7F7F7F","구분"),
    ]
    for ci,bg,hdr in headers:
        _st(ws,4,ci,hdr,bg=bg,fg=WHITE,bold=True,size=8,wrap=True)
    ws.row_dimensions[4].height=32

    # ── 데이터 행 ──────────────────────────────────────────────
    sorted_shops = sorted(STORE_CODE_MAP.keys(), key=lambda x: STORE_CODE_MAP[x])

    tot = {k:0 for k in ["pf","vp","ps","vn","cg","op","mgr_off","mgr_on","mgr","bosm","boss1",
                          "arba","labor","exp","total_exp","profit"]}
    ei = 0
    for shop in sorted_shops:
        sv = sales_vp.get(shop,{})
        sc = shop_calc.get(shop,{})
        구분 = shop_구분.get(shop,"")

        pf  = sv.get("pf",0)                    # 판매금액
        vp  = sv.get("vp",0)                    # 수금액V+
        ps  = pf - vp                           # 점수수료
        비율  = ps/pf if pf else 0
        vn  = round(vp/1.1)                     # 수금액V-
        cg  = sv.get("cg",0)                    # 생산원가
        op  = vn - cg                           # 영업이익

        # 매니저(인티포함): 중간관리는 수수료율×매출, 본사지급은 calc 결과
        if 구분 == "중간관리" and shop in shop_mgr_rate:
            mr       = shop_mgr_rate[shop]
            sv_m     = sales_vp.get(shop, {})
            off_pf_v = sv_m.get("off_pf", 0)   # 오프라인 판매금액
            on_pf_v  = sv_m.get("on_pf",  0)   # 온라인 판매금액
            # 매니저 합산: 오프+온 합산 후 10원 미만 올림 (run.py calc()의 commission 산식과 동일 방식)
            # round()로 부동소수점 오차 제거 후 ceil 적용 (예: 20194580.000000004 방지)
            combined = round(off_pf_v * mr["rate_off"] + on_pf_v * mr["rate_on"])
            mgr = math.ceil(combined / 10) * 10
            # J/K열 오프·온 분해(표시용, 원 미만 반올림) — 합산은 위 mgr을 그대로 사용
            mgr_off_val = round(off_pf_v * mr["rate_off"])
            mgr_on_val  = round(on_pf_v  * mr["rate_on"])
        else:
            mgr_off_val = 0
            mgr_on_val  = 0
            mgr = sc.get("mgr", 0)
        bosm  = sc.get("bosm",0)
        boss1 = sc.get("boss1",0)
        # 아르바이트: 아르바이트_정산서 백화점 소계 우선
        # 백화점 소계에 없는 경우(중간관리 매장 등) → expense_report의 직배비 계산에 포함됨
        arba  = arba_subs.get(shop, 0)
        labor = mgr + bosm + boss1 + arba       # 인건비총액
        exp   = expense_data.get(shop,0)        # 경비
        total_exp = labor + exp                 # 총경비
        profit    = op - total_exp              # 순이익
        profit_r  = profit/op if op else 0      # 순이익율: 영업이익(V-) 기준
        exp_r     = total_exp/op if op else 0   # 경비율: 영업이익(V-) 기준

        row = ei + 5
        rb  = AMBER if 구분=="중간관리" else (WHITE if ei%2==0 else GRAY2)

        def wc(c, v, fmt=None, fg="000000", bold=False):
            cel = ws.cell(row,c,v)
            cel.font = Font(name="맑은 고딕",bold=bold,
                           color=RED if isinstance(v,(int,float)) and v<0 else fg,
                           size=9)
            cel.fill = PatternFill("solid",fgColor=rb)
            cel.alignment = Alignment(horizontal="left" if c==2 else "center",
                                      vertical="center")
            if fmt: cel.number_format = fmt

        wc(1,  STORE_CODE_MAP.get(shop,""), fg="555555")
        wc(2,  shop, fg="000000")
        wc(3,  pf,    FMT)
        wc(4,  vp,    FMT)
        wc(5,  ps,    FMT)
        wc(6,  비율,   "0.0%")
        wc(7,  vn,    FMT)
        wc(8,  cg,    FMT)
        wc(9,  op,    FMT)
        wc(10, mgr_off_val, FMT)          # J: 오프수수료 (원 미만 반올림)
        wc(11, mgr_on_val,  FMT)          # K: 온라인수수료 (원 미만 반올림)
        # L: 합산 (10원 미만 올림) — 강조 셀
        cl = ws.cell(row,12, mgr)
        cl.font = Font(name="맑은 고딕",bold=True,color=WHITE,size=9)
        cl.fill = PatternFill("solid",fgColor="4472C4")
        cl.number_format = FMT
        cl.alignment = Alignment(horizontal="center",vertical="center")
        wc(13, bosm,  FMT)
        wc(14, boss1, FMT)
        wc(15, arba,  FMT)
        wc(16, arba/op if op else 0, "0.0%")
        wc(17, labor, FMT)
        wc(18, labor/op if op else 0, "0.0%")
        wc(19, exp,   FMT)
        wc(20, exp/op if op else 0, "0.0%")
        wc(21, total_exp, FMT)

        # 순이익 셀 강조
        c22 = ws.cell(row,22,profit)
        c22.font = Font(name="맑은 고딕",bold=True,
                        color=WHITE if profit>0 else RED,size=9)
        c22.fill = PatternFill("solid",fgColor=GREEN if profit>0 else REDL)
        c22.number_format = FMT
        c22.alignment = Alignment(horizontal="center",vertical="center")

        wc(23, profit_r, PCT)
        wc(24, exp_r,    PCT)

        # 구분 셀
        c25 = ws.cell(row,25,구분)
        c25.font = Font(name="맑은 고딕",bold=bool(구분),
                        color="7F4F00" if 구분=="중간관리" else "000000",size=9)
        c25.fill = PatternFill("solid",fgColor=AMBER if 구분=="중간관리" else rb)
        c25.alignment = Alignment(horizontal="center",vertical="center")

        ws.row_dimensions[row].height=17

        # 합계 누적
        for k,v in [("pf",pf),("vp",vp),("ps",ps),("vn",vn),("cg",cg),
                    ("op",op),("mgr_off",mgr_off_val),("mgr_on",mgr_on_val),
                    ("mgr",mgr),("bosm",bosm),("boss1",boss1),
                    ("arba",arba),("labor",labor),("exp",exp),
                    ("total_exp",total_exp),("profit",profit)]:
            tot[k] += v
        ei += 1

    # ── 합계 행 ────────────────────────────────────────────────
    sr = ei + 5
    tot_profit_r = tot["profit"]/tot["op"] if tot["op"] else 0
    tot_exp_r    = tot["total_exp"]/tot["op"] if tot["op"] else 0

    sum_vals = [
        (1,"합 계",None),(2,f"({ei}개 매장)",None),
        (3,tot["pf"],FMT),(4,tot["vp"],FMT),
        (5,tot["ps"],FMT),(6,tot["ps"]/tot["pf"] if tot["pf"] else 0,"0.0%"),
        (7,tot["vn"],FMT),(8,tot["cg"],FMT),(9,tot["op"],FMT),
        (10,tot["mgr_off"],FMT),(11,tot["mgr_on"],FMT),(12,tot["mgr"],FMT),
        (13,tot["bosm"],FMT),(14,tot["boss1"],FMT),
        (15,tot["arba"],FMT),(16,None,None),
        (17,tot["labor"],FMT),(18,None,None),
        (19,tot["exp"],FMT),(20,None,None),
        (21,tot["total_exp"],FMT),(22,tot["profit"],FMT),
        (23,tot_profit_r,PCT),(24,tot_exp_r,PCT),(25,None,None),
    ]
    for ci,v,fmt in sum_vals:
        bg = GREEN if ci==21 and tot["profit"]>0 else NAVY
        fg = WHITE
        c = ws.cell(sr,ci,v if v is not None else "")
        c.font = Font(name="맑은 고딕",bold=True,color=fg,size=10)
        c.fill = PatternFill("solid",fgColor=bg)
        if fmt: c.number_format = fmt
        c.alignment = Alignment(horizontal="left" if ci==2 else "center",
                                vertical="center")
    ws.row_dimensions[sr].height=22

    _bdr(ws,4,1,sr,24)

    out_path = OUTPUT / "매출집계_분석.xlsx"
    wb.save(out_path)
    return out_path


# ══════════════════════════════════════════════════════════════
def run(ym: str, base_dir: Path, results: list,
        expense_totals: dict, sales_path: Path, master: list,
        arba_path: Path = None):
    print(f"  📊 매출집계(분석) 보고서 처리 중...")
    try:
        path = make_sales_analysis(ym, base_dir, results,
                                   expense_totals, sales_path, master,
                                   arba_path)
        print(f"     ✅ 매출집계_분석.xlsx")
        return path
    except Exception as e:
        print(f"     ⚠️  매출집계(분석) 오류: {e}")
        import traceback; traceback.print_exc()
        return None
