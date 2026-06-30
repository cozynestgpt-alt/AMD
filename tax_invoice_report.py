"""
매입세금계산서 자동 출력 모듈 v2
- 템플릿 파일을 매장별로 직접 로드·수정하여 각 시트 생성
- 인쇄 설정을 템플릿과 동일하게 유지 (A4 세로, scale=99, 1페이지 맞춤)
"""
import math, copy, calendar
from pathlib import Path
from collections import defaultdict
from datetime import date

try:
    from openpyxl import load_workbook
    import openpyxl
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable,"-m","pip","install","openpyxl","--quiet"])
    from openpyxl import load_workbook
    import openpyxl


def load_supplier_info(path: Path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    본사 = {
        "상호명":   ws.cell(3,4).value or "(주)더코지네스트컴퍼니",
        "사업자번호": ws.cell(3,5).value or "211-81-83236",
        "대표자":   ws.cell(3,6).value or "김창수",
        "주소":     ws.cell(3,7).value or "서울시 강남구 논현로 752, 3층",
        "업태":     ws.cell(3,8).value or "제조,도소매",
        "종목":     ws.cell(3,9).value or "침구 및 관련제품",
    }
    data = {}
    for r in range(4, ws.max_row+1):
        no = ws.cell(r,2).value
        if not isinstance(no,(int,float)): continue
        전산명 = str(ws.cell(r,3).value or "").strip()
        if not 전산명: continue
        data[전산명] = {
            "no":       int(no),
            "상호명":   ws.cell(r,4).value or "",
            "사업자번호": ws.cell(r,5).value or "",
            "대표자":   ws.cell(r,6).value or "",
            "주소":     ws.cell(r,7).value or "",
            "업태":     ws.cell(r,8).value or "",
            "종목":     ws.cell(r,9).value or "",
            "담당자":   ws.cell(r,11).value or "" if ws.max_column>=11 else "",
        }
    return 본사, data


def _clear_formulas(ws):
    """수식·외부참조 제거"""
    for row in ws.iter_rows():
        for cell in row:
            if cell.__class__.__name__ == "MergedCell": continue
            if isinstance(cell.value, str) and ("=" in cell.value or "#REF!" in cell.value):
                cell.value = None


def _sv(ws, r, c, v):
    """안전 셀 입력 (MergedCell 건너뜀)"""
    try:
        cell = ws.cell(r, c)
        if cell.__class__.__name__ != "MergedCell":
            cell.value = v
    except: pass


def _fill_block(ws, base_row, info, 공급가액, 세액, 작성일):
    """
    세금계산서 1블록 값 입력
    base_row=11 → 상단(공급받는자보관용, 행11~30)
    base_row=33 → 하단(공급자보관용, 행33~52)
    """
    br = base_row

    # 공급자 정보
    _sv(ws, br+3, 6,  info.get("사업자번호",""))   # 등록번호
    _sv(ws, br+4, 6,  info.get("상호명",""))        # 상호
    _sv(ws, br+4, 14, info.get("대표자",""))        # 성명
    _sv(ws, br+6, 6,  info.get("주소",""))          # 사업장주소
    _sv(ws, br+8, 6,  info.get("업태",""))          # 업태
    _sv(ws, br+8, 12, info.get("종목",""))          # 종목

    # 작성년월일 (행22/44 = br+11)
    _sv(ws, br+11, 2, 작성일.year)
    _sv(ws, br+11, 4, 작성일.month)
    _sv(ws, br+11, 5, 작성일.day)

    # 공급가액 자릿별 (G~Q = 열7~17)
    for i, d in enumerate(reversed(str(공급가액))):
        col = 17 - i
        if 7 <= col <= 17: _sv(ws, br+11, col, d)
    # 세액 자릿별 (R~AA = 열18~27)
    for i, d in enumerate(reversed(str(세액))):
        col = 27 - i
        if 18 <= col <= 27: _sv(ws, br+11, col, d)

    # 품목행 (행24/46 = br+13)
    _sv(ws, br+13, 4, "판매수수료")
    ws.cell(br+13, 20).value = 공급가액
    ws.cell(br+13, 20).number_format = "#,##0"
    ws.cell(br+13, 26).value = 세액
    ws.cell(br+13, 26).number_format = "#,##0"

    # 합계금액 (행29/51 = br+18)
    ws.cell(br+18, 2).value = 공급가액 + 세액
    ws.cell(br+18, 2).number_format = "#,##0"


def _apply_page_setup(ws, print_area="A10:AH53"):
    """템플릿과 동일한 인쇄 설정 적용"""
    ws.print_area = print_area
    ws.page_setup.paperSize   = 9          # A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.scale       = 99
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered   = True
    ws.page_margins.left   = 0.19685   # 5mm
    ws.page_margins.right  = 0.19685
    ws.page_margins.top    = 0.59055   # 15mm
    ws.page_margins.bottom = 0.59055
    ws.page_margins.header = 0.31496
    ws.page_margins.footer = 0.31496


def make_tax_invoice(ym: str, base_dir: Path,
                     results: list, expense_rows: dict, master: list) -> Path:
    INPUT    = base_dir / "input"
    TMPL_DIR = base_dir / "templates"
    OUTPUT   = base_dir / "output" / ym
    OUTPUT.mkdir(parents=True, exist_ok=True)

    info_path = next(iter(INPUT.glob("중간관리매장_매입세금계산서*.xlsx")), None)
    tmpl_path = TMPL_DIR / "세금계산서출력_템플릿.xlsx"
    if not info_path:
        print(f"     ⚠️  기재사항정보 파일 없음: input/중간관리매장_매입세금계산서_기재사항정보.xlsx")
        return None
    if not tmpl_path.exists():
        print(f"     ⚠️  템플릿 없음: templates/세금계산서출력_템플릿.xlsx")
        return None

    본사_info, supplier_info = load_supplier_info(info_path)

    from expense_report import STORE_CODE_MAP
    event_shops = {e["shop"] for e in master if (e.get("note","") or "").strip()=="행사매장"}
    mgr_shops   = {e["shop"] for e in master if e.get("pay_type","")=="중간관리"}

    shop_fee = defaultdict(int); shop_income = {}
    for r in results:
        emp, res = r["emp"], r["res"]
        if not emp["name"]: continue
        shop_fee[emp["shop"]] += emp["base_fee"] + res["commission"]
        if emp["shop"] not in shop_income:
            shop_income[emp["shop"]] = emp["income"]

    # 중간관리 순서
    mgr_order = []; seen = set()
    for r in results:
        shop = r["emp"]["shop"]
        if r["emp"].get("pay_type","")=="중간관리" and shop not in seen:
            seen.add(shop); mgr_order.append(shop)
    invoice_shops = [s for s in mgr_order if shop_income.get(s)=="사업소득"]

    year,month = int(ym[:4]),int(ym[5:])
    작성일 = date(year, month, calendar.monthrange(year,month)[1])

    # ── 매장별 시트 생성 ─────────────────────────────────────
    # 각 매장마다 템플릿을 열고 값을 주입, 하나의 워크북에 합침
    wb_out = openpyxl.Workbook()
    del wb_out["Sheet"]

    shop_values = {}

    for shop in invoice_shops:
        info = supplier_info.get(shop, {})
        er   = expense_rows.get(shop, {})
        fee_c = math.ceil(shop_fee[shop] / 10) * 10
        공제  = (er.get("shortfall",0)+er.get("gift",0)+er.get("pos",0)
                 -er.get("t_refund",0)+er.get("loss",0))
        공급가액 = fee_c + er.get("j_direct",0) - 공제
        세액     = int(공급가액 * 0.1)
        shop_values[shop] = {"공급가액": 공급가액, "세액": 세액}

        # 템플릿 로드 (매장마다 새로 열기)
        wb_tmpl = load_workbook(tmpl_path)
        ws_tmpl = wb_tmpl["세금계산서(출력)"]

        # 수식 제거
        _clear_formulas(ws_tmpl)

        # 값 주입
        _fill_block(ws_tmpl, 11, info, 공급가액, 세액, 작성일)  # 상단
        _fill_block(ws_tmpl, 33, info, 공급가액, 세액, 작성일)  # 하단

        # 인쇄 설정
        _apply_page_setup(ws_tmpl)

        # 출력 워크북에 시트 복사
        순번 = info.get("no", invoice_shops.index(shop)+1)
        sheet_nm = f"{순번:02d}_{shop}"[:31]
        ws_out = wb_out.create_sheet(sheet_nm)

        # 열너비/행높이 복사
        for col, cd in ws_tmpl.column_dimensions.items():
            ws_out.column_dimensions[col].width = cd.width
        for row, rd in ws_tmpl.row_dimensions.items():
            ws_out.row_dimensions[row].height = rd.height

        # 스타일+값 복사 (MergedCell 건너뜀)
        for sr in ws_tmpl.iter_rows():
            for sc in sr:
                if sc.__class__.__name__ == "MergedCell": continue
                dc = ws_out.cell(sc.row, sc.column)
                dc.value = sc.value
                if sc.has_style:
                    dc.font      = copy.copy(sc.font)
                    dc.fill      = copy.copy(sc.fill)
                    dc.border    = copy.copy(sc.border)
                    dc.alignment = copy.copy(sc.alignment)
                    dc.number_format = sc.number_format

        # 병합셀 (스타일 복사 후)
        for mr in ws_tmpl.merged_cells.ranges:
            try: ws_out.merge_cells(str(mr))
            except: pass

        # 인쇄 설정 (복사 후 다시 적용)
        _apply_page_setup(ws_out)
        ws_out.sheet_view.showGridLines = False

    # ── 매입세금계산서_list 시트 추가 ───────────────────────
    wb_tmpl2 = load_workbook(tmpl_path)
    ws_src = wb_tmpl2["매입세금계산서_list"]
    ws_list = wb_out.create_sheet("매입세금계산서_list")

    for col, cd in ws_src.column_dimensions.items():
        ws_list.column_dimensions[col].width = cd.width
    for row, rd in ws_src.row_dimensions.items():
        ws_list.row_dimensions[row].height = rd.height
    for sr in ws_src.iter_rows():
        for sc in sr:
            if sc.__class__.__name__ == "MergedCell": continue
            dc = ws_list.cell(sc.row, sc.column)
            if isinstance(sc.value,str) and ("#REF!" in sc.value or sc.value.startswith("=")):
                dc.value = None
            else:
                dc.value = sc.value
            if sc.has_style:
                dc.font=copy.copy(sc.font); dc.fill=copy.copy(sc.fill)
                dc.border=copy.copy(sc.border); dc.alignment=copy.copy(sc.alignment)
                dc.number_format=sc.number_format
    for mr in ws_src.merged_cells.ranges:
        try: ws_list.merge_cells(str(mr))
        except: pass

    FMT="#,##0"
    total_공급=0; total_세액=0
    for i,shop in enumerate(invoice_shops):
        row=6+i
        공급가액=shop_values[shop]["공급가액"]
        세액=shop_values[shop]["세액"]
        합계=공급가액+세액
        for c,v in [(12,공급가액),(13,공급가액),(14,세액),(15,합계)]:
            try:
                ws_list.cell(row,c).value=v
                ws_list.cell(row,c).number_format=FMT
            except: pass
        total_공급+=공급가액; total_세액+=세액
    for c,v in [(12,total_공급),(13,total_공급),(14,total_세액),(15,total_공급+total_세액)]:
        try:
            ws_list.cell(2,c).value=v
            ws_list.cell(2,c).number_format=FMT
        except: pass
    ws_list.sheet_view.showGridLines=False

    out_path = OUTPUT / f"세금계산서_{ym}.xlsx"
    wb_out.save(out_path)
    return out_path


def run(ym, base_dir, results, expense_rows, master):
    print(f"  🧾 세금계산서 처리 중...")
    try:
        path = make_tax_invoice(ym, base_dir, results, expense_rows, master)
        if path:
            print(f"     ✅ 세금계산서_{ym}.xlsx  (24개 매장)")
        return path
    except Exception as e:
        print(f"     ⚠️  세금계산서 처리 오류: {e}")
        import traceback; traceback.print_exc()
        return None
