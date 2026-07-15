"""
판매수수료집계.xlsx 생성 모듈
템플릿: salary_system/templates/판매수수료집계_템플릿.xlsx
집계 항목:
  매출: 정상/행사/합계
  판매수수료: 근로소득/중간관리/합계
  지원금: 직배비+경비
  공제: 덜받음/임의사은품/POS공제/POS환급/유통하자LOSS
  아르바이트비
  공급가액 = ①+②-③+④
  부가세합계
  기타공제: 전화요금/직배비공제(V+)/세액공제(별도)
  송금액
"""
import sys
from pathlib import Path
from collections import defaultdict

from common import find_input_one

try:
    import openpyxl
    from openpyxl import load_workbook
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable,"-m","pip","install","openpyxl","--quiet"])
    import openpyxl
    from openpyxl import load_workbook


def make_sales_summary(ym: str, base_dir: Path,
                       results: list,
                       sales_detail: dict,
                       expense_rows: dict,
                       master: list,
                       pos_unit: int = 0,
                       event_shops: set = None) -> Path:
    """
    results       : calc() 결과 리스트
    sales_detail  : load_sales_file() 결과 {shop: {off_normal,...}}
    expense_rows  : {shop: {j_direct, shortfall, loss, gift, pos, t_refund}}
    master        : 사원마스터 리스트
    pos_unit      : POS환급 단가
    event_shops   : 행사매장 set
    """
    INPUT  = base_dir / "input" / ym
    OUTPUT = base_dir / "output" / ym
    OUTPUT.mkdir(parents=True, exist_ok=True)
    TMPL   = base_dir / "templates" / "판매수수료집계_템플릿.xlsx"

    if not TMPL.exists():
        print(f"     ⚠️  템플릿 없음: {TMPL}")
        return None

    if event_shops is None:
        event_shops = set()

    # ── 소스 데이터 로드 (정산월 폴더 안에서만 찾음) ─────────────────
    def _find_in(patterns):
        return find_input_one(patterns, base=INPUT)

    try:
        from amd_report import (load_phone_fee, load_tax_deduction,
                                load_delivery_fee, load_arba_subtotals)
        phone_path  = _find_in(["매장전화요금내역*.xlsx", "*전화요금*.xlsx"])
        tax_paths   = [
            _find_in(["급여상여명세서*일용직*.xlsx"]),
            _find_in(["급여상여명세서*매장직*.xlsx"]),
        ]
        deliv_path  = _find_in(["*로젠*고객직배*.xlsx", "*고객직배*.xlsx"])
        arba_path   = OUTPUT / "아르바이트_정산서.xlsx"
    except Exception as e:
        print(f"     ⚠️  보조 데이터 파일 탐색 오류: {e}")
        phone_path = deliv_path = None
        tax_paths  = []
        arba_path  = OUTPUT / "아르바이트_정산서.xlsx"

    try:
        phone_data = load_phone_fee(phone_path, ym) if phone_path else {}
    except Exception as e:
        print(f"     ⚠️  전화요금 데이터 로드 오류: {e}")
        phone_data = {}
    try:
        tax_data = load_tax_deduction(tax_paths)
    except Exception as e:
        print(f"     ⚠️  세액공제 데이터 로드 오류: {e}")
        tax_data = {}
    try:
        deliv_data = load_delivery_fee(deliv_path) if deliv_path else {}
    except Exception as e:
        print(f"     ⚠️  직배비 데이터 로드 오류: {e}")
        deliv_data = {}
    try:
        arba_data = load_arba_subtotals(arba_path)
    except Exception as e:
        print(f"     ⚠️  아르바이트 데이터 로드 오류: {e}")
        arba_data = {}

    mgr_shops = {e["shop"] for e in master if e.get("pay_type","") == "중간관리"}

    # ── 집계 계산 ─────────────────────────────────────────────
    # 매출
    정상매출  = sum(d.get("off_normal",0)+d.get("on_normal",0) for d in sales_detail.values())
    행사매출  = sum(d.get("off_event", 0)+d.get("on_event", 0) for d in sales_detail.values())

    # 수수료
    근로수수료 = sum(r["res"]["base_fee"]+r["res"]["commission"]
                  for r in results
                  if r["emp"]["income"]=="근로소득" and r["emp"]["name"])
    중간관리수수료 = sum(r["res"]["base_fee"]+r["res"]["commission"]
                     for r in results
                     if r["emp"].get("pay_type","")=="중간관리" and r["emp"]["name"])
    수수료합계 = 근로수수료 + 중간관리수수료

    # 이번 달 매출이 있는 매장만 집계 대상 (매출 0/데이터 없음 → 완전 제외)
    sales_shops = {s for s, d in sales_detail.items() if d.get("grand_total", 0)}

    # 지원금/공제
    from expense_report import STORE_CODE_MAP
    _scm_sales  = [s for s in STORE_CODE_MAP if s in sales_shops]
    직배비합계  = sum(expense_rows.get(s,{}).get("j_direct",  0) for s in _scm_sales)
    덜받음합계  = sum(expense_rows.get(s,{}).get("shortfall", 0) for s in _scm_sales)
    사은품합계  = sum(expense_rows.get(s,{}).get("gift",      0) for s in _scm_sales)
    POS공제합계 = sum(expense_rows.get(s,{}).get("pos",       0) for s in _scm_sales)
    POS환급합계 = sum(expense_rows.get(s,{}).get("t_refund",  0) for s in _scm_sales)
    LOSS합계    = sum(expense_rows.get(s,{}).get("loss",      0) for s in _scm_sales)
    공제합계    = 덜받음합계 + 사은품합계 + POS공제합계 - POS환급합계 + LOSS합계

    # 아르바이트 (백화점 소계 + 월별입력)
    # 아르바이트: 모든 백화점 소계 합산 (행사매장 제외 없음)
    아르바이트합계 = sum(arba_data.values()) if arba_data else 0

    # 공급가액
    공급가액 = 수수료합계 + 직배비합계 - 공제합계 + 아르바이트합계

    # 부가세: 매장별 공급가액 × 10% (중간관리 + 사업소득 매장만, 세금계산서 발행 대상과 동일 조건)
    # 공급가액(매장별) = 수수료(10원올림) + 직배비 - 공제
    # → AMD와 동일한 방식으로 계산
    import math as _math2
    from collections import defaultdict as _dd2
    _shop_fee = _dd2(int)
    _shop_income = {}
    for _r in results:
        _e, _res = _r["emp"], _r["res"]
        if not _e["name"]: continue
        _shop_fee[_e["shop"]] += _e["base_fee"] + _res["commission"]
        if _e["shop"] not in _shop_income:
            _shop_income[_e["shop"]] = _e["income"]
    부가세합계 = 0
    for _shop, _inc in _shop_income.items():
        if _inc != "사업소득" or _shop not in mgr_shops: continue
        _er = expense_rows.get(_shop, {})
        _fee_c = _math2.ceil(_shop_fee[_shop] / 10) * 10
        _공제 = (_er.get("shortfall",0)+_er.get("gift",0)+_er.get("pos",0)
                 -_er.get("t_refund",0)+_er.get("loss",0))
        _공급 = _fee_c + _er.get("j_direct",0) - _공제
        부가세합계 += int(_공급 * 0.1)

    # 총지급액
    총지급액 = 공급가액 + 부가세합계

    # 기타공제
    전화요금합계    = sum(phone_data.get(s,0) for s in mgr_shops)
    직배비공제합계  = sum(deliv_data.get(s,0) for s in mgr_shops)
    # 세액공제: 매출 있는 매장 기준 (POS환급 외 행사매장 제외 없음)
    세액공제합계    = sum(v for shop,v in tax_data.items()
                        if shop in sales_shops)
    기타공제합계    = 전화요금합계 + 직배비공제합계 + 세액공제합계

    # 송금액
    송금액 = 총지급액 - 기타공제합계

    # 구분 집계
    사업소득자수 = len({r["emp"]["shop"] for r in results if r["emp"]["income"]=="사업소득"})
    근로소득자수 = len({r["emp"]["shop"] for r in results if r["emp"]["income"]=="근로소득"})
    # 사업소득자 송금액 / 근로소득자 송금액 추정 (비례)
    사업소득자수total = 사업소득자수 + 근로소득자수

    # ── 템플릿에 값 주입 ─────────────────────────────────────
    wb = load_workbook(TMPL)
    ws = wb.active
    FMT = "#,##0"

    def sv(r, v):
        cell = ws.cell(r, 7)
        cell.value = v if v else 0
        cell.number_format = FMT

    ws.cell(2,2).value = f"{ym[:4]}년 {int(ym[5:])}월 판매수수료 내역 집계"
    ws.cell(4,5).value = f"{ym[:4]}년  {int(ym[5:])+1}월  10일"
    ws.cell(4,5).number_format = "@"   # 텍스트로

    sv(6,  정상매출)
    sv(7,  행사매출)
    sv(8,  정상매출 + 행사매출)
    sv(10, 근로수수료)
    sv(11, 중간관리수수료)
    sv(12, 수수료합계)
    sv(13, 직배비합계)
    sv(14, 덜받음합계)
    sv(15, 사은품합계)
    sv(16, POS공제합계)
    sv(17, POS환급합계)     # POS환급은 양수로 표시 (공급가액 계산식에서 이미 차감)
    sv(18, LOSS합계)
    sv(19, 공제합계)
    sv(20, 아르바이트합계)
    sv(21, 공급가액)
    sv(22, 부가세합계)
    sv(23, 총지급액)
    sv(24, 전화요금합계)
    sv(25, 직배비공제합계)
    sv(27, 세액공제합계)
    sv(28, 기타공제합계)
    sv(29, 송금액)
    # ── 2. 구분 섹션: 매장별 송금액 집계 ──────────────────────────
    import math as _math
    from collections import defaultdict as _dd
    from expense_report import STORE_CODE_MAP as _SCM3

    # 위에서 이미 로드한 데이터 재사용 (중복 로드 방지)
    _phone  = phone_data
    _tax    = tax_data
    _deliv  = deliv_data
    _arba2  = arba_data
    _results_map = {r["emp"]["shop"]: r for r in results if r["emp"]["name"]}
    _mgr_shops2  = mgr_shops
    _event2      = {e["shop"] for e in master if (e.get("note","") or "").strip()=="행사매장"}

    shop_송금  = _dd(int)
    shop_구분  = {}
    processed  = set()

    for r in results:
        emp, res = r["emp"], r["res"]
        if not emp["name"]: continue
        shop = emp["shop"]
        if shop in processed: continue
        processed.add(shop)

        er_shop = expense_rows.get(shop, {})
        mgr_fee_s = _math.ceil(
            sum((rr["res"]["commission"]+rr["res"]["base_fee"])
                for rr in results if rr["emp"]["shop"]==shop and rr["emp"]["name"])
            / 10) * 10
        j_dir   = er_shop.get("j_direct",0)
        공제계_s = (er_shop.get("shortfall",0)+er_shop.get("gift",0)
                    +er_shop.get("pos",0)-er_shop.get("t_refund",0)+er_shop.get("loss",0))
        공급가액_s = mgr_fee_s + j_dir - 공제계_s
        # 부가세: 중간관리 + 사업소득 매장만 (세금계산서 발행 대상과 동일 조건)
        _shop_inc = next((rr["emp"]["income"] for rr in results
                          if rr["emp"]["shop"]==shop and rr["emp"]["name"]), "")
        부가세_s   = int(공급가액_s * 0.1) if (_shop_inc == "사업소득" and shop in _mgr_shops2) else 0
        지급할총액_s= 공급가액_s + 부가세_s
        tel   = _phone.get(shop,0) if shop in _mgr_shops2 else 0
        deli  = _deliv.get(shop,0) if shop in _mgr_shops2 else 0
        tax_s = _tax.get(shop,0)   if shop in _SCM3 else 0
        기타_s= tel + deli + tax_s
        송금_s = 지급할총액_s - 기타_s
        # 아르바이트 가산
        송금_s += _arba2.get(shop, 0)
        shop_송금[shop] = 송금_s

        # 구분
        if shop in _event2:
            shop_구분[shop] = "행사매장"
        elif emp.get("pay_type","") == "중간관리":
            shop_구분[shop] = "백화점판매대행계약점"
        else:
            shop_구분[shop] = "본사지급"

    # 행사매장(name=None)의 아르바이트도 합산 — 대표매장 없어서 루프에서 처리 안됨
    for _shop, _amt in _arba2.items():
        if _shop in _event2 and _shop not in processed:
            # 아르바이트 가산, 세액공제 차감 적용
            _tax_ev = _tax.get(_shop, 0) if _shop in _SCM3 else 0
            shop_송금[_shop] += _amt - _tax_ev
            shop_구분[_shop] = "행사매장"

    구분별_송금  = _dd(int)
    구분별_매장수 = _dd(int)
    for shop, 구분 in shop_구분.items():
        구분별_송금[구분]  += shop_송금[shop]
        구분별_매장수[구분] += 1

    total_매장 = len(shop_구분)
    total_송금2 = sum(shop_송금.values())

    # 행사매장 송금액이 있을 때만 행 추가
    행사_있음 = 구분별_매장수.get("행사매장", 0) > 0 and 구분별_송금.get("행사매장",0) != 0

    def _sv2(r, c, v, fmt=FMT):
        cell = ws.cell(r, c)
        cell.value = v if v is not None else 0
        if fmt: cell.number_format = fmt

    # 32행: 실송금액 합계
    _sv2(32, 6, total_매장)
    _sv2(32, 7, total_송금2)

    # 33행: 백화점판매대행계약점
    _sv2(33, 6, 구분별_매장수.get("백화점판매대행계약점",0))
    _sv2(33, 7, 구분별_송금.get("백화점판매대행계약점",0))

    # 34행: 본사지급
    _sv2(34, 6, 구분별_매장수.get("본사지급",0))
    _sv2(34, 7, 구분별_송금.get("본사지급",0))

    # 35행: 행사매장 (매출 있을 때만)
    if 행사_있음:
        ws.cell(35,2).value = "- "
        ws.cell(35,3).value = "행사매장"
        # 병합셀 C35:E35
        try:
            ws.merge_cells("C35:E35")
        except Exception:
            pass
        import copy
        # 스타일 복사 (34행 기준)
        for c in range(1,8):
            src_cell = ws.cell(34, c)
            dst_cell = ws.cell(35, c)
            if src_cell.has_style:
                dst_cell.font      = copy.copy(src_cell.font)
                dst_cell.fill      = copy.copy(src_cell.fill)
                dst_cell.border    = copy.copy(src_cell.border)
                dst_cell.alignment = copy.copy(src_cell.alignment)
        _sv2(35, 6, 구분별_매장수.get("행사매장",0))
        _sv2(35, 7, 구분별_송금.get("행사매장",0))

    # 인쇄 영역: 템플릿의 고정값(B2:G34)이 아니라 실제 마지막 데이터 행 기준으로 설정
    # (행사매장 행이 있을 때 35행까지 늘어나므로, 그렇지 않으면 마지막 줄이 잘림)
    last_row = 35 if 행사_있음 else 34
    ws.print_area = f"B2:G{last_row}"

    out_path = OUTPUT / "판매수수료집계.xlsx"
    wb.save(out_path)
    return out_path


def run(ym: str, base_dir: Path, results: list,
        sales_detail: dict, expense_rows: dict,
        master: list, pos_unit: int = 0, event_shops: set = None):
    print(f"  📊 판매수수료집계 처리 중...")
    try:
        path = make_sales_summary(ym, base_dir, results, sales_detail,
                                  expense_rows, master, pos_unit, event_shops)
        if path:
            print(f"     ✅ 판매수수료집계.xlsx")
        return path
    except Exception as e:
        print(f"     ⚠️  판매수수료집계 오류: {e}")
        import traceback; traceback.print_exc()
        return None
