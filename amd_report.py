"""
판매수수료작업시트(AMD) 생성 모듈
데이터 소스:
  영*판매*.xlsx          → 정상/행사/합계 매출
  사원마스터.xlsx         → 기본수수료, 지급방법
  calc() 결과            → 추가지급수수료, 부가세, 지급할총액 등
  ★경비내역서.xlsx        → 직배비+경비
  아르바이트_정산서.xlsx   → 아르바이트비 (생성 후)
  ◈매장공제건집계.xlsx     → 덜받음, 임의사은품, POS정정공제/환급
  재고실사.xlsx           → 유통하자/LOSS
  급여상여명세서_일용직.xlsx + 급여상여명세서_매장직.xlsx  → 세액공제(별도)
  매장전화요금내역.xlsx    → 전화요금공제 (중간관리만)
  로젠_고객직배*.xlsx     → 직배비공제(V+) (중간관리만, F열×1.1)
"""

import math
import sys
from pathlib import Path
from collections import defaultdict

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

# ── 부서명 → 매장명 매핑 (세액공제용) ─────────────────────────
DEPT_TO_STORE = {
    "롯데잠실점":"롯데잠실점","롯데관악점":"롯데관악점","롯데강남점":"롯데강남점",
    "롯데광주점":"롯데광주점","롯데포항점":"롯데포항점","롯데울산점":"롯데울산점",
    "롯데동래점":"롯데동래점","롯데창원점":"롯데창원점","롯데상인점":"롯데상인점",
    "롯데미아점":"롯데미아점","롯데구리점":"롯데구리점","롯데영등포점":"롯데영등포점",
    "롯데부산본점":"롯데부산본점","롯데평촌점":"롯데평촌점","롯데안산점":"롯데안산점",
    "롯데노원점":"롯데노원점","롯데광복점":"롯데광복점","롯데수원점":"롯데수원점",
    "롯데대구점":"롯데대구점","롯데전주점":"롯데전주점","롯데인천터미널점":"롯데인천터미널점",
    "롯데동탄점":"롯데동탄점","롯데잠실캐슬프라자점":"롯데잠실캐슬프라자점",
    "현대천호점":"현대천호점","현대울산점":"현대울산점","현대압구정본점":"현대압구정본점",
    "현대미아점":"현대미아점","현대중동점":"현대중동점","현대킨텍스점":"현대킨텍스점",
    "현대대구점":"현대대구점","현대충청점":"현대충청점","현대동구점":"현대동구점",
    "현대판교점":"현대판교점","현대신촌점":"현대신촌점","현대울산점":"현대울산점",
    "신세계광주점":"신세계광주점","신세계마산점":"신세계마산점","신세계강남점":"신세계강남점",
    "신세계경기점":"신세계경기점","신세계센텀점":"신세계센텀점",
    "신세계타임스퀘어점":"신세계타임스퀘어점","신세계천안아산점":"신세계천안아산점",
    "신세계의정부점":"신세계의정부점","신세계김해점":"신세계김해점",
    "스타필드하남점":"신세계스타필드하남점","신세계대구점":"신세계대구점",
    "신세계대전점":"신세계대전점","신세계충청점":"신세계천안아산점",
    "갤러리아천안점":"갤러리아천안점","갤러리아광교점":"갤러리아광교점",
    "갤러리아진주점":"갤러리아진주점","갤러리아타임월드점":"갤러리아타임월드점",
    "AK광명점":"AK프라자광명점(직영점)","AK분당점":"AK프라자분당점","AK수원점":"AK프라자수원점",
    "파주점_신세계아울렛":"신세계아울렛파주점","양주점(LF몰)":"LF스퀘어양주점",
    "고양점(롯데아울렛)":"롯데아울렛고양점","광교점(롯데아울렛)":"롯데아울렛광교점",
    "광명점(롯데아울렛)":"롯데아울렛광명점","군산점(롯데아울렛)":"롯데아울렛군산점",
    "동부산점(롯데아울렛)":"롯데아울렛동부산점","이천점(롯데아울렛)":"롯데아울렛이천점",
    "김해점(롯데아울렛)":"롯데아울렛김해점",
    "현대대전점(아울렛)":"현대아울렛대전점","현대남양주점(아울렛)":"현대아울렛남양주점",
    "신세계본점":"신세계본점",
    # ── 직영점 (판매수수료 대상 아님) ──────────────────────────
    "일산점":             "일산점(직영점)",
    "양재점":             "양재점(직영점)",
    "경기광주점":         "경기광주점(직영점)",
    "NC충장점":           "NC충장점(직영점)",
    "NC일산점":           "NC일산점(직영점)",
    "NC해운대점":         "NC해운대점(직영점)",
    "청주점":             "청주점(직영점)",
    "현대커넥트부산점":   "현대커넥트부산점(직영점)",
    "전주점":             "전주점(직영점)",
    "가든5점_현대아울렛": "가든5점(직영점)",
    "NC불광점":           "NC불광점(직영점)",
    # ── 미입점행사 (판매수수료 대상 아님) ──────────────────────
    "ET_월평점":          "ET_월평점(미입점행사)",
}

# ── 직배비공제(V+) 제외 매장 목록 ─────────────────────────────
# 로젠파일에 있더라도 아래 매장은 직배비공제(V+) 적용 안 함
# 매월 제외 매장 추가/삭제 시 이 목록만 수정하면 됩니다
DELIVERY_FEE_EXCLUDE = {
    "롯데창원점",    # 직배비공제V+ 제외
    "롯데대구점",    # 직배비공제V+ 제외
    "신세계마산점",   # 직배비공제V+ 제외
}

# ── 전화요금 B열 약칭 → 정식 매장명 ──────────────────────────
PHONE_NAME_MAP = {
    "롯데관악":"롯데관악점","롯데광주":"롯데광주점","롯데포항":"롯데포항점",
    "롯데동래":"롯데동래점","롯데창원":"롯데창원점","롯데상인":"롯데상인점",
    "롯데미아":"롯데미아점","롯데노원":"롯데노원점","현대천호":"현대천호점",
    "현대울산":"현대울산점","현대미아":"현대미아점","현대중동":"현대중동점",
    "신세계마산":"신세계마산점","신세계경기":"신세계경기점",
    "신세계센텀시티":"신세계센텀점","신세계타임스퀘어":"신세계타임스퀘어점",
    "롯데아울렛 이천점":"롯데아울렛이천점","신세계본점":"신세계본점",
    "신세계아라리오":"신세계천안아산점","광교 갤러리아":"갤러리아광교점",
    "현대대구":"현대대구점","롯데영등포":"롯데영등포점",
    "AK프라자분당":"AK프라자분당점","롯데평촌":"롯데평촌점",
    "롯데광복점":"롯데광복점","현대충청":"현대충청점",
    "천안갤러리아":"갤러리아천안점","롯데안산점":"롯데안산점",
    "롯데김해아울렛":"롯데아울렛김해점","롯데수원":"롯데수원점",
    "신세계김해점":"신세계김해점","신세계하남스타필드점":"신세계스타필드하남점",
    "신세계백화점대구점":"신세계대구점","롯데아울렛고양점":"롯데아울렛고양점",
    "롯데아울렛광교점":"롯데아울렛광교점","롯데대구":"롯데대구점",
    "롯데전주":"롯데전주점","롯데아울렛광명점":"롯데아울렛광명점",
    "LF스퀘어양주점":"LF스퀘어양주점","롯데잠실":"롯데잠실점",
    "롯데아울렛군산점":"롯데아울렛군산점","롯데강남점":"롯데강남점",
    "롯데인천터미널점":"롯데인천터미널점","롯데구리점":"롯데구리점",
    "현대동구점":"현대동구점","현대판교":"현대판교점",
    "롯데부산본점":"롯데부산본점","현대아울렛대전점":"현대아울렛대전점",
    "현대아울렛남양주점":"현대아울렛남양주점","롯데울산점":"롯데울산점",
    "롯데동부산점":"롯데아울렛동부산점","신세계광주":"신세계광주점",
    "롯데동탄점":"롯데동탄점","신세계파주아울렛":"신세계아울렛파주점",
    "현대일산킨텍스점":"현대킨텍스점","신세계강남점":"신세계강남점",
    "신세계대전점":"신세계대전점","신세계의정부점":"신세계의정부점",
    "갤러리아진주점":"갤러리아진주점","갤러리아타임월드점":"갤러리아타임월드점",
    "현대신촌점":"현대신촌점","AK프라자수원":"AK프라자수원점",
}

# ══════════════════════════════════════════════════════════
# 소스 데이터 로더
# ══════════════════════════════════════════════════════════
def load_tax_deduction(paths: list) -> dict:
    """일용직+매장직 급여명세서 → {매장명: 세액공제합계}"""
    data = defaultdict(int)
    for path in paths:
        if not path or not path.exists(): continue
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            dept = str(row[3] or "").strip()
            tax  = row[10]   # 11열(idx10) = 공제총액
            if dept and isinstance(tax, (int,float)) and tax:
                shop = DEPT_TO_STORE.get(dept)
                if shop:
                    data[shop] += int(tax)
    return dict(data)

def load_phone_fee(path: Path, ym: str) -> dict:
    """전화요금내역 → {매장명: 해당월 요금} (중간관리 매장만)"""
    if not path or not path.exists(): return {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    # 1행에서 전월 열 자동 탐지 (AMD는 전월 요금 공제)
    month = int(ym.split("-")[1])
    prev_month = month - 1 if month > 1 else 12
    may_col = None
    for c in range(1, ws.max_column+1):
        v = str(ws.cell(1,c).value or "")
        if f"{prev_month}월" in v:
            may_col = c
            break
    if not may_col: return {}
    data = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        raw = str(row[1] or "").strip()
        amt = row[may_col-1] if may_col else None
        if not raw or "합" in raw: continue
        key = raw.split("(")[0].strip()
        shop = PHONE_NAME_MAP.get(key)
        if shop and isinstance(amt, (int,float)) and amt:
            data[shop] = int(amt)
    return data

def load_delivery_fee(path: Path) -> dict:
    """로젠 고객직배 → {매장명: 신용합계×1.1} (중간관리 매장만)
    운송장번호가 있는 행만 집계 — 없는 행은 소계행이므로 제외
    컬럼 위치는 매달 바뀔 수 있어 헤더 행에서 실제 위치를 찾아 사용"""
    if not path or not path.exists(): return {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    # "로젠택배▶" 포함 시트 탐지
    sheet = next((s for s in wb.sheetnames if "로젠택배" in s), None)
    if not sheet: return {}
    ws = wb[sheet]

    rows_iter = ws.iter_rows(values_only=True)
    header = None
    for row in rows_iter:
        if row and "운송장번호" in row:
            header = row
            break
    if not header:
        return {}
    idx = {name: i for i, name in enumerate(header)}
    try:
        i_track, i_shop, i_credit = idx["운송장번호"], idx["물품옵션"], idx["신용"]
    except KeyError:
        return {}

    data = defaultdict(int)
    for row in rows_iter:
        운송장 = row[i_track]
        shop   = str(row[i_shop] or "").strip()
        amt    = row[i_credit]
        # 운송장번호 없는 행 = 소계행 → 제외
        if not 운송장: continue
        if shop and isinstance(amt, (int,float)):
            data[shop] += int(amt)
    # ×1.1 적용
    return {shop: round(total * 1.1) for shop, total in data.items()
            if shop not in DELIVERY_FEE_EXCLUDE}

def load_arba_subtotals(path: Path) -> dict:
    """아르바이트_정산서 점포별소계(백화점) → {매장명: 총지급액}"""
    if not path or not path.exists(): return {}
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if "점포별소계" not in wb.sheetnames: return {}
        ws = wb["점포별소계"]
        data = {}
        current_cat = None
        for row in ws.iter_rows(min_row=2, values_only=True):
            a = str(row[0] or "").strip()
            b = str(row[1] or "").strip()
            f = row[5]
            if a and a not in ("구분",""):
                current_cat = a
            if current_cat == "백화점" and "소계" in b and isinstance(f,(int,float)):
                shop = b.replace("소계","").strip()
                data[shop] = int(f)
        return data
    except Exception:
        return {}

# ══════════════════════════════════════════════════════════
# 보고서 생성
# ══════════════════════════════════════════════════════════
def make_amd_report(ym: str, base_dir: Path, results: list,
                    sales_detail: dict, expense_rows: dict,
                    master: list) -> Path:
    INPUT  = base_dir / "input"
    OUTPUT = base_dir / "output" / ym
    OUTPUT.mkdir(parents=True, exist_ok=True)

    # ── 추가 소스 파일 탐지 ──────────────────────────────
    def find_in(*patterns):
        for pat in patterns:
            found = list(INPUT.glob(pat))
            if found: return found[0]
        return None

    arba_path = OUTPUT / "아르바이트_정산서.xlsx"
    tax_paths = [
        find_in("급여상여명세서_일용직*.xlsx", "급여상여명세서*일용직*.xlsx"),
        find_in("급여상여명세서_매장직*.xlsx", "급여상여명세서*매장직*.xlsx"),
    ]
    phone_path    = find_in("매장전화요금내역.xlsx", "*전화요금*.xlsx")
    delivery_path = find_in("*로젠*고객직배*.xlsx", "*고객직배*.xlsx")

    print(f"     세액공제 파일: {[p.name if p else '없음' for p in tax_paths]}")
    print(f"     전화요금 파일: {phone_path.name if phone_path else '없음'}")
    print(f"     직배비 파일:   {delivery_path.name if delivery_path else '없음'}")

    # ── 데이터 로드 ──────────────────────────────────────
    tax_data  = load_tax_deduction(tax_paths)
    phone_data = load_phone_fee(phone_path, ym)
    deliv_data = load_delivery_fee(delivery_path)
    arba_data  = load_arba_subtotals(arba_path)

    # 중간관리 매장 목록
    mgr_shops = {e["shop"] for e in master
                 if (e.get("pay_type","") or "") == "중간관리"}

    # 사원마스터에서 calc 결과 병합: emp당 1행
    # results = [{"emp":{...},"res":{...}}]
    # 매장+성명 조합으로 유일 행 구성
    # 매장별 전체 수수료 합산 (대표 사원 행에 집계)
    shop_total_fee = defaultdict(int)   # {shop: 전체사원 판매수수료합계}
    shop_total_vat = defaultdict(int)   # {shop: 전체사원 부가세합계}
    for r in results:
        emp, res = r["emp"], r["res"]
        if not emp["name"]: continue
        shop = emp["shop"]
        shop_total_fee[shop] += emp.get("base_fee",0) + res.get("commission",0)
        shop_total_vat[shop] += res.get("vat",0)

    # 매장별 대표 사원 판단 (중간관리=매니저, 본사지급=본사-M)
    shop_repr = {}  # {shop: emp}
    for r in results:
        emp = r["emp"]
        if not emp["name"]: continue
        shop = emp["shop"]
        grade = emp.get("grade","")
        pay_type = emp.get("pay_type","")
        if shop not in shop_repr:
            shop_repr[shop] = r
        else:
            prev = shop_repr[shop]["emp"]
            # 매니저 또는 본사-M 우선
            if grade in ("매니저","본사-M"):
                shop_repr[shop] = r

    person_rows = []
    processed_shops = set()
    for r in results:
        emp, res = r["emp"], r["res"]
        if not emp["name"]: continue
        shop = emp["shop"]
        is_repr = (shop_repr.get(shop) is r)   # 대표 사원 여부
        er   = (expense_rows or {}).get(shop, {})
        sd   = (sales_detail or {}).get(shop, {})

        # 매출 (대표 사원만 표시)
        off_n = sd.get("off_normal", 0) if is_repr else 0
        off_e = sd.get("off_event",  0) if is_repr else 0
        on_n  = sd.get("on_normal",  0) if is_repr else 0
        on_e  = sd.get("on_event",   0) if is_repr else 0
        정상매출 = off_n + on_n
        행사매출 = off_e + on_e
        매출합계 = 정상매출 + 행사매출

        # 수수료: 개인 기본+추가
        기본수수료    = emp.get("base_fee", 0)
        추가지급수수료  = res.get("commission", 0)
        판매수수료합계  = 기본수수료 + 추가지급수수료

        # 지원금/공제/경비는 대표 사원 행에만 집계
        if is_repr:
            직배비경비  = er.get("j_direct", 0)
            아르바이트비 = arba_data.get(shop, 0)
            덜받음   = er.get("shortfall", 0)
            사은품   = er.get("gift", 0)
            POS공제  = er.get("pos", 0)
            POS환급  = er.get("t_refund", 0)
            LOSS    = er.get("loss", 0)
            공제계   = 덜받음 + 사은품 + POS공제 - POS환급 + LOSS
            # 공급가액 = 매니저수수료(10원올림) + 직배비 - 공제 (아르바이트 제외)
            # 공급가액 공식: + 매니저수수료 + 지원금(직배비·경비)
            #               - 덜받음 - 사은품 - POS공제 + POS환급 - LOSS
            _fee_ceil = math.ceil(shop_total_fee[shop] / 10) * 10
            공급가액  = _fee_ceil + 직배비경비 - 공제계
            # 부가세 = 공급가액 × 10% (사업소득자 매장만, 10원 미만 올림 전 기준)
            is_biz = any(r["emp"]["income"]=="사업소득"
                        for r in results if r["emp"]["shop"]==shop and r["emp"]["name"])
            부가세   = int(공급가액 * 0.1) if is_biz else 0
            지급할총액 = 공급가액 + 부가세
            전화요금공제 = phone_data.get(shop, 0) if shop in mgr_shops else 0
            세액공제   = tax_data.get(shop, 0)
            직배비공제vp = deliv_data.get(shop, 0) if shop in mgr_shops else 0
            공제합계   = 전화요금공제 + 세액공제 + 직배비공제vp
            송금액 = 지급할총액 - 공제합계
        else:
            직배비경비=0; 아르바이트비=0; 덜받음=0; 사은품=0
            POS공제=0; POS환급=0; LOSS=0; 공제계=0
            공급가액=0; 부가세=0; 지급할총액=0
            전화요금공제=0; 세액공제=0; 직배비공제vp=0; 공제합계=0; 송금액=0

        pay_type = emp.get("pay_type", "")
        person_rows.append({
            "매장명":     shop,
            "성명":      emp["name"],
            "지급방법":   pay_type,
            "정상매출":   정상매출,
            "행사매출":   행사매출,
            "매출합계":   매출합계,
            "기본수수료":  기본수수료,
            "추가지급수수료": 추가지급수수료,
            "판매수수료합계": 판매수수료합계,
            "직배비+경비": 직배비경비,
            "아르바이트비": 아르바이트비,
            "①공급가액":  공급가액,
            "②부가세":   부가세,
            "③지급할총액": 지급할총액,
            "덜받음":    덜받음,
            "임의사은품지급": 사은품,
            "POS정정요청공제": POS공제,
            "POS정정요청공제환급금": POS환급,
            "유통하자/재고LOSS": LOSS,
            "전화요금공제": 전화요금공제,
            "세액공제(별도)": 세액공제,
            "직배비공제(V+)": 직배비공제vp,
            "ⓕ공제합계":  공제합계,
            "ⓗ송금액":   송금액,
        })

    # ── 엑셀 생성 ────────────────────────────────────────
    NAVY="1F3864"; BLUE="2E5EAA"; PALE="EEF3FB"; WHITE="FFFFFF"
    GREEN="1E6B3C"; GRAY="F2F2F2"; AMBER="FFF2CC"; RED="C00000"
    FMT="#,##0"

    ITEMS=[
        ("매장명","매장명","left",None,22),
        ("성명","성명","center",None,8),
        ("지급방법","지급방법","center",None,8),
        ("정상매출","정상매출","center",FMT,12),
        ("행사매출","행사매출","center",FMT,12),
        ("매출합계","매출합계","center",FMT,12),
        ("기본수수료","기본수수료","center",FMT,12),
        ("추가지급수수료","추가지급\n수수료","center",FMT,10),
        ("판매수수료합계","판매수수료\n합계","center",FMT,12),
        ("직배비+경비","직배비+경비","center",FMT,10),
        ("아르바이트비","아르바이트비","center",FMT,10),
        ("①공급가액","①공급가액","center",FMT,12),
        ("②부가세","②부가세","center",FMT,10),
        ("③지급할총액","③지급할\n총액","center",FMT,12),
        ("덜받음","덜받음","center",FMT,10),
        ("임의사은품지급","임의사은품\n지급","center",FMT,10),
        ("POS정정요청공제","POS정정\n요청공제","center",FMT,10),
        ("POS정정요청공제환급금","POS정정\n환급금","center",FMT,10),
        ("유통하자/재고LOSS","유통하자\n/LOSS","center",FMT,10),
        ("전화요금공제","전화요금\n공제","center",FMT,10),
        ("세액공제(별도)","세액공제\n(별도)","center",FMT,10),
        ("직배비공제(V+)","직배비공제\n(V+)","center",FMT,10),
        ("ⓕ공제합계","ⓕ공제합계","center",FMT,10),
        ("ⓗ송금액","ⓗ송금액","center",FMT,12),
    ]
    GROUPS=[
        (1,3,"기본정보",NAVY),(4,6,"매출",BLUE),(7,9,"수수료",BLUE),
        (10,11,"지원금","375623"),(12,14,"지급금액","4472C4"),
        (15,23,"공제항목",RED),(24,24,"송금액","276221"),
    ]

    wb = openpyxl.Workbook(); ws = wb.active
    ws.title="판매수수료작업시트(AMD)"
    ws.sheet_view.showGridLines=False; ws.freeze_panes="D3"

    def st(r,c,v=None,bg=None,fg="000000",bold=False,size=9,
           align="center",wrap=False,fmt=None):
        cell=ws.cell(r,c)
        if v is not None: cell.value=v
        cell.font=Font(name="맑은 고딕",bold=bold,color=fg,size=size)
        cell.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap)
        if bg: cell.fill=PatternFill("solid",fgColor=bg)
        if fmt: cell.number_format=fmt
        return cell

    def bdr(r1,c1,r2,c2):
        tk=Side(style="medium"); tn=Side(style="thin")
        for r in range(r1,r2+1):
            for c in range(c1,c2+1):
                ws.cell(r,c).border=Border(
                    left=tk if c==c1 else tn, right=tk if c==c2 else tn,
                    top=tk if r==r1 else tn, bottom=tk if r==r2 else tn)

    ws.merge_cells(f"A1:{get_column_letter(len(ITEMS))}1")
    st(1,1,f"판매수수료 작업시트 (AMD)  ▶  {ym}",
       bg=NAVY,fg=WHITE,bold=True,size=13)
    ws.row_dimensions[1].height=28

    for c1,c2,lbl,bg in GROUPS:
        if c1!=c2: ws.merge_cells(f"{get_column_letter(c1)}2:{get_column_letter(c2)}2")
        st(2,c1,lbl,bg=bg,fg=WHITE,bold=True,size=8)
    ws.row_dimensions[2].height=14

    for ci,(key,hdr,align,fmt,w) in enumerate(ITEMS,1):
        ws.column_dimensions[get_column_letter(ci)].width=w
        bg=next((color for s,e,_,color in GROUPS if s<=ci<=e),NAVY)
        st(3,ci,hdr,bg=bg,fg=WHITE,bold=True,size=8,wrap=True)
    ws.row_dimensions[3].height=30

    prev_shop=""; bg_toggle=False
    for ei,p in enumerate(person_rows):
        row=ei+4
        shop=str(p.get("매장명") or "")
        pay=str(p.get("지급방법") or "")
        송금액=p.get("ⓗ송금액") or 0
        if shop!=prev_shop:
            bg_toggle=not bg_toggle if prev_shop else False
            prev_shop=shop
        rb=WHITE if not bg_toggle else GRAY

        for ci,(key,hdr,align,fmt,w) in enumerate(ITEMS,1):
            v=p.get(key)
            if v is None: v=0 if fmt else ""
            bg=rb; fg="000000"; bold=False

            if key=="지급방법":
                if pay=="중간관리": bg=AMBER; fg="7F4F00"; bold=True
                elif pay=="본사지급": bg=PALE; fg=NAVY
            elif key=="③지급할총액": bg="E2EFDA"; bold=True
            elif key=="ⓗ송금액":
                if isinstance(송금액,(int,float)) and 송금액<0:
                    bg="FFE0E0"; fg=RED; bold=True
                else:
                    bg="C6EFCE"; fg="276221"; bold=True

            if isinstance(v,(int,float)) and v<0 and key not in ("ⓗ송금액","③지급할총액"):
                fg=RED

            c=ws.cell(row,ci,v)
            c.font=Font(name="맑은 고딕",bold=bold,color=fg,size=9)
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Alignment(horizontal=align,vertical="center")
            if fmt and isinstance(v,(int,float)): c.number_format=fmt
        ws.row_dimensions[row].height=17

    sr=len(person_rows)+4
    ws.merge_cells(f"A{sr}:C{sr}")
    st(sr,1,f"합  계  ({len(person_rows)}명)",bg=NAVY,fg=WHITE,bold=True,size=10)
    for ci,(key,hdr,align,fmt,w) in enumerate(ITEMS,1):
        if not fmt: continue
        total=sum((p.get(key) or 0) for p in person_rows if isinstance(p.get(key),(int,float)))
        bg="C6EFCE" if key=="ⓗ송금액" else ("E2EFDA" if key=="③지급할총액" else NAVY)
        fg="276221" if key=="ⓗ송금액" else WHITE
        c=ws.cell(sr,ci,total)
        c.font=Font(name="맑은 고딕",bold=True,color=fg,size=10)
        c.fill=PatternFill("solid",fgColor=bg); c.number_format=FMT
        c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[sr].height=22
    bdr(2,1,sr,len(ITEMS))

    out_path = OUTPUT / "판매수수료작업시트_AMD.xlsx"
    wb.save(out_path)
    return out_path


def run(ym: str, base_dir: Path, results: list,
        sales_detail: dict, expense_rows: dict, master: list):
    print(f"  📋 판매수수료작업시트(AMD) 처리 중...")
    try:
        path = make_amd_report(ym, base_dir, results,
                               sales_detail, expense_rows, master)
        print(f"     ✅ 판매수수료작업시트_AMD.xlsx  ({sum(1 for r in results if r['emp']['name'])}명)")
        return path
    except Exception as e:
        print(f"     ⚠️  AMD 처리 오류: {e}")
        import traceback; traceback.print_exc()
        return None
