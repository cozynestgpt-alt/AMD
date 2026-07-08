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

from common import STORE_CODE_MAP

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
    "김해점(롯데아울렛)":"롯데아울렛김해점","롯데아울렛서울역점":"롯데아울렛서울역점",
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

# ── 로젠 파일 물품옵션(매장명) 약칭 → 정식 매장명 매핑 ────────
# (직영점/거래처 표기는 원래 매칭 대상이 아니므로 여기 넣지 않음)
DELIVERY_NAME_MAP = {
    "현대울산": "현대울산점",
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
    """일용직+매장직 급여명세서 → {매장명: 세액공제합계}
    두 파일은 컬럼 배치가 서로 달라 헤더에서 "부서"/"공제총액" 위치를 파일별로 찾아 사용"""
    data = defaultdict(int)
    for path in paths:
        if not path or not path.exists(): continue
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        header = next(rows_iter, None)
        if not header:
            continue
        try:
            i_dept = header.index("부서")
            i_tax  = header.index("공제총액")
        except ValueError:
            continue
        for row in rows_iter:
            dept = str(row[i_dept] or "").strip()
            tax  = row[i_tax]
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
    """로젠 고객직배 → {매장명: (신용+제주운임/산간료)합계×1.1}
    - 운송장번호가 있는 행만 집계 (없는 행 = 소계행 → 제외)
    - 집배구분="요청반품" & 물품명에 "입금완료" 포함 → 이미 정산된 반품건이므로 제외
    - 중간관리 매장 필터링은 호출부에서 처리 (mgr_shops 체크)
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

    def find_col(*keywords):
        for i, h in enumerate(header):
            if h and all(k in str(h) for k in keywords):
                return i
        return None

    i_cat    = find_col("집배구분")
    i_track  = find_col("운송장번호")
    i_shop   = find_col("물품옵션")
    i_item   = find_col("물품명")
    i_credit = find_col("신용")
    i_jeju   = find_col("제주")
    if None in (i_cat, i_track, i_shop, i_item, i_credit, i_jeju):
        return {}

    data = defaultdict(int)
    for row in rows_iter:
        운송장 = row[i_track]
        if not 운송장: continue                                  # 소계행 제외
        cat  = row[i_cat]
        item = str(row[i_item] or "")
        if cat == "요청반품" and "입금완료" in item: continue        # 정산완료 반품건 제외

        shop   = str(row[i_shop] or "").strip()
        shop   = DELIVERY_NAME_MAP.get(shop, shop)
        credit = row[i_credit] if isinstance(row[i_credit], (int, float)) else 0
        jeju   = row[i_jeju]   if isinstance(row[i_jeju],   (int, float)) else 0
        amt = credit + jeju
        if shop and amt:
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
    # scoped: 연월(ym_compact) 또는 연+월을 동시에 요구하는 패턴(다른 달과 안 겹침)
    #         → input/ 최상위 우선, 없으면 input/이전/까지 재귀 탐색 (과거 달 재조회 지원)
    # loose : 연월 정보 없는 느슨한 패턴 → 최상위에서만, scoped가 전부 실패했을 때 최후 수단
    def find_in(scoped, loose=()):
        for pat in scoped:
            found = sorted(INPUT.glob(pat))
            if found: return found[0]
        for pat in scoped:
            found = sorted(INPUT.rglob(pat))
            if found: return found[0]
        for pat in loose:
            found = sorted(INPUT.glob(pat))
            if found: return found[0]
        return None

    _ym_compact = ym.replace("-", "")
    _yy, _mm_pad = ym[2:4], ym[5:]
    _mm_bare = str(int(_mm_pad))

    arba_path = OUTPUT / "아르바이트_정산서.xlsx"
    tax_paths = [
        find_in([f"급여상여명세서_일용직*{_ym_compact}*.xlsx"], ["급여상여명세서_일용직*.xlsx", "급여상여명세서*일용직*.xlsx"]),
        find_in([f"급여상여명세서_매장직*{_ym_compact}*.xlsx"], ["급여상여명세서_매장직*.xlsx", "급여상여명세서*매장직*.xlsx"]),
    ]
    phone_path    = find_in([], ["매장전화요금내역.xlsx", "*전화요금*.xlsx"])
    delivery_path = find_in(
        [f"*{_yy}.{_mm_pad}월*로젠*고객직배*.xlsx", f"*{_yy}.{_mm_bare}월*로젠*고객직배*.xlsx",
         f"*{_yy}.{_mm_pad}월*고객직배*.xlsx", f"*{_yy}.{_mm_bare}월*고객직배*.xlsx"],
        ["*로젠*고객직배*.xlsx", "*고객직배*.xlsx"])

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

    # 매장별 사원(results) 목록
    shop_emps = defaultdict(list)   # {shop: [r, ...]}
    for r in results:
        if not r["emp"]["name"]: continue
        shop_emps[r["emp"]["shop"]].append(r)

    # ── 매장별 1행 집계 ──────────────────────────────────
    store_rows = []
    for shop, code in sorted(STORE_CODE_MAP.items(), key=lambda x: x[1]):
        emps = shop_emps.get(shop, [])
        er   = (expense_rows or {}).get(shop, {})
        sd   = (sales_detail or {}).get(shop, {})

        # 대표자 이름: 매니저(중간관리) 우선, 없으면 본사-M
        대표자 = next((r["emp"]["name"] for r in emps
                     if r["emp"].get("grade","")=="매니저"), "")
        if not 대표자:
            대표자 = next((r["emp"]["name"] for r in emps
                         if r["emp"].get("grade","")=="본사-M"), "")

        매출합계 = sd.get("grand_total", 0)

        def _pay(r):
            return r["emp"].get("base_fee",0) + r["res"].get("commission",0)

        본사M  = sum(_pay(r) for r in emps if r["emp"].get("grade","")=="본사-M")
        본사S1 = sum(_pay(r) for r in emps if r["emp"].get("grade","")=="본사-S1")
        중간관리수수료 = sum(_pay(r) for r in emps
                        if r["emp"].get("grade","")=="매니저"
                        and r["emp"].get("pay_type","")=="중간관리")

        판매수수료합계 = 본사M + 본사S1 + 중간관리수수료   # H = SUM(E:G)

        직배비경비  = er.get("j_direct", 0)
        아르바이트비 = arba_data.get(shop, 0)
        덜받음   = er.get("shortfall", 0)
        사은품   = er.get("gift", 0)
        POS공제  = er.get("pos", 0)
        POS환급  = er.get("t_refund", 0)
        LOSS    = er.get("loss", 0)

        # 공급가액 = SUM(H:J) - SUM(K:M,O) + N
        #        = (판매수수료합계+직배비경비+아르바이트비) - (덜받음+임의사은품+POS정정공제+유통LOSS) + POS정정환급금
        공급가액  = (판매수수료합계 + 직배비경비 + 아르바이트비) \
                  - (덜받음 + 사은품 + POS공제 + LOSS) + POS환급
        # 부가세 = 공급가액 × 10% (사업소득자 매장만)
        is_biz = any(r["emp"]["income"]=="사업소득" for r in emps)
        부가세   = int(공급가액 * 0.1) if is_biz else 0
        지급할총액 = 공급가액 + 부가세   # R = SUM(P:Q)

        is_mgr_shop = shop in mgr_shops
        전화요금공제 = phone_data.get(shop, 0) if is_mgr_shop else 0
        세액공제   = tax_data.get(shop, 0) if not is_mgr_shop else 0   # 본사지급 매장만
        직배비공제vp = deliv_data.get(shop, 0) if is_mgr_shop else 0
        기타공제합계 = 전화요금공제 + 세액공제 + 직배비공제vp   # V = SUM(S:U)
        송금액 = 지급할총액 - 기타공제합계                        # W = R - V

        store_rows.append({
            "매장코드":   code,
            "매장명":     shop,
            "매니저":     대표자,
            "매출합계":   매출합계,
            "본사-M":    본사M,
            "본사-S1":   본사S1,
            "중간관리수수료": 중간관리수수료,
            "직배비+경비": 직배비경비,
            "아르바이트비": 아르바이트비,
            "덜받음":    덜받음,
            "임의사은품지급": 사은품,
            "POS정정요청공제": POS공제,
            "POS정정요청공제환급금": POS환급,
            "유통하자/재고LOSS": LOSS,
            "①공급가액":  공급가액,
            "②부가세":   부가세,
            "③지급할총액": 지급할총액,
            "전화요금공제": 전화요금공제,
            "세액공제(별도)": 세액공제,
            "직배비공제(V+)": 직배비공제vp,
            "ⓕ기타공제합계": 기타공제합계,
            "ⓗ송금액":   송금액,
        })

    # ── 엑셀 생성 ────────────────────────────────────────
    NAVY="1F3864"; BLUE="2E5EAA"; PALE="EEF3FB"; WHITE="FFFFFF"
    GREEN="1E6B3C"; GRAY="F2F2F2"; AMBER="FFF2CC"; RED="C00000"
    FMT="#,##0"

    # (key, header, align, number_format, width) — 판매수수료작업시트_AMD_템플릿.xlsx 컬럼 순서 반영
    # 열너비는 억 단위 합계(예: 3,624,908,030)도 잘리지 않도록 넉넉히 지정
    ITEMS=[
        ("매장코드","매장\n코드","center",None,9),
        ("매장명","매장명","left",None,20),
        ("매니저","매니저","center",None,10),
        ("매출합계","매출합계","center",FMT,17),
        ("본사-M","본사-M","center",FMT,14),
        ("본사-S1","본사-S1","center",FMT,14),
        ("중간관리수수료","중간관리\n수수료","center",FMT,14),
        ("판매수수료합계","판매수수료\n합계","center",FMT,16),
        ("직배비+경비","직배비+경비","center",FMT,14),
        ("아르바이트비","아르바이트비","center",FMT,14),
        ("덜받음","덜받음","center",FMT,12),
        ("임의사은품지급","임의사은품\n지급","center",FMT,13),
        ("POS정정요청공제","POS정정\n요청공제","center",FMT,13),
        ("POS정정요청공제환급금","POS정정\n환급금","center",FMT,13),
        ("유통하자/재고LOSS","유통하자\n/LOSS","center",FMT,13),
        ("①공급가액","①공급가액","center",FMT,16),
        ("②부가세","②부가세","center",FMT,14),
        ("③지급할총액","③지급할\n총액","center",FMT,16),
        ("전화요금공제","전화요금\n공제","center",FMT,13),
        ("세액공제(별도)","세액공제\n(별도)","center",FMT,13),
        ("직배비공제(V+)","직배비공제\n(V+)","center",FMT,13),
        ("ⓕ기타공제합계","ⓕ기타공제합계","center",FMT,14),
        ("ⓗ송금액","ⓗ송금액","center",FMT,16),
    ]
    GROUPS=[
        (1,3,"기본정보",NAVY),(4,4,"매출",BLUE),(5,8,"수수료",BLUE),
        (9,10,"지원금","375623"),(11,15,"공제항목",RED),
        (16,18,"지급금액","4472C4"),(19,22,"기타공제","AA5B1E"),
        (23,23,"송금액","276221"),
    ]
    # 열 문자 기준 Excel 수식 (판매수수료작업시트_AMD_템플릿.xlsx 그대로)
    FORMULA_COLS = {
        8:  "=SUM(E{r}:G{r})",
        16: "=SUM(H{r}:J{r})-SUM(K{r}:M{r},O{r})+N{r}",
        18: "=SUM(P{r}:Q{r})",
        22: "=SUM(S{r}:U{r})",
        23: "=R{r}-V{r}",
    }

    wb = openpyxl.Workbook(); ws = wb.active
    ws.title="판매수수료작업시트(AMD)"
    ws.sheet_view.showGridLines=False; ws.freeze_panes="D4"

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

    for ei,p in enumerate(store_rows):
        row=ei+4
        송금액=p.get("ⓗ송금액") or 0
        rb=WHITE if ei%2==0 else GRAY

        for ci,(key,hdr,align,fmt,w) in enumerate(ITEMS,1):
            bg=rb; fg="000000"; bold=False

            if key=="③지급할총액": bg="E2EFDA"; bold=True
            elif key=="ⓗ송금액":
                if isinstance(송금액,(int,float)) and 송금액<0:
                    bg="FFE0E0"; fg=RED; bold=True
                else:
                    bg="C6EFCE"; fg="276221"; bold=True

            if ci in FORMULA_COLS:
                v = FORMULA_COLS[ci].format(r=row)
            else:
                v=p.get(key)
                if v is None: v=0 if fmt else ""
                if isinstance(v,(int,float)) and v<0 and key not in ("ⓗ송금액","③지급할총액"):
                    fg=RED

            c=ws.cell(row,ci,v)
            c.font=Font(name="맑은 고딕",bold=bold,color=fg,size=9)
            c.fill=PatternFill("solid",fgColor=bg)
            c.alignment=Alignment(horizontal=align,vertical="center")
            if fmt: c.number_format=fmt
        ws.row_dimensions[row].height=17

    first_row = 4
    sr = len(store_rows) + 4
    ws.merge_cells(f"A{sr}:C{sr}")
    st(sr,1,f"합  계  ({len(store_rows)}개 매장)",bg=NAVY,fg=WHITE,bold=True,size=10)
    for ci,(key,hdr,align,fmt,w) in enumerate(ITEMS,1):
        if not fmt: continue
        col = get_column_letter(ci)
        bg="C6EFCE" if key=="ⓗ송금액" else ("E2EFDA" if key=="③지급할총액" else NAVY)
        fg="276221" if key=="ⓗ송금액" else WHITE
        # 합계행: 수식 열은 합계행 자기 자신을 참조하는 동일 수식(연쇄 계산),
        # 나머지는 데이터 행 전체를 SUM
        if ci in FORMULA_COLS:
            v = FORMULA_COLS[ci].format(r=sr)
        else:
            v = f"=SUM({col}{first_row}:{col}{sr-1})"
        c=ws.cell(sr,ci,v)
        c.font=Font(name="맑은 고딕",bold=True,color=fg,size=10)
        c.fill=PatternFill("solid",fgColor=bg); c.number_format=FMT
        c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[sr].height=22
    bdr(2,1,sr,len(ITEMS))

    out_path = OUTPUT / "판매수수료작업시트_AMD.xlsx"
    wb.save(out_path)
    return out_path, len(store_rows)


def run(ym: str, base_dir: Path, results: list,
        sales_detail: dict, expense_rows: dict, master: list):
    print(f"  📋 판매수수료작업시트(AMD) 처리 중...")
    try:
        path, cnt = make_amd_report(ym, base_dir, results,
                               sales_detail, expense_rows, master)
        print(f"     ✅ 판매수수료작업시트_AMD.xlsx  ({cnt}개 매장)")
        return path
    except Exception as e:
        print(f"     ⚠️  AMD 처리 오류: {e}")
        import traceback; traceback.print_exc()
        return None
