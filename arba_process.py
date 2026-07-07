"""
아르바이트 급여 자동 취합 모듈
입력:
  input/사원현황_YYYY-MM.xlsx   → 성명·사번·부서·주민번호
  input/급여상여명세서_YYYY-MM.xlsx → 총지급액·공제내역
  input/월별근태_YYYY-MM.xlsx   → 실근(E열) = 근무일수
출력:
  output/YYYY-MM/아르바이트_정산서.xlsx
"""

import sys
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "--quiet"])
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

# ── 부서명 → 사원마스터 매장명 매핑 테이블 ────────────────────
# 왼쪽: 사원현황·급여명세서·근태 파일의 부서명
# 오른쪽: 사원마스터의 매장명
DEPT_TO_STORE = {
    # 완전 일치
    "롯데잠실점":          "롯데잠실점",
    "롯데관악점":          "롯데관악점",
    "롯데강남점":          "롯데강남점",
    "롯데상인점":          "롯데상인점",
    "롯데미아점":          "롯데미아점",
    "롯데구리점":          "롯데구리점",
    "롯데영등포점":        "롯데영등포점",
    "롯데부산본점":        "롯데부산본점",
    "롯데안산점":          "롯데안산점",
    "롯데노원점":          "롯데노원점",
    "롯데광복점":          "롯데광복점",
    "롯데수원점":          "롯데수원점",
    "롯데대구점":          "롯데대구점",
    "롯데전주점":          "롯데전주점",
    "롯데인천터미널점":    "롯데인천터미널점",
    "롯데동탄점":          "롯데동탄점",
    "롯데잠실캐슬프라자점":"롯데잠실캐슬프라자점",
    "현대천호점":          "현대천호점",
    "현대울산점":          "현대울산점",
    "현대압구정본점":      "현대압구정본점",
    "현대미아점":          "현대미아점",
    "현대중동점":          "현대중동점",
    "현대킨텍스점":        "현대킨텍스점",
    "현대대구점":          "현대대구점",
    "현대충청점":          "현대충청점",
    "현대동구점":          "현대동구점",
    "현대판교점":          "현대판교점",
    "현대신촌점":          "현대신촌점",
    "신세계강남점":        "신세계강남점",
    "신세계경기점":        "신세계경기점",
    "신세계센텀점":        "신세계센텀점",
    "신세계김해점":        "신세계김해점",
    "신세계대구점":        "신세계대구점",
    "신세계대전점":        "신세계대전점",
    "신세계의정부점":      "신세계의정부점",
    "신세계본점":          "신세계본점",
    "갤러리아천안점":      "갤러리아천안점",
    "갤러리아광교점":      "갤러리아광교점",
    "갤러리아진주점":      "갤러리아진주점",
    "갤러리아타임월드점":  "갤러리아타임월드점",
    # 이름이 다른 매핑
    "고양점(롯데아울렛)":  "롯데아울렛고양점",
    "광교점(롯데아울렛)":  "롯데아울렛광교점",
    "광명점(롯데아울렛)":  "롯데아울렛광명점",
    "동부산점(롯데아울렛)":"롯데아울렛동부산점",
    "스타필드하남점":      "신세계스타필드하남점",
    "신세계충청점":        "신세계천안아산점",   # 충청점 = 천안아산점
    "양주점(LF몰)":        "LF스퀘어양주점",
    "파주점_신세계아울렛": "신세계아울렛파주점",
    "AK분당점":            "AK프라자분당점",
    "AK수원점":            "AK프라자수원점",
    "AK광명점":            "AK프라자광명점(직영점)",  # 직영점 — 판매수수료 대상 아님
    # ── 직영점 (판매수수료 대상 아님) ──────────────────────────────
    "일산점":              "일산점(직영점)",
    "양재점":              "양재점(직영점)",
    "경기광주점":          "경기광주점(직영점)",
    "NC충장점":            "NC충장점(직영점)",
    "NC일산점":            "NC일산점(직영점)",
    "NC해운대점":          "NC해운대점(직영점)",
    "청주점":              "청주점(직영점)",
    "현대커넥트부산점":    "현대커넥트부산점(직영점)",
    "전주점":              "전주점(직영점)",
    "가든5점_현대아울렛":  "가든5점(직영점)",
    "NC불광점":            "NC불광점(직영점)",
    # ── 미입점행사 (판매수수료 대상 아님) ───────────────────────
    "ET_월평점":           "ET_월평점(미입점행사)",
}

# ── 색상 ────────────────────────────────────────────────────────
NAVY    = "1F3864"; BLUE    = "2E5EAA"; PALE    = "EEF3FB"
WHITE   = "FFFFFF"; GRAY    = "F2F2F2"; GRAY2   = "E8E8E8"
GREEN   = "1E6B3C"; LGREEN  = "E2EFDA"; AMBER   = "FFF2CC"
RED     = "C00000"; REDL    = "FFE0E0"; ORANGE  = "F4B942"
ORANGEL = "FFF0CC"; DARK_RED= "8B0000"
INPUT_C = "0000FF"; LINK_G  = "008000"
FMT     = "#,##0"

def _st(ws, r, c, v=None, bg=None, fg="000000", bold=False,
        size=9, align="center", wrap=False, fmt=None):
    cell = ws.cell(r, c)
    if v is not None:
        cell.value = v
    cell.font = Font(name="맑은 고딕", bold=bold, color=fg, size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    if fmt:
        cell.number_format = fmt
    return cell

def _bdr(ws, r1, c1, r2, c2):
    tk = Side(style="medium"); tn = Side(style="thin")
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            ws.cell(r,c).border = Border(
                left=tk if c==c1 else tn, right=tk if c==c2 else tn,
                top=tk if r==r1 else tn, bottom=tk if r==r2 else tn)

# ══════════════════════════════════════════════════════════════
# 1. 파일 읽기
# ══════════════════════════════════════════════════════════════
def load_staff(path: Path) -> dict:
    """사원현황 → {사번: {name, dept, store, jumin, phone}}"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    data = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        name  = str(row[0]).strip()
        sabun = str(row[1]).strip() if row[1] else ""
        dept  = str(row[2]).strip() if row[2] else ""
        jumin = str(row[5]).strip() if row[5] else ""
        phone = str(row[27]).strip() if row[27] else ""
        store = DEPT_TO_STORE.get(dept, dept)  # 매핑 없으면 부서명 그대로
        data[sabun] = {
            "name": name, "dept": dept, "store": store,
            "jumin": jumin, "phone": phone,
        }
    return data

def load_salary(path: Path) -> dict:
    """급여명세서 → {사번: {total_pay, deduct, real_pay, emp_ins, health_ins, pension, income_tax}}"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    data = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[1]:
            continue
        sabun     = str(row[2]).strip() if row[2] else ""
        total_pay = row[9]  or 0   # 총지급액 (J열)
        deduct    = row[10] or 0   # 공제총액 (K열)
        real_pay  = row[11] or 0   # 실지급액 (L열)
        income_tax= row[15] or 0   # 소득세   (P열)
        local_tax = row[16] or 0   # 지방소득세(Q열)
        health    = row[17] or 0   # 의료보험 (R열)
        pension   = row[18] or 0   # 국민연금 (S열)
        emp_ins   = row[19] or 0   # 고용보험 (T열)
        data[sabun] = {
            "total_pay": int(total_pay), "deduct": int(deduct),
            "real_pay":  int(real_pay),  "emp_ins": int(emp_ins),
            "health":    int(health),    "pension": int(pension),
            "income_tax":int(income_tax),"local_tax":int(local_tax),
        }
    return data

def load_attendance(path: Path) -> dict:
    """월별근태 → {사번: 실근일수(E열)}"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    data = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        sabun    = str(row[1]).strip() if row[1] else ""
        work_days= row[4] or 0   # E열 = 실근
        data[sabun] = int(work_days)
    return data

# ══════════════════════════════════════════════════════════════
# 2. 데이터 병합
# ══════════════════════════════════════════════════════════════
def merge(staff: dict, salary: dict, attendance: dict) -> list:
    """3개 파일을 사번 기준으로 병합"""
    results = []
    for sabun, s in staff.items():
        sal  = salary.get(sabun, {})
        days = attendance.get(sabun, 0)
        results.append({
            "sabun":      sabun,
            "name":       s["name"],
            "dept":       s["dept"],
            "store":      s["store"],
            "jumin":      s["jumin"],
            "phone":      s["phone"],
            "work_days":  days,
            "total_pay":  sal.get("total_pay", 0),
            "emp_ins":    sal.get("emp_ins",   0),
            "health":     sal.get("health",    0),
            "pension":    sal.get("pension",   0),
            "income_tax": sal.get("income_tax",0),
            "deduct":     sal.get("deduct",    0),
            "real_pay":   sal.get("real_pay",  0),
            "no_salary":  sabun not in salary,
            "no_attend":  sabun not in attendance,
        })
    # 매장코드 오름차순 정렬 (sales_report의 STORE_CODE_MAP 참조)
    try:
        from sales_report import STORE_CODE_MAP as _SCM
        _code_order = {name: code for name, code in _SCM.items()}
    except ImportError:
        _code_order = {}
    results.sort(key=lambda x: (
        _code_order.get(x["store"] or "", "99999"),  # 코드순, 미매핑은 맨 뒤
        x["name"]
    ))
    return results

# ══════════════════════════════════════════════════════════════
# 3. 출력 파일 생성
# ══════════════════════════════════════════════════════════════
def make_arba_xlsx(results: list, ym: str, out_dir: Path) -> Path:
    wb = openpyxl.Workbook()

    # ── 시트1: 전체목록 ──────────────────────────────────────
    ws = wb.active
    ws.title = "전체목록"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    col_w = [5, 6, 20, 20, 8, 14, 10, 6, 12, 10, 10, 10, 10, 14, 10]
    hdrs = [
        ("NO",    NAVY), ("사번",   NAVY), ("매장명",  NAVY), ("부서명(원본)", NAVY),
        ("성명",  NAVY), ("주민번호", BLUE), ("연락처", BLUE),
        ("근무일수", "BF8F00"), ("총지급액", NAVY), ("고용보험", NAVY),
        ("의료보험", NAVY), ("국민연금", NAVY), ("소득세", NAVY),
        ("실지급액", GREEN), ("비고", NAVY),
    ]
    for i, w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:O1")
    _st(ws,1,1,f"📋 아르바이트 급여 정산서  ▶  {ym}",
        bg=NAVY,fg=WHITE,bold=True,size=13)
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:O2")
    _st(ws,2,1,
        "📌 사번 기준으로 [사원현황] + [급여명세서] + [월별근태] 3개 파일을 자동 병합한 결과입니다.",
        bg=AMBER,fg="664400",size=9,align="left")
    ws.row_dimensions[2].height = 16
    ws.merge_cells("A3:O3")
    _st(ws,3,1,
        f"⚠️ 빨간배경 = 급여 또는 근태 파일 미매칭(사번 확인 필요)  |  생성: {datetime.now():%Y-%m-%d %H:%M}",
        bg=REDL,fg=RED,size=9,align="left")
    ws.row_dimensions[3].height = 14

    for ci, (h, bg) in enumerate(hdrs, 1):
        _st(ws, 4, ci, h, bg=bg, fg=WHITE, bold=True, size=9, wrap=True)
    ws.row_dimensions[4].height = 28

    no = 1
    store_totals = {}
    for ei, r in enumerate(results):
        row = ei + 5
        err = r["no_salary"] or r["no_attend"]
        rb  = REDL if err else (WHITE if ei%2==0 else GRAY2)
        note = []
        if r["no_salary"]:  note.append("급여파일 미매칭")
        if r["no_attend"]:  note.append("근태파일 미매칭")
        if r["store"] is None: note.append("매장매핑 미확인")

        vals = [
            no, r["sabun"], r["store"] or f"[미매핑]{r['dept']}",
            r["dept"], r["name"], r["jumin"], r["phone"],
            r["work_days"], r["total_pay"], r["emp_ins"],
            r["health"], r["pension"], r["income_tax"],
            r["real_pay"], ", ".join(note),
        ]
        for ci, v in enumerate(vals, 1):
            fg = WHITE if ci==14 else (RED if err else "000000")
            bg = GREEN if ci==14 else rb
            bold = ci==14
            fmt_v = FMT if ci in range(9,14) or ci==14 else None
            aln = "left" if ci in (3,4,5,7,15) else "center"
            c = ws.cell(row, ci, v)
            c.font = Font(name="맑은 고딕", bold=bold, color=fg, size=9)
            c.alignment = Alignment(horizontal=aln, vertical="center")
            if bg:
                c.fill = PatternFill("solid", fgColor=bg)
            if fmt_v:
                c.number_format = fmt_v
            if ci in (6,): # 주민번호 텍스트
                c.number_format = "@"

        ws.row_dimensions[row].height = 17
        no += 1

        # 매장별 집계용
        store = r["store"] or r["dept"]
        if store not in store_totals:
            store_totals[store] = {"cnt":0,"days":0,"total":0,"ins":0,"real":0}
        store_totals[store]["cnt"]   += 1
        store_totals[store]["days"]  += r["work_days"]
        store_totals[store]["total"] += r["total_pay"]
        store_totals[store]["ins"]   += r["emp_ins"]+r["health"]+r["pension"]+r["income_tax"]
        store_totals[store]["real"]  += r["real_pay"]

    # 합계
    sr = len(results)+5
    ws.merge_cells(f"A{sr}:H{sr}")
    _st(ws,sr,1,f"합계  ({len(results)}명)",bg=NAVY,fg=WHITE,bold=True,size=10)
    for ci,col_l in [(9,"I"),(10,"J"),(11,"K"),(12,"L"),(13,"M"),(14,"N")]:
        c = ws.cell(sr,ci,f"=SUM({col_l}5:{col_l}{sr-1})")
        bg = GREEN if ci==14 else NAVY
        c.font=Font(name="맑은 고딕",bold=True,color=WHITE,size=10)
        c.fill=PatternFill("solid",fgColor=bg); c.number_format=FMT
        c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[sr].height=22
    _bdr(ws,4,1,sr,15)

    # ── 시트2: 매장별집계 ──────────────────────────────────
    ws2 = wb.create_sheet("매장별집계")
    ws2.sheet_view.showGridLines = False
    ws2.freeze_panes = "A4"

    ws2.merge_cells("A1:G1")
    _st(ws2,1,1,f"📊 매장별 아르바이트 집계  ▶  {ym}",
        bg=NAVY,fg=WHITE,bold=True,size=13)
    ws2.row_dimensions[1].height=28

    for i,w in enumerate([5,22,8,10,14,14,14],1):
        ws2.column_dimensions[get_column_letter(i)].width=w
    for ci,h in enumerate(
        [("NO",NAVY),("매장명",NAVY),("인원",NAVY),
         ("총근무일수",NAVY),("총지급액",NAVY),("공제합계",NAVY),("실지급액",GREEN)],1):
        _st(ws2,3,ci,h[0],bg=h[1],fg=WHITE,bold=True,size=9)
    ws2.row_dimensions[3].height=22

    try:
        from sales_report import STORE_CODE_MAP as _SCM2
        _co2 = {name: code for name, code in _SCM2.items()}
    except ImportError:
        _co2 = {}
    for ei,(store,t) in enumerate(sorted(store_totals.items(),
            key=lambda x: (_co2.get(x[0], "99999"), x[0]))):
        r=ei+4
        rb=WHITE if ei%2==0 else LGREEN
        _st(ws2,r,1,ei+1,bg=rb,fg="888888",size=9)
        _st(ws2,r,2,store,bg=rb,size=9,align="left")
        _st(ws2,r,3,t["cnt"],bg=rb,size=9)
        _st(ws2,r,4,t["days"],bg=rb,size=9,fmt=FMT)
        _st(ws2,r,5,t["total"],bg=rb,size=9,fmt=FMT)
        _st(ws2,r,6,t["ins"],bg=rb,size=9,fmt=FMT)
        c=ws2.cell(r,7,t["real"])
        c.font=Font(name="맑은 고딕",bold=True,color=WHITE,size=9)
        c.fill=PatternFill("solid",fgColor=GREEN); c.number_format=FMT
        c.alignment=Alignment(horizontal="center",vertical="center")
        ws2.row_dimensions[r].height=17

    sr2=len(store_totals)+4
    ws2.merge_cells(f"A{sr2}:B{sr2}")
    _st(ws2,sr2,1,f"합계  ({len(store_totals)}개 매장)",bg=NAVY,fg=WHITE,bold=True,size=10)
    for ci in range(3,8):
        col_l=get_column_letter(ci)
        c=ws2.cell(sr2,ci,f"=SUM({col_l}4:{col_l}{sr2-1})")
        bg=GREEN if ci==7 else NAVY
        c.font=Font(name="맑은 고딕",bold=True,color=WHITE,size=10)
        c.fill=PatternFill("solid",fgColor=bg); c.number_format=FMT
        c.alignment=Alignment(horizontal="center",vertical="center")
    ws2.row_dimensions[sr2].height=22
    _bdr(ws2,3,1,sr2,7)

    # ── 시트3: 부서매핑확인 ────────────────────────────────
    ws3 = wb.create_sheet("부서매핑확인")
    ws3.sheet_view.showGridLines=False
    ws3.column_dimensions["A"].width=5
    ws3.column_dimensions["B"].width=25
    ws3.column_dimensions["C"].width=25
    ws3.column_dimensions["D"].width=15

    ws3.merge_cells("A1:D1")
    _st(ws3,1,1,"🔗 부서명 → 매장명 매핑 현황",bg=NAVY,fg=WHITE,bold=True,size=12)
    ws3.row_dimensions[1].height=28
    ws3.merge_cells("A2:D2")
    _st(ws3,2,1,
        "⚠️ '매핑없음' 항목은 arba_process.py의 DEPT_TO_STORE 딕셔너리에 추가하세요.",
        bg=REDL,fg=RED,size=9,align="left")
    ws3.row_dimensions[2].height=16

    for ci,h in enumerate([("NO",NAVY),("부서명(원본)",NAVY),("→ 매장명",GREEN),("상태",NAVY)],1):
        _st(ws3,3,ci,h[0],bg=h[1],fg=WHITE,bold=True,size=9)
    ws3.row_dimensions[3].height=20

    seen = {}
    for r in results:
        if r["dept"] not in seen:
            seen[r["dept"]] = r["store"]

    try:
        from sales_report import STORE_CODE_MAP as _SCM3
        _co3 = {name: code for name, code in _SCM3.items()}
    except ImportError:
        _co3 = {}
    for ei,(dept,store) in enumerate(sorted(seen.items(),
            key=lambda x: (_co3.get(x[1] or "", "99999"), x[0]))):
        r=ei+4
        rb=WHITE if ei%2==0 else GRAY2
        _st(ws3,r,1,ei+1,bg=rb,fg="888888",size=9)
        _st(ws3,r,2,dept,bg=rb,size=9,align="left")
        if store:
            _st(ws3,r,3,store,bg=LGREEN,fg=LINK_G,size=9,align="left")
            _st(ws3,r,4,"✅ 매핑완료",bg=LGREEN,fg=GREEN,size=9,bold=True)
        else:
            _st(ws3,r,3,"미매핑",bg=REDL,fg=RED,size=9)
            _st(ws3,r,4,"❌ 매핑없음",bg=REDL,fg=RED,size=9,bold=True)
        ws3.row_dimensions[r].height=17
    _bdr(ws3,3,1,len(seen)+3,4)

    # ── 시트4: 점포별소계 ──────────────────────────────────
    ws4 = wb.create_sheet("점포별소계")
    ws4.sheet_view.showGridLines = False
    ws4.freeze_panes = "A3"

    # ── 구분 분류 함수 ──────────────────────────────────────
    def get_category(store: str, dept: str = "") -> str:
        """매핑된 매장명(우선) 또는 원본 부서명(dept) → 구분(백화점/직영점/미입점행사)"""
        s = (store or "").strip()
        if s.endswith("(직영점)"):
            return "직영점"
        if s.endswith("(미입점행사)"):
            return "미입점행사"
        # DEPT_TO_STORE에 아직 등록 안 된 신규 미입점행사 매장 대비 폴백
        if (dept or "").strip().startswith("ET_"):
            return "미입점행사"
        return "백화점"

    # 구분 순서 및 색상
    CAT_ORDER  = ["백화점", "직영점", "미입점행사"]
    CAT_COLOR  = {
        "백화점":   "BDD7EE",   # 연파랑
        "직영점":   "C6EFCE",   # 연초록
        "미입점행사":"FCE4D6",  # 연주황
    }
    CAT_HDR_COLOR = {
        "백화점":   "2E75B6",
        "직영점":   "375623",
        "미입점행사":"843C0C",
    }
    SUBCAT_BG  = {
        "백화점":   "DEEAF1",
        "직영점":   "E2EFDA",
        "미입점행사":"FCE4D6",
    }
    CAT_TOT_BG = {
        "백화점":   "2E75B6",
        "직영점":   "375623",
        "미입점행사":"843C0C",
    }

    # 열 너비 (A=구분, B=점포명, C=성명, D=근무일수, E=일급, F=금액)
    for i, w in enumerate([10, 20, 8, 10, 12, 14], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w

    # 제목
    ws4.merge_cells("A1:F1")
    _st(ws4,1,1,f"점포별 아르바이트비 소계  ▶  {ym}",
        bg=NAVY, fg=WHITE, bold=True, size=13)
    ws4.row_dimensions[1].height = 30

    # 헤더
    for ci, (h, bg) in enumerate([
        ("구분",  NAVY), ("점포명", NAVY), ("성명", NAVY),
        ("근무일수", NAVY), ("일급", NAVY), ("금액", GREEN)
    ], 1):
        _st(ws4, 2, ci, h, bg=bg, fg=WHITE, bold=True, size=10)
    ws4.row_dimensions[2].height = 22

    # ── 데이터 그룹핑: 구분 → 점포명 → 인원 ──────────────
    from collections import defaultdict, OrderedDict
    cat_store = defaultdict(lambda: defaultdict(list))
    for r in results:
        dept  = r["dept"]
        store = r["store"] or r["dept"]
        cat   = get_category(store, dept)
        cat_store[cat][store].append(r)

    tk = Side(style="medium"); tn = Side(style="thin")
    row = 3
    grand_total = 0

    def daily_rate(m):
        if m["work_days"] and m["work_days"] > 0 and m["total_pay"] > 0:
            unit = m["total_pay"] // m["work_days"]
            std  = [80000, 85000, 90000, 95000, 100000]
            closest = min(std, key=lambda x: abs(x - unit))
            return closest if abs(closest - unit) < 5000 else unit
        return 0

    def cell_bdr(ws, r, c, left=None, right=None, top=None, bot=None):
        """단일 셀 테두리 지정"""
        ws.cell(r, c).border = Border(
            left=left or Side(style=None),
            right=right or Side(style=None),
            top=top or Side(style=None),
            bottom=bot or Side(style=None),
        )

    TK = Side(style="medium"); TN = Side(style="thin")

    for cat in CAT_ORDER:
        if cat not in cat_store:
            continue

        cat_bg      = CAT_COLOR[cat]
        cat_hdr_bg  = CAT_HDR_COLOR[cat]
        sub_bg      = SUBCAT_BG[cat]
        cat_tot_bg  = CAT_TOT_BG[cat]

        try:
            from sales_report import STORE_CODE_MAP as _SCM4
            _co4 = {name: code for name, code in _SCM4.items()}
        except ImportError:
            _co4 = {}
        stores = dict(sorted(cat_store[cat].items(),
                     key=lambda x: (_co4.get(x[0], "99999"), x[0])))
        cat_start   = row
        cat_total   = 0

        for si, (store, members) in enumerate(stores.items()):
            store_start = row
            store_total = 0

            for mi, m in enumerate(members):
                rb = WHITE if mi % 2 == 0 else cat_bg

                # A: 구분 — 첫 번째 점포의 첫 번째 행만 표시 (나중에 병합)
                ca = ws4.cell(row, 1, "")
                ca.fill = PatternFill("solid", fgColor=cat_bg)
                ca.alignment = Alignment(horizontal="center", vertical="center",
                                         wrap_text=False)

                # B: 점포명 — 점포 첫 행만 표시
                cb = ws4.cell(row, 2, store if mi == 0 else "")
                cb.font = Font(name="맑은 고딕", bold=(mi==0), color="000000", size=10)
                cb.fill = PatternFill("solid", fgColor=rb)
                cb.alignment = Alignment(horizontal="center", vertical="center")

                # C: 성명
                cc = ws4.cell(row, 3, m["name"])
                cc.font = Font(name="맑은 고딕", color="000000", size=10)
                cc.fill = PatternFill("solid", fgColor=rb)
                cc.alignment = Alignment(horizontal="center", vertical="center")

                # D: 근무일수
                cd = ws4.cell(row, 4, m["work_days"])
                cd.font = Font(name="맑은 고딕", color="000000", size=10)
                cd.fill = PatternFill("solid", fgColor=rb)
                cd.alignment = Alignment(horizontal="center", vertical="center")

                # E: 일급
                dr = daily_rate(m)
                ce = ws4.cell(row, 5, dr if dr else "-")
                ce.font = Font(name="맑은 고딕", color="000000", size=10)
                ce.fill = PatternFill("solid", fgColor=rb)
                ce.number_format = FMT if dr else "@"
                ce.alignment = Alignment(horizontal="center", vertical="center")

                # F: 금액
                amt = m["total_pay"]
                cf = ws4.cell(row, 6, amt)
                cf.font = Font(name="맑은 고딕", color="000000", size=10)
                cf.fill = PatternFill("solid", fgColor=rb)
                cf.number_format = FMT
                cf.alignment = Alignment(horizontal="center", vertical="center")

                store_total += amt
                ws4.row_dimensions[row].height = 18
                row += 1

            # ── 점포 소계 행 ──────────────────────────────
            ws4.merge_cells(f"B{row}:E{row}")
            c = ws4.cell(row, 1, "")
            c.fill = PatternFill("solid", fgColor=cat_bg)

            cs = ws4.cell(row, 2, f"{store}  소계")
            cs.font = Font(name="맑은 고딕", bold=True, color="000000", size=10)
            cs.fill = PatternFill("solid", fgColor=sub_bg)
            cs.alignment = Alignment(horizontal="right", vertical="center")

            cf = ws4.cell(row, 6, store_total)
            cf.font = Font(name="맑은 고딕", bold=True, color="000000", size=10)
            cf.fill = PatternFill("solid", fgColor=sub_bg)
            cf.number_format = FMT
            cf.alignment = Alignment(horizontal="center", vertical="center")
            ws4.row_dimensions[row].height = 20

            # 점포 소계 테두리
            for c in range(1, 7):
                ws4.cell(row, c).border = Border(
                    left=TK if c==1 else TN,
                    right=TK if c==6 else TN,
                    top=TK, bottom=TK,
                )
            row += 1
            cat_total += store_total

        # ── 구분 합계 행 ──────────────────────────────────
        ws4.merge_cells(f"A{row}:E{row}")
        ch = ws4.cell(row, 1, f"{cat}  합계")
        ch.font = Font(name="맑은 고딕", bold=True, color=WHITE, size=11)
        ch.fill = PatternFill("solid", fgColor=cat_tot_bg)
        ch.alignment = Alignment(horizontal="center", vertical="center")

        cht = ws4.cell(row, 6, cat_total)
        cht.font = Font(name="맑은 고딕", bold=True, color=WHITE, size=11)
        cht.fill = PatternFill("solid", fgColor=cat_tot_bg)
        cht.number_format = FMT
        cht.alignment = Alignment(horizontal="center", vertical="center")
        ws4.row_dimensions[row].height = 24

        for c in range(1, 7):
            ws4.cell(row, c).border = Border(
                left=TK, right=TK, top=TK, bottom=TK)
        row += 1
        grand_total += cat_total

        # ── 구분 A열 세로 병합 ────────────────────────────
        if row - 1 > cat_start:
            ws4.merge_cells(f"A{cat_start}:A{row-2}")  # 합계 행 제외한 데이터 행
            mc = ws4.cell(cat_start, 1, cat)
            mc.font = Font(name="맑은 고딕", bold=True, color="000000", size=11)
            mc.fill = PatternFill("solid", fgColor=cat_bg)
            mc.alignment = Alignment(horizontal="center", vertical="center",
                                      wrap_text=False)
            # 병합 블록 외곽 테두리
            for dr in range(cat_start, row-1):
                for c in range(1, 7):
                    cell = ws4.cell(dr, c)
                    b = cell.border
                    cell.border = Border(
                        left=TK if c==1 else (b.left if b.left.style else TN),
                        right=TK if c==6 else (b.right if b.right.style else TN),
                        top=TK if dr==cat_start else (b.top if b.top.style else TN),
                        bottom=TK if dr==row-2 else (b.bottom if b.bottom.style else TN),
                    )

    # ── 총합계 행 ─────────────────────────────────────────
    ws4.merge_cells(f"A{row}:E{row}")
    cg = ws4.cell(row, 1, "총  합  계")
    cg.font = Font(name="맑은 고딕", bold=True, color=WHITE, size=12)
    cg.fill = PatternFill("solid", fgColor=GREEN)
    cg.alignment = Alignment(horizontal="center", vertical="center")

    cgt = ws4.cell(row, 6, grand_total)
    cgt.font = Font(name="맑은 고딕", bold=True, color=WHITE, size=13)
    cgt.fill = PatternFill("solid", fgColor=GREEN)
    cgt.number_format = FMT
    cgt.alignment = Alignment(horizontal="center", vertical="center")

    for c in range(1, 7):
        ws4.cell(row, c).border = Border(
            left=TK, right=TK, top=TK, bottom=TK)
    ws4.row_dimensions[row].height = 28

    # ── 보험료 집계 테이블 ─────────────────────────────────
    # 구분별로 보험료 합산
    # 전체목록 시트에서 매장명(C열)·부서명원본(D열)·고용보험(J열)·의료보험(K열)·국민연금(L열)·실지급액(N열)·총지급액(I열) 읽기
    ws_all = wb["전체목록"]
    ins_totals = {cat: {"emp":0, "health":0, "pension":0, "real":0, "total":0}
                  for cat in CAT_ORDER}

    for dr in range(5, ws_all.max_row + 1):
        no_val = ws_all.cell(dr, 1).value
        if not no_val or not isinstance(no_val, (int, float)):
            continue
        store_v  = ws_all.cell(dr, 3).value or ""
        dept_v   = ws_all.cell(dr, 4).value or ""
        cat_v    = get_category(str(store_v), str(dept_v))
        emp_v    = ws_all.cell(dr, 10).value or 0   # 고용보험
        health_v = ws_all.cell(dr, 11).value or 0   # 의료보험
        pen_v    = ws_all.cell(dr, 12).value or 0   # 국민연금
        real_v   = ws_all.cell(dr, 14).value or 0   # 실지급액
        total_v  = ws_all.cell(dr, 9).value  or 0   # 총지급액
        ins_totals[cat_v]["emp"]     += int(emp_v)
        ins_totals[cat_v]["health"]  += int(health_v)
        ins_totals[cat_v]["pension"] += int(pen_v)
        ins_totals[cat_v]["real"]    += int(real_v)
        ins_totals[cat_v]["total"]   += int(total_v)

    # 빈 줄 하나 띄우기
    row += 2
    ws4.row_dimensions[row - 1].height = 10

    # 집계 테이블 열 너비 재조정 (A=구분, B=고용보험, C=의료보험, D=국민연금, E=실지급액, F=금액)
    for i, w in enumerate([12, 16, 12, 12, 16, 16], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w

    # 테이블 제목
    ws4.merge_cells(f"A{row}:F{row}")
    _st(ws4, row, 1, "▣ 구분별 공제 및 지급 집계",
        bg=NAVY, fg=WHITE, bold=True, size=11, align="left")
    ws4.row_dimensions[row].height = 24
    row += 1

    # 테이블 헤더
    tbl_hdrs = [("구분", NAVY), ("고용보험", BLUE), ("의료보험", BLUE),
                ("국민연금", BLUE), ("실지급액", GREEN), ("금액(총지급)", NAVY)]
    for ci, (h, bg) in enumerate(tbl_hdrs, 1):
        _st(ws4, row, ci, h, bg=bg, fg=WHITE, bold=True, size=10)
    ws4.row_dimensions[row].height = 22
    tbl_hdr_row = row
    row += 1

    # 구분별 행
    cat_row_start = row
    cat_colors_bg = {"백화점": "BDD7EE", "직영점": "C6EFCE", "미입점행사": "FCE4D6"}
    cat_colors_hdr = {"백화점": "2E75B6", "직영점": "375623", "미입점행사": "843C0C"}

    sum_emp = sum_health = sum_pen = sum_real = sum_total = 0

    for cat in CAT_ORDER:
        t   = ins_totals[cat]
        cbg = cat_colors_bg[cat]
        # 구분명
        c1 = ws4.cell(row, 1, cat)
        c1.font = Font(name="맑은 고딕", bold=True, color="000000", size=10)
        c1.fill = PatternFill("solid", fgColor=cbg)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        # 고용보험
        c2 = ws4.cell(row, 2, t["emp"])
        c2.font = Font(name="맑은 고딕", color="000000", size=10)
        c2.fill = PatternFill("solid", fgColor=cbg)
        c2.number_format = FMT
        c2.alignment = Alignment(horizontal="center", vertical="center")
        # 의료보험
        c3 = ws4.cell(row, 3, t["health"])
        c3.font = Font(name="맑은 고딕", color="000000", size=10)
        c3.fill = PatternFill("solid", fgColor=cbg)
        c3.number_format = FMT
        c3.alignment = Alignment(horizontal="center", vertical="center")
        # 국민연금
        c4 = ws4.cell(row, 4, t["pension"])
        c4.font = Font(name="맑은 고딕", color="000000", size=10)
        c4.fill = PatternFill("solid", fgColor=cbg)
        c4.number_format = FMT
        c4.alignment = Alignment(horizontal="center", vertical="center")
        # 실지급액
        c5 = ws4.cell(row, 5, t["real"])
        c5.font = Font(name="맑은 고딕", bold=True, color="000000", size=10)
        c5.fill = PatternFill("solid", fgColor=cbg)
        c5.number_format = FMT
        c5.alignment = Alignment(horizontal="center", vertical="center")
        # 총지급액
        c6 = ws4.cell(row, 6, t["total"])
        c6.font = Font(name="맑은 고딕", bold=True, color=WHITE, size=10)
        c6.fill = PatternFill("solid", fgColor=cat_colors_hdr[cat])
        c6.number_format = FMT
        c6.alignment = Alignment(horizontal="center", vertical="center")

        sum_emp    += t["emp"]
        sum_health += t["health"]
        sum_pen    += t["pension"]
        sum_real   += t["real"]
        sum_total  += t["total"]
        ws4.row_dimensions[row].height = 20
        row += 1

    # 합계 행
    sum_vals = [("합  계", None), (sum_emp, FMT), (sum_health, FMT),
                (sum_pen, FMT), (sum_real, FMT), (sum_total, FMT)]
    for ci, (v, fmt_v) in enumerate(sum_vals, 1):
        bg = GREEN if ci == 5 else NAVY
        c = ws4.cell(row, ci, v)
        c.font = Font(name="맑은 고딕", bold=True, color=WHITE, size=10)
        c.fill = PatternFill("solid", fgColor=bg)
        if fmt_v:
            c.number_format = fmt_v
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[row].height = 22
    sum_row = row
    row += 1

    # 집계 테이블 전체 테두리
    for dr in range(tbl_hdr_row, sum_row + 1):
        for c in range(1, 7):
            ws4.cell(dr, c).border = Border(
                left=TK if c == 1 else TN,
                right=TK if c == 6 else TN,
                top=TK if dr == tbl_hdr_row else TN,
                bottom=TK if dr == sum_row else TN,
            )

    path = out_dir / "아르바이트_정산서.xlsx"
    wb.save(path)
    return path

# ══════════════════════════════════════════════════════════════
# 4. 메인
# ══════════════════════════════════════════════════════════════
def run(ym: str, base_dir: Path):
    INPUT  = base_dir / "input"
    OUTPUT = base_dir / "output" / ym

    # 파일 경로 — 파일명에 '일용직' 포함 여부 무관하게 glob 검색
    ymc = ym.replace("-", "")
    def _find(patterns):
        for pat in patterns:
            found = list(INPUT.glob(pat))
            if found: return found[0]
        return None

    staff_path  = _find([f"사원현황_{ymc}.xlsx",      f"사원현황*{ymc}*.xlsx",      f"사원현황*일용직*{ymc}*.xlsx"])
    # 일용직 파일 우선 탐지 (급여상여명세서_일용직_YYYYMM.xlsx)
    salary_path = _find([f"급여상여명세서*일용직*{ymc}*.xlsx", f"급여상여명세서_{ymc}.xlsx", f"급여상여명세서*{ymc}*.xlsx"])
    attend_path = _find([f"월별근태_{ymc}.xlsx",       f"월별근태*{ymc}*.xlsx",       f"월별근태*일용직*{ymc}*.xlsx"])

    errors = []
    if not staff_path:  errors.append(f"  ❌ 없음: 사원현황_{ymc}.xlsx (또는 사원현황_일용직_{ymc}.xlsx)")
    if not salary_path: errors.append(f"  ❌ 없음: 급여상여명세서_{ymc}.xlsx")
    if not attend_path: errors.append(f"  ❌ 없음: 월별근태_{ymc}.xlsx")
    if errors:
        print("\n".join(errors))
        return None

    print("     사원현황 읽는 중...")
    staff = load_staff(staff_path)
    print(f"     → {len(staff)}명")

    print("     급여명세서 읽는 중...")
    salary = load_salary(salary_path)
    print(f"     → {len(salary)}건")

    print("     월별근태 읽는 중...")
    attend = load_attendance(attend_path)
    print(f"     → {len(attend)}건")

    print("     데이터 병합 중...")
    results = merge(staff, salary, attend)

    no_sal = [r["name"] for r in results if r["no_salary"]]
    no_att = [r["name"] for r in results if r["no_attend"]]
    no_map = [r["dept"] for r in results if r["store"] is None]

    if no_sal:
        print(f"  ⚠️  급여 미매칭 {len(no_sal)}명: {', '.join(no_sal[:5])}{'...' if len(no_sal)>5 else ''}")
    if no_att:
        print(f"  ⚠️  근태 미매칭 {len(no_att)}명: {', '.join(no_att[:5])}{'...' if len(no_att)>5 else ''}")
    if no_map:
        unique_no_map = list(set(no_map))
        print(f"  ⚠️  매장매핑 미확인 부서: {', '.join(unique_no_map)}")

    OUTPUT.mkdir(parents=True, exist_ok=True)
    path = make_arba_xlsx(results, ym, OUTPUT)
    print(f"     ✅ 아르바이트_정산서.xlsx  ({len(results)}명)")
    return path

if __name__ == "__main__":
    ym = input("정산 월 입력 (예: 2026-05): ").strip()
    if len(ym)==6 and ym.isdigit():
        ym = ym[:4]+"-"+ym[4:]
    run(ym, Path(__file__).parent)
