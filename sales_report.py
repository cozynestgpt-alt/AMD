import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

from common import STORE_CODE_MAP, CODE_TO_STORE, ONLINE_CODE_MAP

# ── 색상 ──────────────────────────────────────────────────────
NAVY    = "1F3864"; BLUE    = "2E5EAA"; PALE    = "EEF3FB"
WHITE   = "FFFFFF"; GRAY    = "F2F2F2"; GRAY2   = "E8E8E8"
GREEN   = "1E6B3C"; LGREEN  = "E2EFDA"; AMBER   = "FFF2CC"
RED     = "C00000"; REDL    = "FFE0E0"; DARK_RED= "8B0000"
FMT     = "#,##0"

def st(ws,r,c,v=None,bg=None,fg="000000",bold=False,
       size=9,align="center",wrap=False,fmt=None):
    cell = ws.cell(r,c)
    if v is not None: cell.value = v
    cell.font = Font(name="맑은 고딕", bold=bold, color=fg, size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if bg: cell.fill = PatternFill("solid", fgColor=bg)
    if fmt: cell.number_format = fmt
    return cell

def bdr(ws,r1,c1,r2,c2,outer="medium",inner="thin"):
    tk=Side(style=outer); tn=Side(style=inner)
    for r in range(r1,r2+1):
        for c in range(c1,c2+1):
            ws.cell(r,c).border = Border(
                left=tk if c==c1 else tn, right=tk if c==c2 else tn,
                top=tk if r==r1 else tn, bottom=tk if r==r2 else tn)

# ── 1. 영업매출 파일 읽기 ─────────────────────────────────────
def load_sales_file(path: Path) -> dict:
    """
    영5-9_판매_YYYYMM.xlsx 파싱
    컬럼: A=매장코드, B=매장상호, C~F=정상(수량,판매금액,판매원가,생산원가),
          G~J=행사, K~N=균일, O~R=용품, S~V=기획, W=소계수량, X=소계판매금액, ...
    → {store_name: {off_normal, off_event, off_total, online_normal, online_event, online_total, total}}
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    # 오프라인 / 온라인 분리하여 집계
    off_data  = {}  # store_name → {normal, event, total}
    on_data   = {}  # store_name → {normal, event, total}

    for row in ws.iter_rows(min_row=4, values_only=True):
        code = str(row[0]).strip() if row[0] else ""
        if not code or code == "총계":
            continue
        # 소계 판매금액 = X열(index 23)
        total_sales = int(row[23]) if isinstance(row[23], (int,float)) else 0
        # 정상 = 정상판매(D열,idx3) + 행사판매(H열,idx7)
        # 행사·기획 = 균일(L열,idx11) + 용품(P열,idx15) + 기획(T열,idx19)
        def _v(i): return int(row[i]) if isinstance(row[i],(int,float)) else 0
        normal_sales = _v(3) + _v(7)
        event_sales  = _v(11) + _v(15) + _v(19)

        if code.startswith("1") and len(code)==5:  # 온라인 코드
            store = ONLINE_CODE_MAP.get(code)
            if store:
                if store not in on_data:
                    on_data[store] = {"normal":0,"event":0,"total":0}
                on_data[store]["normal"] += normal_sales
                on_data[store]["event"]  += event_sales
                on_data[store]["total"]  += total_sales
        else:  # 오프라인
            store = CODE_TO_STORE.get(code)
            if store:
                if store not in off_data:
                    off_data[store] = {"normal":0,"event":0,"total":0}
                off_data[store]["normal"] += normal_sales
                off_data[store]["event"]  += event_sales
                off_data[store]["total"]  += total_sales

    # 병합
    all_stores = set(list(off_data.keys()) + list(on_data.keys()))
    result = {}
    for s in all_stores:
        off = off_data.get(s, {"normal":0,"event":0,"total":0})
        on  = on_data.get(s,  {"normal":0,"event":0,"total":0})
        result[s] = {
            "off_normal": off["normal"],
            "off_event":  off["event"],
            "off_total":  off["total"],
            "on_normal":  on["normal"],
            "on_event":   on["event"],
            "on_total":   on["total"],
            "grand_total": off["total"] + on["total"],
        }
    return result

# ── 2. 매출집계 보고서 생성 ───────────────────────────────────
def make_sales_report(sales_data: dict, ym: str, out_dir: Path,
                      manager_map: dict = None,
                      income_map: dict = None) -> Path:
    """
    매장코드 오름차순으로 정렬하여 매출집계 보고서 생성
    manager_map: {store_name: manager_name}
    income_map:  {store_name: "중간관리" | "근로소득" | "사업소득"}
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "매출집계"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    # 열 너비  A~K (H=오프합계, I=온합계신규, J=매출합계, K=비고)
    col_w = [6, 22, 8, 14, 14, 14, 14, 14, 14, 16, 12]
    for i, w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 제목
    ws.merge_cells("A1:K1")
    st(ws,1,1,f"매출 집계 보고서  ▶  {ym}",
       bg=NAVY,fg=WHITE,bold=True,size=14)
    ws.row_dimensions[1].height = 32

    # 안내
    ws.merge_cells("A2:K2")
    st(ws,2,1,f"📌 입력파일: 영{ym[2:4]}-{ym[5:]}_판매_{ym.replace('-','')}.xlsx  |  매장코드 오름차순 정렬",
       bg=PALE,fg="444444",size=9,align="left")
    ws.row_dimensions[2].height = 16

    # 헤더 (2단)
    ws.merge_cells("A3:A4"); st(ws,3,1,"매장\n코드",bg=NAVY,fg=WHITE,bold=True,size=9,wrap=True)
    ws.merge_cells("B3:B4"); st(ws,3,2,"매장명",bg=NAVY,fg=WHITE,bold=True,size=9)
    ws.merge_cells("C3:C4"); st(ws,3,3,"매니저",bg=NAVY,fg=WHITE,bold=True,size=9)
    ws.merge_cells("D3:E3"); st(ws,3,4,"오프라인 매출",bg=BLUE,fg=WHITE,bold=True,size=9)
    ws.merge_cells("F3:G3"); st(ws,3,6,"온라인 매출",bg="1F5C99",fg=WHITE,bold=True,size=9)
    ws.merge_cells("H3:H4"); st(ws,3,8,"오프\n합계",bg=BLUE,fg=WHITE,bold=True,size=9,wrap=True)
    ws.merge_cells("I3:I4"); st(ws,3,9,"온\n합계",bg="1F5C99",fg=WHITE,bold=True,size=9,wrap=True)
    ws.merge_cells("J3:J4"); st(ws,3,10,"매출\n합계",bg=GREEN,fg=WHITE,bold=True,size=10,wrap=True)
    ws.merge_cells("K3:K4"); st(ws,3,11,"비고",bg=NAVY,fg=WHITE,bold=True,size=9)
    st(ws,4,4,"정상",bg=BLUE,fg=WHITE,bold=True,size=9)
    st(ws,4,5,"행사·기획",bg=BLUE,fg=WHITE,bold=True,size=9)
    st(ws,4,6,"정상",bg="1F5C99",fg=WHITE,bold=True,size=9)
    st(ws,4,7,"행사",bg="1F5C99",fg=WHITE,bold=True,size=9)
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 18

    # 매장코드 오름차순 정렬
    sorted_stores = sorted(
        [(code, name) for name, code in STORE_CODE_MAP.items()
         if name in sales_data or True],
        key=lambda x: x[0]
    )

    grand = {"off_n":0,"off_e":0,"off_t":0,"on_n":0,"on_e":0,"on_t":0,"total":0}
    ei = 0
    for code, store in sorted_stores:
        d = sales_data.get(store)
        if not d:
            continue
        row = ei + 5
        rb  = WHITE if ei % 2 == 0 else LGREEN
        mgr = (manager_map or {}).get(store, "")

        st(ws,row,1,code,bg=rb,size=9,fg="555555")
        st(ws,row,2,store,bg=rb,size=9,align="left")
        st(ws,row,3,mgr,bg=rb,size=9)
        st(ws,row,4,d["off_normal"],bg=rb,size=9,fmt=FMT)
        # 행사·기획 = 소계 - 정상 = 행사+균일+용품+기획 합산
        off_event_etc = d["off_total"] - d["off_normal"]
        st(ws,row,5,off_event_etc,bg=rb,size=9,fmt=FMT)
        st(ws,row,6,d["on_normal"],bg=rb,size=9,fmt=FMT)
        on_event_etc = d["on_total"] - d["on_normal"]
        st(ws,row,7,on_event_etc,bg=rb,size=9,fmt=FMT)
        # H: 오프합계
        c8 = ws.cell(row,8,d["off_total"])
        c8.font=Font(name="맑은 고딕",bold=True,color="000000",size=9)
        c8.fill=PatternFill("solid",fgColor=PALE); c8.number_format=FMT
        c8.alignment=Alignment(horizontal="center",vertical="center")
        # I: 온합계 (온라인 정상+행사) ← 신규
        c9 = ws.cell(row,9,d["on_total"])
        c9.font=Font(name="맑은 고딕",bold=True,color=WHITE,size=9)
        c9.fill=PatternFill("solid",fgColor="1F5C99"); c9.number_format=FMT
        c9.alignment=Alignment(horizontal="center",vertical="center")
        # J: 매출합계
        c10 = ws.cell(row,10,d["grand_total"])
        c10.font=Font(name="맑은 고딕",bold=True,color=WHITE,size=9)
        c10.fill=PatternFill("solid",fgColor=GREEN); c10.number_format=FMT
        c10.alignment=Alignment(horizontal="center",vertical="center")
        # K: 비고
        income = (income_map or {}).get(store, "")
        if income == "중간관리":
            note_val = "중간관리"; note_bg="FFF2CC"; note_fg="7F4F00"
        elif income in ("근로소득","사업소득"):
            note_val = income; note_bg="EEF3FB"; note_fg="1F3864"
        else:
            note_val = ""; note_bg=rb; note_fg="000000"
        c11 = ws.cell(row,11,note_val)
        c11.font=Font(name="맑은 고딕",bold=(income=="중간관리"),color=note_fg,size=9)
        c11.fill=PatternFill("solid",fgColor=note_bg)
        c11.alignment=Alignment(horizontal="center",vertical="center")

        ws.row_dimensions[row].height = 17
        grand["off_n"] += d["off_normal"]
        grand["off_e"] += off_event_etc
        grand["off_t"] += d["off_total"]
        grand["on_n"]  += d["on_normal"]
        grand["on_e"]  += on_event_etc
        grand["on_t"]  += d["on_total"]
        grand["total"] += d["grand_total"]
        ei += 1

    # 합계 행
    sr = ei + 5
    ws.merge_cells(f"A{sr}:C{sr}")
    st(ws,sr,1,f"합  계  ({ei}개 매장)",bg=NAVY,fg=WHITE,bold=True,size=10)
    for ci,val,bg,fg in [
        (4,grand["off_n"],NAVY,WHITE),(5,grand["off_e"],NAVY,WHITE),
        (6,grand["on_n"],NAVY,WHITE),(7,grand["on_e"],NAVY,WHITE),
        (8,grand["off_t"],PALE,"000000"),
        (9,grand["on_t"],"1F5C99",WHITE),
        (10,grand["total"],GREEN,WHITE),
    ]:
        c=ws.cell(sr,ci,val)
        c.font=Font(name="맑은 고딕",bold=True,color=fg,size=10)
        c.fill=PatternFill("solid",fgColor=bg); c.number_format=FMT
        c.alignment=Alignment(horizontal="center",vertical="center")
    ws.cell(sr,11).fill=PatternFill("solid",fgColor=NAVY)
    ws.row_dimensions[sr].height = 22

    bdr(ws,3,1,sr,11)
    path = out_dir / "매출집계.xlsx"
    wb.save(path)
    print(f"  ✅ 매출집계.xlsx  ({ei}개 매장, 합계: {grand['total']:,}원)")
    return path

# ── 메인 ─────────────────────────────────────────────────────
if __name__ == "__main__":
    ym = "2026-05"
    src = Path("/mnt/user-data/uploads/영5-9_판매_202605.xlsx")
    out = Path("/mnt/user-data/outputs")
    out.mkdir(exist_ok=True)

    print("매출 파일 읽는 중...")
    sales = load_sales_file(src)
    print(f"  → {len(sales)}개 매장")

    # 매니저 정보 (사원마스터에서 읽을 수 있지만 여기선 기존 파일 참조)
    mgr_map = {
        "롯데잠실점":"조종숙","롯데관악점":"백경님","롯데광주점":"송미랑",
        "롯데강남점":"마선예","롯데포항점":"박서영","롯데울산점":"양미경",
        "롯데동래점":"최경진","롯데창원점":"함미란","롯데상인점":"이영옥",
        "롯데미아점":"차정희","롯데구리점":"김주희","롯데영등포점":"최종수",
        "롯데부산본점":"양현숙","롯데평촌점":"최인영","롯데안산점":"이경남",
        "롯데노원점":"임희선","롯데광복점":"양선희","롯데수원점":"장양옥",
        "롯데대구점":"황정애","롯데전주점":"송은주","롯데인천터미널점":"정광옥",
        "롯데동탄점":"최혜진","현대천호점":"이은영","현대울산점":"이현래",
        "현대미아점":"전혜정","현대중동점":"이명희","현대킨텍스점":"이희선",
        "현대대구점":"김종숙","현대충청점":"송경아","현대동구점":"백수연",
        "현대판교점":"김은영","현대신촌점":"이은정",
        "신세계광주점":"김문희","신세계마산점":"이윤옥","신세계강남점":"유태하",
        "신세계경기점":"류영순","신세계센텀점":"강은정","신세계타임스퀘어점":"배향선",
        "신세계천안아산점":"김우진","신세계의정부점":"강보인","신세계김해점":"조명희",
        "신세계스타필드하남점":"조정화","신세계대구점":"진미현","신세계대전점":"김수영",
        "갤러리아천안점":"김미옥","갤러리아광교점":"이현주","갤러리아진주점":"구영주",
        "갤러리아타임월드점":"김민주","AK프라자분당점":"최미희자","AK프라자수원점":"김병희",
        "신세계본점":"김은주","롯데아울렛김해점":"김점희","롯데아울렛이천점":"최혜영",
        "롯데아울렛광명점":"홍인숙","롯데아울렛고양점":"윤효숙","롯데아울렛광교점":"최금요",
        "롯데아울렛군산점":"지민숙","롯데아울렛동부산점":"김헌정","LF스퀘어양주점":"신영숙",
        "신세계아울렛파주점":"김나연","현대아울렛대전점":"장은영","현대아울렛남양주점":"김지현",
    }

    # 소득구분 맵 (테스트용 하드코딩 — run.py에서는 사원마스터에서 자동 추출)
    income_map = {
        "롯데광주점":"중간관리","롯데포항점":"중간관리","롯데울산점":"중간관리",
        "롯데동래점":"중간관리","롯데창원점":"중간관리","롯데영등포점":"중간관리",
        "롯데평촌점":"중간관리","롯데노원점":"중간관리","롯데대구점":"중간관리",
        "현대천호점":"중간관리","현대대구점":"중간관리","현대판교점":"중간관리",
        "신세계광주점":"중간관리","신세계마산점":"중간관리","신세계경기점":"중간관리",
        "신세계센텀점":"중간관리","신세계타임스퀘어점":"중간관리","신세계대구점":"중간관리",
        "갤러리아광교점":"중간관리","신세계본점":"중간관리",
        "롯데아울렛김해점":"중간관리","롯데아울렛이천점":"중간관리",
        "현대아울렛대전점":"중간관리","현대아울렛남양주점":"중간관리",
    }
    make_sales_report(sales, ym, out, mgr_map, income_map)
