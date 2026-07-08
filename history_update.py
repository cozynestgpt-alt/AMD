# -*- coding: utf-8 -*-
"""
월별 손익 DB 업데이트 모듈 v4
- 2026년 5월까지는 기존 DB 수치를 기준으로 보존
- 2026년 6월부터 매출집계_분석.xlsx 자료로 업데이트
- 기준일 = 해당월 말일
- 분기 / 반기 / 시즌(SS, FW) 자동 구분
- 같은 년-월 + 매장코드는 기존 행을 삭제 후 새 자료로 교체
- source: output/YYYY-MM/매출집계_분석.xlsx
"""
from pathlib import Path
from datetime import date, datetime
import calendar, re, shutil
from collections import defaultdict
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB_HEADERS = ['년','월','년-월','기준일','분기','반기','시즌연도','시즌','시즌명','매장코드','매장명','판매금액','수금액(V+)','수금액(V-)','생산원가(V-)','영업이익(V-)','총경비','순이익','순이익율']
NAVY='1F3864'; WHITE='FFFFFF'; FMT='#,##0'; PCT='0.00%'

def month_end(year:int, month:int):
    return date(year, month, calendar.monthrange(year, month)[1])

def quarter(month:int):
    return f"{((month-1)//3)+1}Q"

def half(month:int):
    return '상반기' if month <= 6 else '하반기'

def season_info(year:int, month:int):
    if 3 <= month <= 8:
        season_year, season = year, 'SS'
    elif month >= 9:
        season_year, season = year, 'FW'
    else:
        season_year, season = year - 1, 'FW'
    return season_year, season, f'{season_year}{season}'

def _n(v):
    if v is None or v == '': return 0
    if isinstance(v, (int,float)): return v
    s = str(v).replace(',', '').replace('%', '').strip()
    try: return float(s)
    except: return 0

def _clean_code(v):
    s = str(v or '').strip()
    if not s: return ''
    return s.zfill(5) if s.isdigit() else s

def find_header_row(ws):
    for r in range(1, min(ws.max_row, 10)+1):
        values = [str(ws.cell(r,c).value or '').replace('\n','').strip() for c in range(1, ws.max_column+1)]
        if '매장상호' in values and '판매금액' in values:
            return r
    return 4

def read_analysis(path: Path, ym: str):
    year, month = map(int, ym.split('-'))
    sy, ss, sn = season_info(year, month)
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb['매출집계(분석)'] if '매출집계(분석)' in wb.sheetnames else wb.active
    hr = find_header_row(ws)
    rows = []
    for r in range(hr+1, ws.max_row+1):
        code = _clean_code(ws.cell(r,1).value)
        store = str(ws.cell(r,2).value or '').strip()
        if not code or not store or '합계' in store or '총계' in store: continue
        sales = _n(ws.cell(r,3).value)
        vp = _n(ws.cell(r,4).value)
        vn = _n(ws.cell(r,7).value)
        cg = _n(ws.cell(r,8).value)
        op = _n(ws.cell(r,9).value)
        total_exp = _n(ws.cell(r,21).value)
        profit = _n(ws.cell(r,22).value)
        # 순이익율은 V- 기준: 순이익(V-) / (판매금액(V+) / 1.1)
        denom_sales_vminus = sales / 1.1 if sales else 0
        rate = profit / denom_sales_vminus if denom_sales_vminus else 0
        rows.append([year, f'{month:02d}', ym, month_end(year,month), quarter(month), half(month), sy, ss, sn, code, store, sales, vp, vn, cg, op, total_exp, profit, rate])
    return rows

def create_clean_db(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active; ws.title = '년도별월별DB'
    ws.append(DB_HEADERS)
    style_db(ws)
    cfg = wb.create_sheet('Config')
    cfg.append(['항목','값','설명'])
    cfg.append(['기준일','해당월 말일','예: 2026-06 → 2026-06-30'])
    cfg.append(['분기','1Q~4Q','월 기준 자동 구분'])
    cfg.append(['시즌','SS/FW','SS=3~8월, FW=9~다음해 2월'])
    wb.save(path)

def style_db(ws):
    for cell in ws[1]:
        cell.font = Font(name='맑은 고딕', bold=True, color=WHITE, size=10)
        cell.fill = PatternFill('solid', fgColor=NAVY)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for col,w in {'A':7,'B':6,'C':10,'D':12,'E':8,'F':10,'G':10,'H':8,'I':10,'J':10,'K':22,'L':14,'M':14,'N':14,'O':14,'P':14,'Q':14,'R':14,'S':10}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    thin = Side(style='thin', color='D9D9D9')
    for row in ws.iter_rows():
        for c in row:
            c.border = Border(bottom=thin)
            if c.row > 1:
                c.font = Font(name='맑은 고딕', size=9)
    for r in range(2, ws.max_row+1):
        ws.cell(r,4).number_format='yyyy-mm-dd'
        for c in range(12,19): ws.cell(r,c).number_format=FMT
        ws.cell(r,19).number_format=PCT

def update_history(ym: str, base_dir: Path = None, analysis_path: Path = None, db_path: Path = None, force: bool = False):
    base_dir = Path(base_dir or Path(__file__).parent)
    # 운영 기준: 2026년 5월까지는 DB_202605까지.xlsx에서 이관한 기존 DB 수치를 보존합니다.
    # 2026년 6월부터 salary_system의 output/년월/매출집계_분석.xlsx 자료로 월별 DB를 업데이트합니다.
    if ym <= '2026-05' and not force:
        print(f'ℹ️ {ym}은 기존 DB 보존 대상입니다. DB 업데이트를 건너뜁니다. (2026-06부터 자동 업데이트)')
        return {'db_path': str(base_dir/'DB'/'월별손익DB.xlsx'), 'backup_path': '', 'updated_month': ym, 'rows': 0, 'skipped': True}
    analysis_path = Path(analysis_path or base_dir/'output'/ym/'매출집계_분석.xlsx')
    db_path = Path(db_path or base_dir/'DB'/'월별손익DB.xlsx')
    if not analysis_path.exists():
        raise FileNotFoundError(f'매출집계_분석.xlsx 없음: {analysis_path}')
    if not db_path.exists():
        create_clean_db(db_path)
    backup_dir = base_dir/'backup'; backup_dir.mkdir(exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = backup_dir / f'월별손익DB_before_{ym}_{ts}.xlsx'
    shutil.copy2(db_path, backup_path)
    new_rows = read_analysis(analysis_path, ym)
    wb = load_workbook(db_path)
    ws = wb['년도별월별DB'] if '년도별월별DB' in wb.sheetnames else wb.active
    # 기존 같은 년월 행 삭제
    keep = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[2]) != ym:
            keep.append(list(row))
    # 시트 재작성: 외부링크/수식 없이 값만 유지
    wb.remove(ws)
    ws = wb.create_sheet('년도별월별DB', 0)
    ws.append(DB_HEADERS)
    for row in keep + new_rows:
        ws.append(row)
    # 정렬: 년월, 매장코드
    data = list(ws.iter_rows(min_row=2, values_only=True))
    data.sort(key=lambda x: (str(x[2]), str(x[9])))
    ws.delete_rows(2, ws.max_row)
    for row in data: ws.append(list(row))
    style_db(ws)
    wb.save(db_path)
    return {'db_path': str(db_path), 'backup_path': str(backup_path), 'updated_month': ym, 'rows': len(new_rows)}

def run(ym: str, base_dir: Path = None):
    info = update_history(ym, base_dir)
    try:
        from history_report import run as report_run
        report_run(base_dir or Path(__file__).parent)
        try:
            from management_report import run as mgmt_report_run
            mgmt_report_run(base_dir or Path(__file__).parent)
        except Exception as e:
            print(f'⚠️ 경영분석보고서 생성 오류: {e}')
        try:
            from year_compare_report import run as year_compare_run
            year_compare_run(ym, base_dir or Path(__file__).parent)
        except Exception as e:
            print(f'⚠️ 전년대비보고서 생성 오류: {e}')
    except Exception as e:
        print(f'⚠️ 연도별/분기/시즌 보고서 생성 오류: {e}')
    if info.get('skipped'):
        print(f"✅ 기존 DB 보존 대상월 처리 완료: {info['updated_month']} / DB 변경 없음")
    else:
        print(f"✅ 월별손익DB 업데이트 완료: {info['updated_month']} / {info['rows']}행")
        print(f"   백업: {info['backup_path']}")
    print(f"   DB: {info['db_path']}")
    return info

if __name__ == '__main__':
    ym = input('업데이트할 년월을 입력하세요 (예: 2026-06 또는 202606): ').strip()
    if len(ym)==6 and ym.isdigit(): ym = ym[:4]+'-'+ym[4:]
    run(ym, Path(__file__).parent)
    input('\nEnter 키를 눌러 종료합니다...')
