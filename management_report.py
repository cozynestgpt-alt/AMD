# -*- coding: utf-8 -*-
"""
경영분석 보고서 생성 모듈 v5
- 기존 경영진 보고 양식 참조
- 2026년 5월까지는 DB/월별손익DB.xlsx의 기존 DB 수치 사용
- 2026년 6월부터 history_update.py로 업데이트된 DB 수치 사용
- 단위: 천원
"""
from pathlib import Path
from collections import defaultdict
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList

NAVY = '1F3864'; BLUE='DDEBF7'; BLUE2='BFE8F5'; SKY='B7DEE8'; WHITE='FFFFFF'
GRAY='D9E1F2'; LIGHT='F8FBFD'; PEACH='FCE4D6'; RED='C00000'; REDL='FFE0E0'; GREEN='006100'; GREENL='E2F0D9'; YELLOW='FFF2CC'
BORDER = Side(style='thin', color='A6A6A6')
MED = Side(style='medium', color='595959')
UNIT = 1000
METRICS = ['판매금액','수금액(V+)','수금액(V-)','생산원가(V-)','영업이익(V-)','총경비','순이익']
SHORT_METRICS = ['판매금액','수금액(V+)','생산원가(V-)','총경비','영업이익(V-)']

# 기존 보고서에서 행사/폐점으로 관리하던 매장 기준
EVENT_INFO = {
    '00143': '2026-04 행사',
    '00207': '2026-05 행사',
    '05008': '2025-02~03행사',
    '05553': '2026-03행사',
}
CLOSED_INFO = {
    '00140': '2026-03폐점',
    '00209': '2026-03폐점',
    '00216': '2025-06폐점',
}

CLASS_CATEGORIES = ('동일', '신규', '폐점', '행사')


def _normalize_code(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip().replace('.0', '').zfill(5)


def load_store_classification(base_dir: Path, ym: str):
    """
    input/매장구분_행사_신규_폐점_YYYYMM.xlsx 기준으로 매장을 분류한다.
    파일이 없으면 기존 하드코딩(EVENT_INFO/CLOSED_INFO) 기준으로 fallback 한다.
    반환값: {code: {name, category, memo, source}}
    """
    input_dir = base_dir / 'input' / ym
    candidates = sorted(input_dir.glob('매장구분_행사_신규_폐점*.xlsx'))
    if not candidates:
        fallback = {}
        for code, memo in EVENT_INFO.items():
            fallback[code] = {'name': '', 'category': '행사', 'memo': memo, 'source': 'legacy'}
        for code, memo in CLOSED_INFO.items():
            fallback[code] = {'name': '', 'category': '폐점', 'memo': memo, 'source': 'legacy'}
        return fallback, None

    path = candidates[0]
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb['구분'] if '구분' in wb.sheetnames else wb.active
    header_row = None
    headers = []
    for r in range(1, min(ws.max_row, 20) + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        joined = ' '.join(str(v) for v in vals if v is not None)
        if ('구분' in joined) and (('매장' in joined) or ('행 레이블' in joined) or ('코드' in joined)):
            header_row = r
            headers = vals
            break
    if header_row is None:
        header_row = 3
        headers = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]

    def find_col(*keywords):
        for i, h in enumerate(headers):
            text = str(h or '').strip()
            if all(k in text for k in keywords):
                return i
        return None

    code_col = find_col('코드')
    if code_col is None:
        code_col = find_col('행')
    name_col = find_col('매장명')
    cat_col = find_col('구분')
    memo_col = find_col('비고')
    if code_col is None or name_col is None or cat_col is None:
        # 현재 표준 양식: B=매장코드, C=매장명, D=구분
        code_col, name_col, cat_col = 1, 2, 3

    data = {}
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        code = _normalize_code(row[code_col] if code_col < len(row) else None)
        if not code or code == '00000':
            continue
        name = str(row[name_col] or '').strip() if name_col < len(row) else ''
        category = str(row[cat_col] or '').strip() if cat_col < len(row) else ''
        memo = str(row[memo_col] or '').strip() if memo_col is not None and memo_col < len(row) and row[memo_col] else category
        if category not in CLASS_CATEGORIES:
            continue
        data[code] = {'name': name, 'category': category, 'memo': memo or category, 'source': path.name}
    return data, path


def classify_code(code, store_class, prev=None, cur=None):
    code = _normalize_code(code)
    if code in store_class:
        return store_class[code]['category']
    if code in EVENT_INFO:
        return '행사'
    if code in CLOSED_INFO:
        return '폐점'
    if prev is not None and cur is not None:
        if prev.get('판매금액', 0) > 0 and cur.get('판매금액', 0) == 0:
            return '폐점'
        if prev.get('판매금액', 0) == 0 and cur.get('판매금액', 0) > 0:
            return '신규'
    return '동일'


def class_items(store_class, category):
    return [(code, info) for code, info in sorted(store_class.items()) if info.get('category') == category]


def _k(v):
    return int(round((v or 0) / UNIT))


def _pct(n, d):
    return n / d if d else 0


def _sales_vat_excl(sales_vat_incl):
    # 매출금액은 V+ 기준입니다. 생산원가/총경비/영업이익/순이익은 V- 기준이므로
    # 배수와 율 산출 시 매출을 V- 기준(매출/1.1)으로 환산합니다.
    return sales_vat_incl / 1.1 if sales_vat_incl else 0


def _cost_multiple(sales_vat_incl, production_cost_vat_excl):
    return _sales_vat_excl(sales_vat_incl) / production_cost_vat_excl if production_cost_vat_excl else 0


def _vat_excl_rate(amount_vat_excl, sales_vat_incl):
    denom = _sales_vat_excl(sales_vat_incl)
    return amount_vat_excl / denom if denom else 0


def _safe(v):
    return v if v is not None else 0


def ym_sort_key(ym):
    y, m = str(ym).split('-')
    return int(y), int(m)


def load_db(base_dir: Path):
    db_path = base_dir / 'DB' / '월별손익DB.xlsx'
    if not db_path.exists():
        raise FileNotFoundError(f'월별손익DB.xlsx 없음: {db_path}')
    wb = load_workbook(db_path, data_only=True, read_only=True)
    ws = wb['년도별월별DB'] if '년도별월별DB' in wb.sheetnames else wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[idx['년-월']]:
            continue
        rec = {h: row[i] for h, i in idx.items()}
        rec['년'] = int(rec['년'])
        rec['월'] = int(str(rec['월']))
        for m in METRICS:
            rec[m] = _safe(rec.get(m))
        rec['경비율'] = _vat_excl_rate(rec['총경비'], rec['판매금액'])
        rec['영업이익율'] = _vat_excl_rate(rec['영업이익(V-)'], rec['판매금액'])
        rec['순이익율'] = _vat_excl_rate(rec['순이익'], rec['판매금액'])
        rows.append(rec)
    return rows


def latest_year_month(rows):
    yms = sorted({r['년-월'] for r in rows}, key=ym_sort_key)
    return yms[-1]


def sum_rows(rows):
    d = {m: 0 for m in METRICS}
    for r in rows:
        for m in METRICS:
            d[m] += r.get(m, 0) or 0
    return d


def group_by_store(rows):
    out = {}
    for r in rows:
        key = (str(r['매장코드']).zfill(5), r['매장명'])
        if key not in out:
            out[key] = {m: 0 for m in METRICS}
        for m in METRICS:
            out[key][m] += r[m]
    return out


def rows_for_period(rows, year, month_to):
    return [r for r in rows if r['년'] == year and 1 <= r['월'] <= month_to]


def rows_for_month_window(rows, end_ym, months=20):
    yms = sorted({r['년-월'] for r in rows}, key=ym_sort_key)
    if end_ym not in yms:
        end_ym = yms[-1]
    end_idx = yms.index(end_ym)
    selected = set(yms[max(0, end_idx - months + 1):end_idx + 1])
    return [r for r in rows if r['년-월'] in selected], sorted(selected, key=ym_sort_key)


def _apply_table_style(ws, start_row, start_col, end_row, end_col, header_rows=1, font_size=9):
    for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
        for cell in row:
            cell.border = Border(left=BORDER, right=BORDER, top=BORDER, bottom=BORDER)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(name='맑은 고딕', size=font_size)
    for r in range(start_row, start_row + header_rows):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(r, c)
            cell.fill = PatternFill('solid', fgColor=BLUE2)
            cell.font = Font(name='맑은 고딕', size=font_size, bold=True)
    for c in range(start_col, end_col + 1):
        ws.cell(start_row, c).border = Border(top=MED, left=BORDER, right=BORDER, bottom=BORDER)
        ws.cell(end_row, c).border = Border(bottom=MED, left=BORDER, right=BORDER, top=BORDER)


def _title(ws, title, last_col=10):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    cell = ws.cell(1, 1, title)
    cell.font = Font(name='맑은 고딕', size=16, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26
    ws.cell(2, last_col, '단위: 천원')
    ws.cell(2, last_col).font = Font(name='맑은 고딕', size=9)
    ws.cell(2, last_col).alignment = Alignment(horizontal='right')


def _num_fmt(ws, cell_range, pct=False):
    for row in ws[cell_range]:
        for cell in row:
            cell.number_format = '0.0%' if pct else '#,##0'
            if isinstance(cell.value, (int, float)) and cell.value < 0:
                cell.font = Font(name='맑은 고딕', size=9, color=RED)


def _autofit(ws, widths=None):
    if widths:
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
    else:
        for c in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(c)].width = 12
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'A4'



def add_dashboard(wb, rows, cur_year, cur_month, class_path=None):
    """경영분석보고서 첫 화면.

    V8 수정사항:
    - 금액 행과 비율 행의 표시 형식을 분리합니다.
    - 순이익 행이 퍼센트로 표시되는 문제를 수정합니다.
    - 순이익율 행만 % 형식으로 표시합니다.
    """
    ws = wb.create_sheet('Dashboard', 0)
    _title(ws, f'경영분석 Dashboard  {cur_year}년 1~{cur_month}월', 12)
    cur = sum_rows(rows_for_period(rows, cur_year, cur_month))
    prev = sum_rows(rows_for_period(rows, cur_year - 1, cur_month))

    headers = ['항목', f'{cur_year-1}', f'{cur_year}', '증감액', '증감률']
    header_row = 4
    first_data_row = 5
    ws.append([])
    ws.append(headers)

    kpis = ['판매금액', '수금액(V+)', '생산원가(V-)', '영업이익(V-)', '총경비', '순이익']
    for m in kpis:
        diff = cur[m] - prev[m]
        rate = _pct(diff, prev[m])
        ws.append([m, _k(prev[m]), _k(cur[m]), _k(diff), rate])

    profit_rate_prev = _vat_excl_rate(prev['순이익'], prev['판매금액'])
    profit_rate_cur = _vat_excl_rate(cur['순이익'], cur['판매금액'])
    ws.append(['순이익율', profit_rate_prev, profit_rate_cur, profit_rate_cur - profit_rate_prev, None])

    last_amount_row = first_data_row + len(kpis) - 1
    profit_rate_row = last_amount_row + 1
    _apply_table_style(ws, header_row, 1, profit_rate_row, 5)

    # 금액 행: B~D는 천원 단위 정수 표시
    _num_fmt(ws, f'B{first_data_row}:D{last_amount_row}')
    # 순이익율 행: B~D만 % 표시
    _num_fmt(ws, f'B{profit_rate_row}:D{profit_rate_row}', pct=True)
    # 증감률 열: 금액 KPI 행만 % 표시
    _num_fmt(ws, f'E{first_data_row}:E{last_amount_row}', pct=True)

    for r in range(first_data_row, profit_rate_row + 1):
        if isinstance(ws.cell(r, 4).value, (int, float)) and ws.cell(r, 4).value < 0:
            ws.cell(r, 4).font = Font(name='맑은 고딕', size=9, color=RED)

    ws['G3'] = '자동 코멘트'
    ws['G3'].fill = PatternFill('solid', fgColor=NAVY)
    ws['G3'].font = Font(name='맑은 고딕', color=WHITE, bold=True)

    comments = []
    sales_diff = cur['판매금액'] - prev['판매금액']
    profit_diff = cur['순이익'] - prev['순이익']
    comments.append(f"매출은 전년동기 대비 {_k(sales_diff):,}천원 ({_pct(sales_diff, prev['판매금액']):.1%}) 변동했습니다.")
    comments.append(f"순이익은 전년동기 대비 {_k(profit_diff):,}천원 ({_pct(profit_diff, prev['순이익']):.1%}) 변동했습니다.")
    comments.append(f"순이익율은 {profit_rate_cur:.1%}로 전년동기 대비 {(profit_rate_cur-profit_rate_prev):.1%}p 변동했습니다.")
    comments.append('상세 원인은 동일매장/신규매장/폐점매장/행사매장 분석 시트를 함께 확인하세요.')
    if class_path:
        comments.append(f'매장 구분 기준 파일: {class_path.name}')

    for i, txt in enumerate(comments, 4):
        ws.cell(i, 7, txt)
        ws.cell(i, 7).alignment = Alignment(wrap_text=True, vertical='top')
        try:
            ws.merge_cells(start_row=i, start_column=7, end_row=i, end_column=12)
        except ValueError:
            pass
        ws.row_dimensions[i].height = 32
        for c in range(7, 13):
            ws.cell(i, c).fill = PatternFill('solid', fgColor=LIGHT)
            ws.cell(i, c).border = Border(left=BORDER, right=BORDER, top=BORDER, bottom=BORDER)

    _autofit(ws, {'A':18, 'B':14, 'C':14, 'D':14, 'E':12, 'G':26, 'H':14, 'I':14, 'J':14, 'K':14, 'L':14})

def add_category_sheet(wb, rows, cur_year, cur_month, store_class, category, sheet_name):
    ws = wb.create_sheet(sheet_name)
    _title(ws, sheet_name, 12)
    headers = ['구분','매장코드','매장명']
    for m in SHORT_METRICS:
        headers += [m+'\n2025', '2026']
    ws.append(headers)
    prev = group_by_store(rows_for_period(rows, cur_year-1, cur_month))
    cur = group_by_store(rows_for_period(rows, cur_year, cur_month))
    store_names = {str(r['매장코드']).zfill(5): r['매장명'] for r in rows}
    items = class_items(store_class, category)
    # fallback: 매장구분 파일이 없을 때 기존 기준 사용
    if not items and category == '행사':
        items = [(code, {'name': store_names.get(code, ''), 'category': '행사', 'memo': memo}) for code, memo in EVENT_INFO.items()]
    if not items and category == '폐점':
        items = [(code, {'name': store_names.get(code, ''), 'category': '폐점', 'memo': memo}) for code, memo in CLOSED_INFO.items()]
    start = ws.max_row + 1
    total_prev = {m:0 for m in SHORT_METRICS}; total_cur={m:0 for m in SHORT_METRICS}
    zero = {m:0 for m in METRICS}
    for code, info in items:
        name = info.get('name') or store_names.get(code, '')
        p = next((v for (c,n),v in prev.items() if c==code), zero)
        c = next((v for (cc,n),v in cur.items() if cc==code), zero)
        row = [info.get('memo') or category, code, name]
        for m in SHORT_METRICS:
            row += [_k(p[m]), _k(c[m])]
            total_prev[m]+=p[m]; total_cur[m]+=c[m]
        ws.append(row)
    row = ['총합계','','']
    for m in SHORT_METRICS:
        row += [_k(total_prev[m]), _k(total_cur[m])]
    ws.append(row)
    end=ws.max_row
    _apply_table_style(ws, 3, 1, end, len(headers), header_rows=1)
    if start <= end:
        _num_fmt(ws, f'D{start}:M{end}')
    for c in range(1,len(headers)+1):
        ws.cell(end,c).fill = PatternFill('solid', fgColor=BLUE2)
        ws.cell(end,c).font = Font(name='맑은 고딕', bold=True)
    _autofit(ws, {'A':14,'B':10,'C':24,'D':12,'E':12,'F':12,'G':12,'H':12,'I':12,'J':12,'K':12,'L':12,'M':12})


def add_event_sheet(wb, rows, cur_year, cur_month, store_class):
    add_category_sheet(wb, rows, cur_year, cur_month, store_class, '행사', '행사매장분석')


def add_closed_sheet(wb, rows, cur_year, cur_month, store_class):
    add_category_sheet(wb, rows, cur_year, cur_month, store_class, '폐점', '폐점매장분석')


def add_new_store_sheet(wb, rows, cur_year, cur_month, store_class):
    add_category_sheet(wb, rows, cur_year, cur_month, store_class, '신규', '신규매장분석')


def add_same_store_sheet(wb, rows, cur_year, cur_month, store_class):
    ws = wb.create_sheet('동일매장분석')
    _title(ws, '동일매장분석', 18)
    headers = ['동일\n매장코드','매장명','매출금액\n2025','2026','당기기준\n신장률','수금액(V+)\n2025','2026','매출대비\n수금율','생산원가(V-)\n2025','2026','매출(V-)대비\n원가배수','총경비\n2025','2026','매출(V-)대비\n경비율','영업이익(V-)\n2025','2026','매출(V-)대비\n이익률','공헌\n이익률','증감액']
    ws.append(headers)
    prev = group_by_store(rows_for_period(rows, cur_year-1, cur_month))
    cur = group_by_store(rows_for_period(rows, cur_year, cur_month))
    keys = sorted(set(prev)|set(cur), key=lambda x:x[0])
    data=[]
    for key in keys:
        code,name=key
        p=prev.get(key,{m:0 for m in METRICS}); c=cur.get(key,{m:0 for m in METRICS})
        if classify_code(code, store_class, p, c) != '동일':
            continue
        if p['판매금액']>0 and c['판매금액']>0:
            diff = c['영업이익(V-)']-p['영업이익(V-)']
            data.append([code,name,_k(p['판매금액']),_k(c['판매금액']),_pct(c['판매금액']-p['판매금액'],p['판매금액']),_k(p['수금액(V+)']),_k(c['수금액(V+)']),_pct(c['수금액(V+)'],c['판매금액']),_k(p['생산원가(V-)']),_k(c['생산원가(V-)']),_cost_multiple(c['판매금액'],c['생산원가(V-)']),_k(p['총경비']),_k(c['총경비']),_vat_excl_rate(c['총경비'],c['판매금액']),_k(p['영업이익(V-)']),_k(c['영업이익(V-)']),_vat_excl_rate(c['영업이익(V-)'],c['판매금액']),_vat_excl_rate(c['순이익'],c['판매금액']),_k(diff)])
    data.sort(key=lambda x:x[15], reverse=True)
    for r in data: ws.append(r)
    # totals
    ptotal=sum_rows([r for r in rows_for_period(rows,cur_year-1,cur_month) if classify_code(str(r['매장코드']).zfill(5), store_class) == '동일'])
    ctotal=sum_rows([r for r in rows_for_period(rows,cur_year,cur_month) if classify_code(str(r['매장코드']).zfill(5), store_class) == '동일'])
    ws.append(['총합계','',_k(ptotal['판매금액']),_k(ctotal['판매금액']),_pct(ctotal['판매금액']-ptotal['판매금액'],ptotal['판매금액']),_k(ptotal['수금액(V+)']),_k(ctotal['수금액(V+)']),_pct(ctotal['수금액(V+)'],ctotal['판매금액']),_k(ptotal['생산원가(V-)']),_k(ctotal['생산원가(V-)']),_cost_multiple(ctotal['판매금액'],ctotal['생산원가(V-)']),_k(ptotal['총경비']),_k(ctotal['총경비']),_vat_excl_rate(ctotal['총경비'],ctotal['판매금액']),_k(ptotal['영업이익(V-)']),_k(ctotal['영업이익(V-)']),_vat_excl_rate(ctotal['영업이익(V-)'],ctotal['판매금액']),_vat_excl_rate(ctotal['순이익'],ctotal['판매금액']),_k(ctotal['영업이익(V-)']-ptotal['영업이익(V-)'])])
    end=ws.max_row
    _apply_table_style(ws,3,1,end,len(headers))
    for rng in ['C4:D{}'.format(end),'F4:G{}'.format(end),'I4:J{}'.format(end),'L4:M{}'.format(end),'O4:P{}'.format(end),'S4:S{}'.format(end)]: _num_fmt(ws,rng)
    for rng in ['E4:E{}'.format(end),'H4:H{}'.format(end),'K4:K{}'.format(end),'N4:N{}'.format(end),'Q4:R{}'.format(end)]: _num_fmt(ws,rng,pct=True)
    for c in range(1,len(headers)+1): ws.cell(end,c).fill=PatternFill('solid',fgColor=BLUE2); ws.cell(end,c).font=Font(name='맑은 고딕',bold=True)
    ws.auto_filter.ref = f'A3:S{end}'
    _autofit(ws, {get_column_letter(i):10 for i in range(1,20)})
    ws.column_dimensions['B'].width=22


def add_yoy_summary_sheet(wb, rows, cur_year, cur_month, store_class):
    ws = wb.create_sheet('전년동기대비')
    _title(ws, '전년동기대비 증감액 및 증감율', 14)
    prev_by = group_by_store(rows_for_period(rows, cur_year-1, cur_month))
    cur_by = group_by_store(rows_for_period(rows, cur_year, cur_month))
    categories = {'동일':[], '신규':[], '폐점':[], '행사':[]}
    for key in set(prev_by)|set(cur_by):
        code=key[0]
        p=prev_by.get(key,{m:0 for m in METRICS}); c=cur_by.get(key,{m:0 for m in METRICS})
        cat = classify_code(code, store_class, p, c)
        categories[cat].append((p,c))
    headers=['구분']
    for m in SHORT_METRICS: headers += [m+'\n2025', '2026']
    ws.append(headers)
    totals_prev={m:0 for m in SHORT_METRICS}; totals_cur={m:0 for m in SHORT_METRICS}
    for cat in ['동일','신규','폐점','행사']:
        ptotal={m:0 for m in SHORT_METRICS}; ctotal={m:0 for m in SHORT_METRICS}
        for p,c in categories[cat]:
            for m in SHORT_METRICS: ptotal[m]+=p[m]; ctotal[m]+=c[m]
        row=[cat]
        for m in SHORT_METRICS:
            row += [_k(ptotal[m]), _k(ctotal[m])]
            totals_prev[m]+=ptotal[m]; totals_cur[m]+=ctotal[m]
        ws.append(row)
    row=['총합계']
    for m in SHORT_METRICS: row += [_k(totals_prev[m]), _k(totals_cur[m])]
    ws.append(row)
    # difference/contribution section
    start2 = ws.max_row + 3
    ws.cell(start2,1,f'전기대비 매출 증가분의 당기 매출기여도')
    ws.cell(start2,1).fill=PatternFill('solid',fgColor=BLUE2); ws.cell(start2,1).font=Font(name='맑은 고딕',bold=True)
    headers2=['구분']
    for m in SHORT_METRICS: headers2 += [m+'차이', '증감률']
    ws.append([]); ws.append(headers2)
    total_diff={m: totals_cur[m]-totals_prev[m] for m in SHORT_METRICS}
    for cat in ['동일','신규','폐점','행사','총합계']:
        if cat=='총합계': ptotal=totals_prev; ctotal=totals_cur
        else:
            ptotal={m:0 for m in SHORT_METRICS}; ctotal={m:0 for m in SHORT_METRICS}
            for p,c in categories[cat]:
                for m in SHORT_METRICS: ptotal[m]+=p[m]; ctotal[m]+=c[m]
        row=[cat]
        for m in SHORT_METRICS:
            diff=ctotal[m]-ptotal[m]
            row += [_k(diff), _pct(diff, ptotal[m])]
        ws.append(row)
    end=ws.max_row
    _apply_table_style(ws,3,1,8,len(headers))
    _apply_table_style(ws,start2+2,1,end,len(headers2))
    _num_fmt(ws,'B4:K8')
    _num_fmt(ws,f'B{start2+3}:J{end}')
    for c in range(3,12,2): _num_fmt(ws,f'{get_column_letter(c)}{start2+3}:{get_column_letter(c)}{end}', pct=True)
    _autofit(ws, {get_column_letter(i):12 for i in range(1,15)})


def add_expense_ratio_trend(wb, rows, end_ym):
    ws=wb.create_sheet('월별경비율추이')
    _title(ws, '월별경비율추이', 18)
    window, yms = rows_for_month_window(rows, end_ym, 20)
    by=defaultdict(lambda: defaultdict(lambda:{'sales':0,'exp':0}))
    for r in window:
        key=(str(r['매장코드']).zfill(5), r['매장명'])
        by[key][r['년-월']]['sales']+=r['판매금액']; by[key][r['년-월']]['exp']+=r['총경비']
    headers=['매장코드','매장명']+yms+['합계']
    ws.append(headers)
    data=[]
    for key,dic in by.items():
        vals=[]; s_tot=e_tot=0
        for ym in yms:
            s=dic[ym]['sales']; e=dic[ym]['exp']; vals.append(_pct(e,s)); s_tot+=s; e_tot+=e
        data.append([key[0],key[1]]+vals+[_pct(e_tot,s_tot)])
    data.sort(key=lambda x:x[-1], reverse=True)
    for r in data: ws.append(r)
    # total row
    total=['합계','']
    for ym in yms:
        s=sum(r['판매금액'] for r in window if r['년-월']==ym); e=sum(r['총경비'] for r in window if r['년-월']==ym)
        total.append(_pct(e,s))
    s=sum(r['판매금액'] for r in window); e=sum(r['총경비'] for r in window); total.append(_pct(e,s))
    ws.append(total)
    end=ws.max_row
    _apply_table_style(ws,3,1,end,len(headers))
    _num_fmt(ws, f'C4:{get_column_letter(len(headers))}{end}', pct=True)
    ws.auto_filter.ref=f'A3:{get_column_letter(len(headers))}{end}'
    for c in range(1,len(headers)+1): ws.cell(end,c).fill=PatternFill('solid',fgColor=BLUE2); ws.cell(end,c).font=Font(name='맑은 고딕',bold=True)
    _autofit(ws, {'A':10,'B':22, **{get_column_letter(c):10 for c in range(3,len(headers)+1)}})


def add_rank_sheet(wb, rows, latest_ym, sheet_name, sort_key, ascending=True):
    ws=wb.create_sheet(sheet_name)
    _title(ws, sheet_name, 12)
    latest=[r for r in rows if r['년-월']==latest_ym]
    data=[]
    for r in latest:
        sales=r['판매금액']; exp=r['총경비']; profit=r['순이익']; op=r['영업이익(V-)']
        data.append([r['년'], f'{r["월"]:02d}', r['년-월'], str(r['매장코드']).zfill(5), r['매장명'], _k(sales), _k(r['수금액(V+)']), _k(r['수금액(V-)']), _k(r['생산원가(V-)']), _k(op), _k(exp), _k(profit), _vat_excl_rate(op-exp, sales), _vat_excl_rate(profit, sales), _vat_excl_rate(exp, sales)])
    key_idx={'순이익':11,'경비율':14,'영업이익':9}.get(sort_key,11)
    data.sort(key=lambda x:x[key_idx], reverse=not ascending)
    headers=['년','월','년-월','매장코드','매장명','매출금액','수금액(V+)','수금액(V-)','생산원가(V-)','영업이익(V-)','총경비','순이익','영업이익-총경비\n(손익)','순이익율','경비율']
    ws.append(headers)
    for row in data: ws.append(row)
    end=ws.max_row
    _apply_table_style(ws,3,1,end,len(headers))
    _num_fmt(ws,f'F4:M{end}')
    _num_fmt(ws,f'N4:O{end}',pct=True)
    ws.auto_filter.ref=f'A3:O{end}'
    _autofit(ws, {'A':7,'B':6,'C':10,'D':10,'E':24,'F':12,'G':12,'H':12,'I':12,'J':12,'K':12,'L':12,'M':12,'N':10,'O':10})


def add_year_month_profit(wb, rows, end_year):
    ws=wb.create_sheet('연도별월별영업순이익')
    _title(ws,'년도별 월별 총합계 영업 순이익', 12)
    years=[y for y in sorted({r['년'] for r in rows}) if 2023 <= y <= end_year]
    profits=defaultdict(lambda:defaultdict(int))
    for r in rows:
        if r['년'] in years:
            profits[r['월']][r['년']] += r['순이익']
    ws.append(['월']+years+['평균'])
    for m in range(1,13):
        vals=[_k(profits[m][y]) if profits[m][y] else None for y in years]
        avg_vals=[v for v in vals if v is not None]
        ws.append([m]+vals+[round(sum(avg_vals)/len(avg_vals)) if avg_vals else None])
    ws.append(['합계']+[sum(_k(profits[m][y]) for m in range(1,13)) for y in years]+[None])
    end=ws.max_row
    _apply_table_style(ws,3,1,end,2+len(years))
    _num_fmt(ws,f'B4:{get_column_letter(2+len(years))}{end}')
    # cumulative section
    start2=end+6
    ws.cell(start2,1,'년도별 월별 총합계 영업 순이익(누적)')
    ws.cell(start2,1).font=Font(name='맑은 고딕',bold=True,size=13)
    ws.cell(start2+2,1,'월')
    for i,y in enumerate(years,2): ws.cell(start2+2,i,y)
    ws.cell(start2+2,2+len(years),'평균')
    for m in range(1,13):
        ws.cell(start2+2+m,1,m)
        cum=[]
        for i,y in enumerate(years,2):
            val=sum(profits[mm][y] for mm in range(1,m+1))
            ws.cell(start2+2+m,i,_k(val) if val else None); cum.append(_k(val) if val else None)
        av=[v for v in cum if v is not None]
        ws.cell(start2+2+m,2+len(years),round(sum(av)/len(av)) if av else None)
    _apply_table_style(ws,start2+2,1,start2+14,2+len(years))
    _num_fmt(ws,f'B{start2+3}:{get_column_letter(2+len(years))}{start2+14}')
    # charts
    chart = BarChart(); chart.type='col'; chart.style=10; chart.title='월별 영업 순이익'; chart.y_axis.title='천원'; chart.x_axis.title='월'
    data=Reference(ws,min_col=2,max_col=1+len(years),min_row=3,max_row=15)
    cats=Reference(ws,min_col=1,min_row=4,max_row=15)
    chart.add_data(data,titles_from_data=True); chart.set_categories(cats); chart.height=9; chart.width=18; chart.legend.position='b'
    ws.add_chart(chart,'H3')
    chart2=LineChart(); chart2.title='누적 영업 순이익'; chart2.y_axis.title='천원'; chart2.x_axis.title='월'; chart2.height=9; chart2.width=18
    data2=Reference(ws,min_col=2,max_col=1+len(years),min_row=start2+2,max_row=start2+14)
    cats2=Reference(ws,min_col=1,min_row=start2+3,max_row=start2+14)
    chart2.add_data(data2,titles_from_data=True); chart2.set_categories(cats2); chart2.legend.position='r'
    ws.add_chart(chart2,f'H{start2+2}')
    _autofit(ws, {'A':14, **{get_column_letter(c):14 for c in range(2,2+len(years)+1)}})


def add_quarter_season(wb, rows, cur_year):
    for sheet_name, key in [('분기분석','분기'),('시즌분석','시즌명')]:
        ws=wb.create_sheet(sheet_name)
        _title(ws, sheet_name, 12)
        headers=[key,'판매금액','수금액(V+)','생산원가(V-)','영업이익(V-)','총경비','순이익','순이익율','경비율']
        ws.append(headers)
        groups=defaultdict(list)
        for r in rows:
            if r['년']>=cur_year-2:
                groups[r[key]].append(r)
        for g in sorted(groups.keys(), key=lambda x: str(x)):
            s=sum_rows(groups[g])
            ws.append([g,_k(s['판매금액']),_k(s['수금액(V+)']),_k(s['생산원가(V-)']),_k(s['영업이익(V-)']),_k(s['총경비']),_k(s['순이익']),_vat_excl_rate(s['순이익'],s['판매금액']),_vat_excl_rate(s['총경비'],s['판매금액'])])
        end=ws.max_row
        _apply_table_style(ws,3,1,end,len(headers))
        _num_fmt(ws,f'B4:G{end}')
        _num_fmt(ws,f'H4:I{end}',pct=True)
        _autofit(ws, {'A':12,'B':14,'C':14,'D':14,'E':14,'F':14,'G':14,'H':10,'I':10})


def make_management_report(base_dir: Path = None, out_path: Path = None):
    base_dir = Path(base_dir or Path(__file__).parent)
    rows = load_db(base_dir)
    latest_ym = latest_year_month(rows)
    cur_year, cur_month = ym_sort_key(latest_ym)
    store_class, class_path = load_store_classification(base_dir, latest_ym)
    out_dir = base_dir / 'output'
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = Path(out_path or out_dir / '경영분석보고서.xlsx')
    wb = Workbook()
    wb.remove(wb.active)
    add_dashboard(wb, rows, cur_year, cur_month, class_path)
    add_yoy_summary_sheet(wb, rows, cur_year, cur_month, store_class)
    add_same_store_sheet(wb, rows, cur_year, cur_month, store_class)
    add_event_sheet(wb, rows, cur_year, cur_month, store_class)
    add_closed_sheet(wb, rows, cur_year, cur_month, store_class)
    add_new_store_sheet(wb, rows, cur_year, cur_month, store_class)
    add_rank_sheet(wb, rows, latest_ym, '영업이익높은순', '영업이익', ascending=False)
    add_rank_sheet(wb, rows, latest_ym, '순이익낮은순', '순이익', ascending=True)
    add_rank_sheet(wb, rows, latest_ym, '경비율높은순', '경비율', ascending=False)
    add_expense_ratio_trend(wb, rows, latest_ym)
    add_year_month_profit(wb, rows, cur_year)
    add_quarter_season(wb, rows, cur_year)
    # freeze panes and page setup
    for ws in wb.worksheets:
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = 9
        ws.page_margins.left = 0.3; ws.page_margins.right = 0.3; ws.page_margins.top = 0.5; ws.page_margins.bottom = 0.5
    wb.save(out_path)
    print(f'✅ 경영분석보고서 생성 완료: {out_path}')
    print(f"   매장구분 기준: {class_path.name if class_path else '기존 자동/legacy 기준'}")
    return out_path


def run(base_dir: Path = None):
    return make_management_report(base_dir or Path(__file__).parent)


if __name__ == '__main__':
    run(Path(__file__).parent)
