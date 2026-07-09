# -*- coding: utf-8 -*-
"""
경영분석 보고서 생성 모듈 v6
- 기존 경영진 보고 양식 참조
- 2026년 5월까지는 DB/월별손익DB.xlsx의 기존 DB 수치 사용
- 2026년 6월부터 history_update.py로 업데이트된 DB 수치 사용
- 단위: Dashboard/연도별월별영업순이익/시즌분석은 천원, 월정산보고/경비율높은순/
  순이익낮은순/월별경비율추이는 원 단위 (템플릿 참조)
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


def _k(v):
    return int(round((v or 0) / UNIT))


def _pct(n, d):
    return n / d if d else 0


def _sales_vat_excl(sales_vat_incl):
    # 매출금액은 V+ 기준입니다. 생산원가/총경비/영업이익/순이익은 V- 기준이므로
    # 배수와 율 산출 시 매출을 V- 기준(매출/1.1)으로 환산합니다.
    return sales_vat_incl / 1.1 if sales_vat_incl else 0


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


def rows_for_period(rows, year, month_to):
    return [r for r in rows if r['년'] == year and 1 <= r['월'] <= month_to]


def rows_for_month_window(rows, end_ym, months=20):
    yms = sorted({r['년-월'] for r in rows}, key=ym_sort_key)
    if end_ym not in yms:
        end_ym = yms[-1]
    end_idx = yms.index(end_ym)
    selected = set(yms[max(0, end_idx - months + 1):end_idx + 1])
    return [r for r in rows if r['년-월'] in selected], sorted(selected, key=ym_sort_key)


def _store_month_map(window):
    """[{매장코드,매장명,년-월,...}, ...] -> {(코드,매장명): {년-월: row}}"""
    m = {}
    for r in window:
        key = (str(r['매장코드']).zfill(5), r['매장명'])
        m.setdefault(key, {})[r['년-월']] = r
    return m


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


def _title(ws, title, last_col=10, unit='천원'):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    cell = ws.cell(1, 1, title)
    cell.font = Font(name='맑은 고딕', size=16, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26
    # unit이 없어도 2행을 비워서라도 만들어 둬야 헤더가 항상 3행에 오는 기존 관례가 유지된다
    ws.cell(2, last_col, f'단위: {unit}' if unit else '')
    if unit:
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



def add_dashboard(wb, rows, cur_year, cur_month):
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

def add_monthly_settlement_sheet(wb, rows, latest_ym, cur_year, cur_month):
    """월정산보고: 최신 1개월, 매장코드순 매장별 정산 스냅샷 (원 단위)"""
    ws = wb.create_sheet('월정산보고')
    _title(ws, f'{cur_year}년 {cur_month:02d}월 백화점 매장 정산 보고(매장코드순)', 13, unit='원')
    headers = ['매장코드','매장명','매출금액','수금액(V+)','점수수료','수수료율','수금액(V-)',
               '생산원가(V-)','영업이익(V-)','총경비','순이익','영업이익_순이익율','영업이익_경비율']
    ws.append(headers)
    latest = sorted([r for r in rows if r['년-월'] == latest_ym], key=lambda r: str(r['매장코드']).zfill(5))
    data = []
    for r in latest:
        sales=r['판매금액']; vp=r['수금액(V+)']; vn=r['수금액(V-)']; cg=r['생산원가(V-)']
        op=r['영업이익(V-)']; exp=r['총경비']; profit=r['순이익']
        commission = sales - vp
        data.append([str(r['매장코드']).zfill(5), r['매장명'], sales, vp, commission,
                     _pct(commission, sales), vn, cg, op, exp, profit,
                     _pct(profit, op), _pct(exp, op)])
    for row in data: ws.append(row)
    tot_sales=sum(x[2] for x in data); tot_vp=sum(x[3] for x in data); tot_commission=sum(x[4] for x in data)
    tot_cg=sum(x[7] for x in data); tot_exp=sum(x[9] for x in data)
    tot_vn = round(tot_vp/1.1) if tot_vp else 0
    tot_op = tot_vn - tot_cg
    tot_profit = tot_op - tot_exp
    ws.append(['합계','', tot_sales, tot_vp, tot_commission, None, tot_vn, tot_cg, tot_op, tot_exp,
               tot_profit, _pct(tot_profit, tot_op), _pct(tot_exp, tot_op)])
    end = ws.max_row
    _apply_table_style(ws, 3, 1, end, len(headers))
    _num_fmt(ws, f'C4:E{end}')
    _num_fmt(ws, f'F4:F{end}', pct=True)
    _num_fmt(ws, f'G4:K{end}')
    _num_fmt(ws, f'L4:M{end}', pct=True)
    ws.auto_filter.ref = f'A3:M{end}'
    for c in range(1, len(headers)+1):
        ws.cell(end, c).fill = PatternFill('solid', fgColor=BLUE2)
        ws.cell(end, c).font = Font(name='맑은 고딕', bold=True)
    store_count = len(data)
    avg_sales = round(tot_sales/store_count) if store_count else 0
    stat_row = end + 3
    for i, (label, val) in enumerate([('총매출', tot_sales), ('매장수', store_count), ('평균매출', avg_sales)]):
        rr = stat_row + i
        ws.cell(rr, 2, label).font = Font(name='맑은 고딕', bold=True)
        ws.cell(rr, 3, val).number_format = '#,##0'
    _autofit(ws, {'A':10,'B':22,'C':14,'D':14,'E':14,'F':10,'G':14,'H':14,'I':14,'J':14,'K':14,'L':10,'M':10})


def add_expense_ratio_trend(wb, rows, end_ym):
    """월별경비율추이: 최근 12개월, 경비율 = 총경비/영업이익(V-), 원 데이터는 매장×월 매트릭스 공유"""
    ws=wb.create_sheet('월별경비율추이')
    _title(ws, '월별경비율추이(총경비/영업이익)', 15, unit=None)
    window, yms = rows_for_month_window(rows, end_ym, 12)
    smap = _store_month_map(window)
    headers=['매장코드','매장명']+yms+['합계']
    ws.append(headers)
    data=[]
    for key,dic in smap.items():
        vals=[]; exp_tot=op_tot=0
        for ym in yms:
            r = dic.get(ym)
            exp = r['총경비'] if r else 0
            op = r['영업이익(V-)'] if r else 0
            vals.append(_pct(exp, op)); exp_tot+=exp; op_tot+=op
        data.append([key[0],key[1]]+vals+[_pct(exp_tot,op_tot)])
    data.sort(key=lambda x:x[-1], reverse=True)
    for r in data: ws.append(r)
    # total row: 총경비합계/영업이익합계
    total=['총경비합계/영업이익합계, %','']
    for ym in yms:
        exp=sum(r['총경비'] for r in window if r['년-월']==ym); op=sum(r['영업이익(V-)'] for r in window if r['년-월']==ym)
        total.append(_pct(exp,op))
    exp_all=sum(r['총경비'] for r in window); op_all=sum(r['영업이익(V-)'] for r in window)
    total.append(_pct(exp_all,op_all))
    ws.append(total)
    end=ws.max_row
    _apply_table_style(ws,3,1,end,len(headers))
    _num_fmt(ws, f'C4:{get_column_letter(len(headers))}{end}', pct=True)
    ws.auto_filter.ref=f'A3:{get_column_letter(len(headers))}{end}'
    for c in range(1,len(headers)+1): ws.cell(end,c).fill=PatternFill('solid',fgColor=BLUE2); ws.cell(end,c).font=Font(name='맑은 고딕',bold=True)
    _autofit(ws, {'A':10,'B':22, **{get_column_letter(c):10 for c in range(3,len(headers)+1)}})


def add_profit_trend_sheet(wb, rows, end_ym):
    """순이익낮은순: 최근 12개월 매장별 순이익 원 단위 추이, 12개월 합계 오름차순"""
    ws=wb.create_sheet('순이익낮은순')
    _title(ws, '순이익낮은순(최근 12개월)', 15, unit='원')
    window, yms = rows_for_month_window(rows, end_ym, 12)
    smap = _store_month_map(window)
    headers=['매장코드','매장명']+yms+['합계']
    ws.append(headers)
    data=[]
    for key,dic in smap.items():
        vals=[dic[ym]['순이익'] if ym in dic else 0 for ym in yms]
        data.append([key[0],key[1]]+vals+[sum(vals)])
    data.sort(key=lambda x:x[-1])
    for r in data: ws.append(r)
    n = len(data)
    sum_row=['이익액 합계','']; avg_row=['이익액 평균','']
    for i in range(len(yms)):
        col = [d[2+i] for d in data]
        sum_row.append(sum(col)); avg_row.append(round(sum(col)/n) if n else 0)
    tot_col = [d[-1] for d in data]
    sum_row.append(sum(tot_col)); avg_row.append(round(sum(tot_col)/n) if n else 0)
    ws.append(sum_row); ws.append(avg_row)
    end=ws.max_row
    _apply_table_style(ws,3,1,end,len(headers))
    _num_fmt(ws, f'C4:{get_column_letter(len(headers))}{end}')
    ws.auto_filter.ref=f'A3:{get_column_letter(len(headers))}{end}'
    for rr in (end-1, end):
        for c in range(1,len(headers)+1): ws.cell(rr,c).fill=PatternFill('solid',fgColor=BLUE2); ws.cell(rr,c).font=Font(name='맑은 고딕',bold=True)
    _autofit(ws, {'A':10,'B':22, **{get_column_letter(c):12 for c in range(3,len(headers)+1)}})


def add_rank_sheet(wb, rows, latest_ym):
    """경비율높은순: 최신 1개월, 총경비율 내림차순, 원 단위"""
    sheet_name = '경비율높은순'
    ws=wb.create_sheet(sheet_name)
    _title(ws, sheet_name, 11, unit='원')
    latest=[r for r in rows if r['년-월']==latest_ym]
    data=[]
    for r in latest:
        sales=r['판매금액']; vp=r['수금액(V+)']; vn=r['수금액(V-)']; cg=r['생산원가(V-)']
        op=r['영업이익(V-)']; exp=r['총경비']; profit=r['순이익']
        data.append([str(r['매장코드']).zfill(5), r['매장명'], sales, vp, vn, cg, op, exp, profit,
                     _vat_excl_rate(profit, sales), _vat_excl_rate(exp, sales)])
    data.sort(key=lambda x:x[10], reverse=True)
    headers=['매장코드','매장명','매출금액(V+)','수금액(V+)','수금액(V-)','생산원가(V-)',
             '영업이익(V-)','총경비','순이익(영업이익-총경비)','순이익율','총경비율']
    ws.append(headers)
    for row in data: ws.append(row)
    tot_sales=sum(x[2] for x in data); tot_vp=sum(x[3] for x in data); tot_vn=sum(x[4] for x in data)
    tot_cg=sum(x[5] for x in data); tot_op=sum(x[6] for x in data); tot_exp=sum(x[7] for x in data)
    tot_profit=sum(x[8] for x in data)
    ws.append(['합계','', tot_sales, tot_vp, tot_vn, tot_cg, tot_op, tot_exp, tot_profit,
               _vat_excl_rate(tot_profit, tot_sales), _vat_excl_rate(tot_exp, tot_sales)])
    end=ws.max_row
    _apply_table_style(ws,3,1,end,len(headers))
    _num_fmt(ws,f'C4:I{end}')
    _num_fmt(ws,f'J4:K{end}',pct=True)
    ws.auto_filter.ref=f'A3:K{end}'
    for c in range(1,len(headers)+1): ws.cell(end,c).fill=PatternFill('solid',fgColor=BLUE2); ws.cell(end,c).font=Font(name='맑은 고딕',bold=True)
    _autofit(ws, {'A':10,'B':22,'C':14,'D':14,'E':14,'F':14,'G':14,'H':14,'I':14,'J':10,'K':10})


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


def add_season_sheet(wb, rows, cur_year):
    ws=wb.create_sheet('시즌분석')
    _title(ws, '시즌분석', 12)
    headers=['시즌명','판매금액','수금액(V+)','생산원가(V-)','영업이익(V-)','총경비','순이익','순이익율','경비율']
    ws.append(headers)
    groups=defaultdict(list)
    for r in rows:
        if r['년']>=cur_year-2:
            groups[r['시즌명']].append(r)
    for g in sorted(groups.keys(), key=lambda x: str(x)):
        s=sum_rows(groups[g])
        ws.append([g,_k(s['판매금액']),_k(s['수금액(V+)']),_k(s['생산원가(V-)']),_k(s['영업이익(V-)']),_k(s['총경비']),_k(s['순이익']),_vat_excl_rate(s['순이익'],s['판매금액']),_vat_excl_rate(s['총경비'],s['판매금액'])])
    end=ws.max_row
    _apply_table_style(ws,3,1,end,len(headers))
    _num_fmt(ws,f'B4:G{end}')
    _num_fmt(ws,f'H4:I{end}',pct=True)
    _autofit(ws, {'A':12,'B':14,'C':14,'D':14,'E':14,'F':14,'G':14,'H':10,'I':10})


def make_management_report(base_dir: Path = None, out_path: Path = None, ym: str = None):
    base_dir = Path(base_dir or Path(__file__).parent)
    rows = load_db(base_dir)
    latest_ym = latest_year_month(rows)
    cur_year, cur_month = ym_sort_key(latest_ym)
    out_dir = base_dir / 'output'
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = Path(out_path or out_dir / '경영분석보고서.xlsx')
    wb = Workbook()
    wb.remove(wb.active)
    add_dashboard(wb, rows, cur_year, cur_month)
    add_monthly_settlement_sheet(wb, rows, latest_ym, cur_year, cur_month)
    add_year_month_profit(wb, rows, cur_year)
    add_rank_sheet(wb, rows, latest_ym)
    add_profit_trend_sheet(wb, rows, latest_ym)
    add_expense_ratio_trend(wb, rows, latest_ym)
    add_season_sheet(wb, rows, cur_year)
    # freeze panes and page setup
    for ws in wb.worksheets:
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = 9
        ws.page_margins.left = 0.3; ws.page_margins.right = 0.3; ws.page_margins.top = 0.5; ws.page_margins.bottom = 0.5
    wb.save(out_path)
    print(f'✅ 경영분석보고서 생성 완료: {out_path}')
    if ym:
        try:
            import shutil
            month_out = base_dir/'output'/ym/'경영분석보고서.xlsx'
            month_out.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(out_path, month_out)
            print(f'   ↳ 월별 사본: {month_out}')
        except Exception as e:
            print(f'   ⚠️ 월별 사본 저장 오류: {e}')
    return out_path


def run(base_dir: Path = None, ym: str = None):
    return make_management_report(base_dir or Path(__file__).parent, ym=ym)


if __name__ == '__main__':
    run(Path(__file__).parent)
