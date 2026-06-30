기존 salary_system의 templates 폴더에서 아래 3개 템플릿 파일만 복사하세요.

[필수 템플릿]
1. 개별통지문_템플릿.xlsx
   - 사용 파일: run.py
   - 생성 결과: output/YYYY-MM/개별통지문.xlsx

2. 판매수수료집계_템플릿.xlsx
   - 사용 파일: sales_summary.py
   - 생성 결과: output/YYYY-MM/판매수수료집계.xlsx

3. 세금계산서출력_템플릿.xlsx
   - 사용 파일: tax_invoice_report.py
   - 생성 결과: output/YYYY-MM/세금계산서_YYYY-MM.xlsx

[현재 코드 기준으로 불필요한 템플릿]
- 중간관리판매수수료이체내역_템플릿.xlsx
  사유: transfer_report.py가 openpyxl.Workbook()으로 이체내역서를 직접 생성하므로 템플릿을 읽지 않습니다.

- 월별입력_템플릿.xlsx
  사유: V3/V4에서는 월별입력 파일을 완전히 폐지했습니다.

- 사원마스터_템플릿.xlsx
  사유: 실행 시 사원마스터.xlsx 입력 파일은 사용하지만, 템플릿 파일은 코드에서 읽지 않습니다.

- 매출_템플릿.xlsx
  사유: 영*판매*.xlsx 입력 파일은 사용하지만, 매출 템플릿 파일은 코드에서 읽지 않습니다.

※ 위 판단은 V4 코드 전수 검색 결과 기준입니다.
