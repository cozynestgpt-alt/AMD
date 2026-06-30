# -*- coding: utf-8 -*-
"""salary_system v3 통합 실행 파일
기존 run.py를 실행한 뒤, run.py 내부에서 월별손익DB와 분기/시즌 보고서까지 자동 업데이트합니다.
"""
import run

if __name__ == '__main__':
    run.main()
    input('\nEnter 키를 눌러 종료합니다...')
