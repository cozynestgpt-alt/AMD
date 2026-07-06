# -*- coding: utf-8 -*-
"""
salary_system V4 점검 스크립트
- Python 코드 내 템플릿 참조 확인
- 월별입력/중간관리 이체 템플릿 사용 여부 확인
- 주요 xlsx 파일의 외부 링크 존재 여부 확인
사용법: python audit_system.py
"""
from pathlib import Path
import zipfile

BASE = Path(__file__).parent

REQUIRED_TEMPLATES = [
    "개별통지문_템플릿.xlsx",
    "판매수수료집계_템플릿.xlsx",
    "세금계산서출력_템플릿.xlsx",
]
NOT_USED_TEMPLATES = [
    "중간관리판매수수료이체내역_템플릿.xlsx",
    "월별입력_템플릿.xlsx",
]


def scan_code():
    print("\n[1] Python 코드 템플릿 참조 검사")
    py_files = [p for p in sorted(BASE.glob("*.py")) if p.name != "audit_system.py"]
    for template in REQUIRED_TEMPLATES + NOT_USED_TEMPLATES:
        hits = []
        for p in py_files:
            text = p.read_text(encoding="utf-8", errors="ignore")
            if template in text:
                hits.append(p.name)
        print(f"  - {template}: {', '.join(hits) if hits else '참조 없음'}")


def has_external_links(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path, "r") as z:
            names = z.namelist()
            return any(n.startswith("xl/externalLinks/") for n in names)
    except Exception:
        return False


def scan_xlsx_external_links():
    print("\n[2] xlsx 외부 링크 검사")
    targets = []
    db = BASE / "DB" / "월별손익DB.xlsx"
    if db.exists():
        targets.append(db)
    out_report = BASE / "output" / "연도별_분기별_시즌별_손익분석.xlsx"
    if out_report.exists():
        targets.append(out_report)

    for path in targets:
        result = "외부 링크 있음" if has_external_links(path) else "외부 링크 없음"
        print(f"  - {path.relative_to(BASE)}: {result}")


def check_template_files():
    print("\n[3] templates 폴더 필수 파일 존재 검사")
    tdir = BASE / "templates"
    for name in REQUIRED_TEMPLATES:
        print(f"  - {name}: {'있음' if (tdir / name).exists() else '없음'}")


def main():
    print("=" * 60)
    print(" salary_system V4 audit")
    print("=" * 60)
    scan_code()
    check_template_files()
    scan_xlsx_external_links()
    print("\n점검 완료")


if __name__ == "__main__":
    main()
