# AMD V8.0.1 Git 관리 정책

## 1. 원칙

GitHub에는 프로그램과 문서만 관리합니다.

포함 대상:

- Python 소스
- BAT 실행 파일
- README / CHANGELOG / LICENSE / CONTRIBUTING
- docs 문서
- templates 양식
- `.gitignore`
- `config.yaml`

제외 대상:

- 월별 input 실무자료
- output 산출 보고서
- backup 백업자료
- 실제 월별손익DB.xlsx

## 2. 제외 이유

`input`, `output`, `backup`, `DB`에는 매출, 수수료, 원가, 순이익, 급여성 자료 등 민감한 경영정보가 포함될 수 있습니다.

따라서 GitHub에는 올리지 않고, NAS 또는 사내 보안 저장소에서 별도 관리하는 것이 안전합니다.

## 3. 폴더 유지

Git은 빈 폴더를 관리하지 않으므로 각 폴더에 `.gitkeep` 파일을 둡니다.

- input/.gitkeep
- output/.gitkeep
- backup/.gitkeep
- DB/.gitkeep

## 4. 이미 Git에 올라간 파일 처리

`.gitignore`를 추가해도 이미 Git에 등록된 파일은 자동으로 제외되지 않습니다.

기존에 Git에 올라간 업무 데이터는 아래 명령으로 Git 추적만 해제해야 합니다.

```bat
git rm -r --cached input output backup DB
git add input/.gitkeep output/.gitkeep backup/.gitkeep DB/.gitkeep
git add .gitignore config.yaml docs/CONFIG_POLICY.md
```

이 명령은 로컬 파일을 삭제하지 않고, GitHub 추적 대상에서만 제외합니다.
