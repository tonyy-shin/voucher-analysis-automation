# 전표분석서 자동화

Excel 원본 데이터(사업비 피벗 테이블)에서 출력전표 코드별로 **사업비 전표분석서 PDF**를 일괄 자동 생성하는 Windows 도구.

---

## 주요 기능

- **출력전표 코드 일괄 처리** — `출력>` 시트에 등록된 코드를 순회하며 코드별 PDF를 자동 생성
- **3중 LEFT JOIN** — 실제발생사업비 ↔ 계정정보 ↔ Cost Center Master ↔ 출력> 시트
- **컬럼·시트명 유연 대응** — `column_config.yaml`로 외부화, 불일치 감지 시 GUI 검증 창 제공
- **PDF 테마 설정** — 포인트 색상과 로고(최대 3개)를 실행 시 선택 가능
- **Fuzzy 매칭** — 컬럼명 버전 변경(예: `대상정의 v3.0_0415`) 시 자동 추천

---

## 요구 환경

| 항목 | 최소 버전 |
|------|-----------|
| Python | 3.10 이상 (권장 3.11+) |
| OS | Windows (맑은 고딕 폰트 필요) |
| 폰트 | `C:\Windows\Fonts\malgun.ttf` |
| tkinter | Python 설치 시 기본 포함 |

---

## 설치

```bash
pip install -r requirements.txt
```

---

## 실행 (개발 환경)

```bash
cd report_generator/src/code
python main.py
```

실행 흐름:
1. 포인트 색상 변경 여부 선택
2. 입력 Excel 파일 선택
3. 컬럼·시트명 검증 (불일치 시 GUI 수정 창)
4. 결과 저장 폴더 선택
5. 코드별 PDF 자동 생성 → 완료 요약

---

## EXE 빌드

```bat
build.bat
```

빌드 완료 후:
- `dist/전표분석서_자동화.exe` — 실행 파일
- `전표분석서_자동화_배포용.zip` — EXE + 설정 파일 2개를 묶은 배포용 패키지

---

## 파일 구조

```
python_data_cleaning/
├── report_generator/src/code/
│   ├── main.py              # 진입점 — tkinter UI, 전체 흐름 제어
│   ├── config_manager.py    # column_config.yaml 로드/저장/검증
│   ├── data_loader.py       # 4개 시트 로드 + Phase 0 클렌징
│   ├── processor.py         # 파이프라인 — 필터링/3중 JOIN/집계
│   ├── pdf_exporter.py      # ReportLab A4 가로 PDF 생성
│   ├── csv_exporter.py      # CSV 출력 (template.csv 기반)
│   ├── company_theme.py     # 색상 테마 데이터클래스
│   ├── theme_manager.py     # theme_config.json 저장·로드
│   ├── template.csv         # CSV 출력 레이아웃 템플릿
│   └── templates/
│       └── report_template.html  # 레이아웃 참조용 HTML
├── images/                  # 기본 로고 이미지 (ABL.png, 우리금융그룹.png)
├── manual/                  # 사용자 매뉴얼
├── build.bat                # 빌드 실행 (더블클릭)
├── build.ps1                # 빌드 자동화 스크립트 (PyInstaller)
├── column_config.yaml       # 컬럼·시트명 설정 (Excel 구조 변경 시 수정)
├── theme_config.json        # 테마 설정 (프로그램이 자동 저장)
└── requirements.txt         # Python 의존성
```

---

## 설정 파일

### `column_config.yaml`

Excel 파일의 시트명·컬럼명이 변경된 경우 이 파일을 수정합니다.

```yaml
sheets:
  transaction: "실제발생사업비"   # 시트 탭 이름과 정확히 일치해야 함
  account: "계정정보"
  ccm: "Cost Center Master"
  output: "출력>"

columns:
  대상정의: "대상정의 v3.0_0415"  # 버전 번호 변경 시 여기를 수정
  cc_name: "Cost Center name"     # 대소문자 주의
  # ... (나머지 컬럼)
```

- 프로그램 실행 중 불일치가 감지되면 GUI 검증 창에서 수정 후 저장할 수 있습니다.
- 파일이 없으면 자동 생성됩니다.

### `theme_config.json`

PDF 포인트 색상과 로고 경로가 저장됩니다. 프로그램 실행 시 자동 저장되므로 직접 수정하지 않아도 됩니다.

---

## 데이터 파이프라인 개요

```
Excel 입력
    │
    ▼
[data_loader] 4개 시트 로드 + Phase 0 클렌징
    │  - 형변환 (JOIN 키 str 통일)
    │  - 공백 제거
    │  - 숫자 결측치 → 0
    │
    ▼
[processor] 코드별 파이프라인
    │  Step 1. 원가요소 기준 필터링
    │  Step 2. 3중 LEFT JOIN + 대상금액 합산
    │  Step 3. 주관/사용부서 귀속 현황
    │  Step 4. 직접비/공통비 분류
    │  Step 5. 성격별 × 부점 교차 집계
    │  Step 6. 계정 헤더 정보 추출
    │
    ▼
[pdf_exporter] A4 가로 PDF 생성 (ReportLab)
```
