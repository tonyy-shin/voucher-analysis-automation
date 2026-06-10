"""
data_loader.py
역할: 4개 시트 로드 + Phase 0 클렌징(형변환·공백제거·결측치 처리).
      클렌징이 완료된 DataFrame만 processor.py 파이프라인에 전달된다.

7대 원칙 중 원칙 4·5·6 담당 모듈.
"""
from __future__ import annotations

import difflib
import re

import pandas as pd

# ── 시트명 상수 ──────────────────────────────────────────────────────────────
_SHEET_TRANSACTION = "실제발생사업비"
_SHEET_ACCOUNT     = "계정정보"
_SHEET_CCM         = "Cost Center Master"
_SHEET_OUTPUT      = "출력>"

# 헤더 탐색 범위 — 모든 _find_*_header_row 함수에서 공통 사용
_HEADER_SCAN_ROWS = 20

# JOIN 미매칭 시 채워지는 표시 텍스트
FALLBACK_TEXT = "(미매칭)"

# 파이프라인 전반에서 사용되는 파생 컬럼명 상수
COL_SUBTOTAL = "합계"   # 행별 성격 합계 컬럼
COL_TOTAL    = "총계"   # 집계 마지막 행의 부서명 값
COL_TEAM     = "팀"     # 실제발생사업비 시트의 팀 컬럼

# ── 컬럼명 중앙 관리 (논리명 → 실제 Excel 컬럼명) ───────────────────────────
# 여기서 수정하면 전체 파이프라인에 반영됨
COLUMN_MAP: dict[str, str] = {
    # 실제발생사업비 컬럼
    "원가요소":           "원가요소",
    "코스트센터":         "코스트 센터",
    # 계정정보 컬럼
    "계정번호":           "계정번호",
    "계정명":             "계정명",
    "계정그룹ID":         "계정그룹ID",
    "계정그룹명":         "계정그룹명",
    "직간접구분_계정":    "직/간접비",
    "대상정의":           "대상정의 v3.0_0415",
    "범위":               "사용 부서",
    "지급대상":           "비용 지급 범위",
    "산출기준":           "산출기준",
    # Cost Center Master 컬럼
    "cc_code":            "Cost Center Code",
    "cc_name":            "Cost Center name",
    "직간접구분":         "변경: 직접/공통 구분",
    # 출력> 시트 컬럼
    "출력전표":           "출력전표",
    "귀속_주관부서":      "귀속_주관부서",
    "귀속_사용부서":      "귀속_사용부서",
}


# ── Fuzzy 컬럼 탐색 패턴 (버전번호·공백 변동이 잦은 컬럼) ─────────────────────
# resolve_col()에서 정규식 탐색에 사용. 키는 COLUMN_MAP 키와 일치해야 함.
_COLUMN_PATTERNS: dict[str, str] = {
    "대상정의":   r"대상정의",
    "직간접구분": r"직간접\s*구분",
    "cc_code":   r"[Cc]ost\s*[Cc]enter\s*[Cc]ode",
    "cc_name":   r"[Cc]ost\s*[Cc]enter\s*[Nn]ame",
    "범위":      r"사용\s*부서",
    "지급대상":  r"비용\s*지급\s*범위",
}
# ── Cost Center Master 리네임 맵 ─────────────────────────────────────────────
# read_excel 시 중복 열 이름(Code, Code.1)을 의미 있는 이름으로 교정
_CCM_RENAME_MAP: dict[str, str] = {
    "Code":   "본부_Code",
    "Code.1": "팀_Code",
}

# ── 각 시트에서 실제로 필요한 컬럼 목록 ──────────────────────────────────────
def _get_ccm_cols_needed() -> list[str]:
    return [COLUMN_MAP["cc_code"], COLUMN_MAP["직간접구분"]]

def _get_account_cols_needed() -> list[str]:
    return [
        COLUMN_MAP["계정번호"],
        COLUMN_MAP["계정명"],
        COLUMN_MAP["계정그룹ID"],
        COLUMN_MAP["계정그룹명"],
        COLUMN_MAP["대상정의"],
        COLUMN_MAP["범위"],
        COLUMN_MAP["지급대상"],
        COLUMN_MAP["산출기준"],
        COLUMN_MAP["직간접구분_계정"],
    ]

# ── 시트별 필수 컬럼: 없으면 JOIN이 전량 NaN으로 오염되므로 즉시 에러 ────────────
# (AMOUNT_COLS·NATURE_COLS 등 선택적 컬럼은 포함하지 않음)
_REQUIRED_TRANSACTION_COLS: list[str] = [
    COLUMN_MAP["원가요소"],    # "원가요소"   — 필터 키 + JOIN 2 키
    COLUMN_MAP["코스트센터"],  # "코스트 센터" — JOIN 1 키
]

_REQUIRED_CCM_COLS: list[str] = [
    COLUMN_MAP["cc_code"],     # "Cost Center Code"     — JOIN 1 키
    COLUMN_MAP["직간접구분"],  # "변경: 직접/공통 구분"    — 직간접 분류 기준
]

_REQUIRED_ACCOUNT_COLS: list[str] = [
    COLUMN_MAP["계정번호"],    # "계정번호" — JOIN 2 키
]


# ── 금액 컬럼 (fillna(0) 대상) ───────────────────────────────────────────────
AMOUNT_COLS: list[str] = [
    "Sum of DA_P", "Sum of DA_N", "Sum of DM", "Sum of DC",
    "Sum of DI",  "Sum of DP",   "Sum of IM",  "Sum of II",
    "Sum of IO",  "Sum of IA",
]

# ── 성격별 분류 컬럼 (fillna(0) 대상) ────────────────────────────────────────
NATURE_COLS: list[str] = [
    "계약비", "유지비", "손해조사비", "투자관리비", "기타사업비",
]


# ════════════════════════════════════════════════════════════════════════════
#  Public API
# ════════════════════════════════════════════════════════════════════════════


def check_unregistered_teams(wb_path: str, config: dict | None = None) -> list[str]:
    """
    실제발생사업비의 팀 목록(Set A)과 출력> 시트 등록 부서(Set B)를 비교하여
    Set A - Set B (미등록 팀)를 반환한다.

    JOIN 없이 두 시트의 최소 컬럼만 읽어 초기 단계에서 경량으로 실행된다.
    load_all_sheets() 이전에 호출되어 PDF 생성 전 데이터 정합성을 사전 검증한다.

    Args:
        wb_path: Excel 파일 경로
        config:  column_config.yaml 설정 (None이면 기본값 사용)

    Returns:
        미등록 팀 이름 목록 (정렬됨). 없으면 [].
    """
    sheet_transaction = _SHEET_TRANSACTION
    sheet_output      = _SHEET_OUTPUT
    col_team          = COL_TEAM
    label_코드        = "출력전표"
    label_주관        = "주관부서"
    label_사용        = "사용부서"

    if config:
        cfg_sheets        = config.get("sheets", {})
        sheet_transaction = cfg_sheets.get("transaction", sheet_transaction)
        sheet_output      = cfg_sheets.get("output",      sheet_output)
        col_team          = config.get("col_team", col_team)
        out_labels        = config.get("output_labels", {})
        label_코드 = out_labels.get("코드", label_코드)
        label_주관 = out_labels.get("주관", label_주관)
        label_사용 = out_labels.get("사용", label_사용)

    # ── Set A: 실제발생사업비 팀 컬럼만 경량 로드 ────────────────────────────
    try:
        header_row = _find_actual_header_row(wb_path, sheet_transaction)
        df_teams = pd.read_excel(
            wb_path,
            sheet_name=sheet_transaction,
            header=header_row,
            usecols=[col_team],
        )
        set_a: set[str] = {
            str(v).strip()
            for v in df_teams[col_team].dropna().unique()
            if str(v).strip() not in ("", "nan", FALLBACK_TEXT)
        }
    except Exception:
        return []

    if not set_a:
        return []

    # ── Set B: 출력> 시트 등록 부서 합집합 ───────────────────────────────────
    try:
        with pd.ExcelFile(wb_path) as xl:
            if sheet_output not in xl.sheet_names:
                return []
            df_raw = pd.read_excel(xl, sheet_name=sheet_output, header=None)
    except Exception:
        return []

    df_out = _preprocess_output(df_raw, label_코드, label_주관, label_사용)
    if df_out.empty:
        return []

    set_b: set[str] = set()
    for col in [COLUMN_MAP["귀속_주관부서"], COLUMN_MAP["귀속_사용부서"]]:
        if col in df_out.columns:
            vals = df_out[col].dropna().astype(str).str.strip()
            set_b.update(vals[~vals.isin({"", "nan", FALLBACK_TEXT})].tolist())

    # ── 차집합 반환 ──────────────────────────────────────────────────────────
    return sorted(set_a - set_b)

def load_all_sheets(wb_path: str, config: dict | None = None) -> dict[str, pd.DataFrame]:
    """
    4개 시트를 로드하고 Phase 0 클렌징까지 완료한 DataFrame dict를 반환한다.

    반환 키:
        "실제발생사업비", "계정정보", "Cost Center Master", "출력>"

    Args:
        wb_path: 원본 데이터가 담긴 Excel 파일 경로 (.xlsx)

    Returns:
        클렌징이 완료된 DataFrame을 담은 dict
    """
    global AMOUNT_COLS, NATURE_COLS, COL_TEAM
    global _REQUIRED_TRANSACTION_COLS, _REQUIRED_CCM_COLS, _REQUIRED_ACCOUNT_COLS

    # ── config 오버라이드 적용 ──────────────────────────────────────────────
    sheet_transaction = _SHEET_TRANSACTION
    sheet_account     = _SHEET_ACCOUNT
    sheet_ccm         = _SHEET_CCM
    sheet_output      = _SHEET_OUTPUT

    # 출력> 시트 레이블 기본값 (config의 output_labels로 오버라이드 가능)
    label_코드 = "출력전표"
    label_주관 = "주관부서"
    label_사용 = "사용부서"

    if config:
        sheets_cfg = config.get("sheets", {})
        sheet_transaction = sheets_cfg.get("transaction", sheet_transaction)
        sheet_account     = sheets_cfg.get("account",     sheet_account)
        sheet_ccm         = sheets_cfg.get("ccm",         sheet_ccm)
        sheet_output      = sheets_cfg.get("output",      sheet_output)
        for k, v in config.get("columns", {}).items():
            if k in COLUMN_MAP:
                COLUMN_MAP[k] = v

        # ── 선택적 컬럼 목록 오버라이드 ────────────────────────────────────
        if "amount_cols" in config and isinstance(config["amount_cols"], list):
            AMOUNT_COLS = list(config["amount_cols"])
        if "nature_cols" in config and isinstance(config["nature_cols"], list):
            NATURE_COLS = list(config["nature_cols"])
        if "col_team" in config:
            COL_TEAM = str(config["col_team"])

        # ── 출력> 시트 레이블 오버라이드 ───────────────────────────────────
        out_labels = config.get("output_labels", {})
        label_코드 = out_labels.get("코드", label_코드)
        label_주관 = out_labels.get("주관", label_주관)
        label_사용 = out_labels.get("사용", label_사용)

    # COLUMN_MAP 업데이트 후 재구성 — 모듈 임포트 시 캡처된 값을 현재 COLUMN_MAP 기준으로 갱신
    _REQUIRED_TRANSACTION_COLS = [COLUMN_MAP["원가요소"],  COLUMN_MAP["코스트센터"]]
    _REQUIRED_CCM_COLS         = [COLUMN_MAP["cc_code"],   COLUMN_MAP["직간접구분"]]
    _REQUIRED_ACCOUNT_COLS     = [COLUMN_MAP["계정번호"]]

    ccm_header_row     = _find_ccm_header_row(wb_path, sheet_ccm)
    actual_header_row  = _find_actual_header_row(wb_path, sheet_transaction)
    account_header_row = _find_account_header_row(wb_path, sheet_account)

    with pd.ExcelFile(wb_path) as xl:
        source_sheet = _find_source_sheet(xl, sheet_transaction)
        df_actual  = pd.read_excel(xl, sheet_name=source_sheet,    header=actual_header_row)
        df_account = pd.read_excel(xl, sheet_name=sheet_account,   header=account_header_row)
        df_ccm     = pd.read_excel(xl, sheet_name=sheet_ccm,       header=ccm_header_row)
        df_output  = (
            pd.read_excel(xl, sheet_name=sheet_output, header=None)
            if sheet_output in xl.sheet_names
            else pd.DataFrame()
        )

    return {
        _SHEET_TRANSACTION: _preprocess_actual(df_actual),
        _SHEET_ACCOUNT:     _preprocess_account(df_account),
        _SHEET_CCM:         _preprocess_ccm(df_ccm),
        _SHEET_OUTPUT:      _preprocess_output(df_output, label_코드, label_주관, label_사용),
    }


# ════════════════════════════════════════════════════════════════════════════
#  Phase 0 — 클렌징 핵심 함수
# ════════════════════════════════════════════════════════════════════════════

def cleanse_types(df: pd.DataFrame, join_key_cols: list[str]) -> pd.DataFrame:
    """
    JOIN 키 컬럼을 str 타입으로 강제 형변환한다.

    Excel에서 숫자로 저장된 코드(예: 1234)가 한쪽은 int, 다른 쪽은 str로 로드되면
    JOIN 결과가 전량 NaN으로 처리되는 조용한 버그가 발생한다.
    형변환과 공백 제거를 동시에 적용하고, 변환 후 생성된 "nan" 문자열을 제거한다.

    Args:
        df: 대상 DataFrame
        join_key_cols: str로 변환할 JOIN 키 컬럼 이름 목록

    Returns:
        형변환이 적용된 DataFrame (in-place 수정 후 반환)
    """
    for col in join_key_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.strip()
                # float 경유 시 "1000.0" → "1000" 교정
                .str.replace(r"\.0$", "", regex=True)
                # 변환 후 nan 문자열 제거
                .replace({"nan": "", "None": "", "NaN": ""})
            )
    return df


def cleanse_whitespace(df: pd.DataFrame) -> pd.DataFrame:
    """
    object dtype 컬럼 전체에 .str.strip()을 적용한다.

    Excel 입력자가 실수로 삽입한 앞뒤 공백("직접비 ", " 계약비")이
    str.contains() 필터링 및 groupby 집계에서 해당 값을 누락시키거나
    별도 그룹으로 집계하는 오류를 방지한다.

    Args:
        df: 대상 DataFrame

    Returns:
        공백 제거가 적용된 DataFrame
    """
    str_cols = df.select_dtypes(include="object").columns
    for col in str_cols:
        df[col] = df[col].str.strip()
    return df


def cleanse_nulls(df: pd.DataFrame, numeric_cols: list[str]) -> pd.DataFrame:
    """
    숫자 컬럼의 NaN을 fillna(0)으로 처리한다.

    Excel 빈 셀은 NaN(float)으로 로드되며, sum()/groupby() 집계 시
    해당 행이 제외되거나 NaN이 전파되어 집계 오류를 유발한다.

    Args:
        df: 대상 DataFrame
        numeric_cols: fillna(0)을 적용할 숫자 컬럼 이름 목록

    Returns:
        결측치가 처리된 DataFrame
    """
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def _validate_columns(
    df: pd.DataFrame,
    required_cols: list[str],
    sheet_name: str,
) -> None:
    """
    필수 컬럼이 DataFrame에 모두 존재하는지 검증한다.

    JOIN 키 컬럼이 없으면 merge() 결과가 전량 NaN이 되어 '(미매칭)'으로
    조용히 오염되는 문제를 사전에 차단한다.

    Args:
        df:            검증 대상 DataFrame
        required_cols: 반드시 존재해야 할 컬럼명 목록
        sheet_name:    에러 메시지에 표시할 시트명

    Raises:
        ValueError: 필수 컬럼 중 하나라도 누락된 경우
    """
    missing = [c for c in required_cols if c not in df.columns]
    if not missing:
        return
    missing_lines = "\n".join(f"  - {c}" for c in missing)
    actual_lines  = "\n".join(f"  - {c}" for c in df.columns.tolist())
    raise ValueError(
        f"[{sheet_name}] 시트에 필수 컬럼이 없습니다.\n\n"
        f"누락된 컬럼:\n{missing_lines}\n\n"
        f"실제 컬럼 목록:\n{actual_lines}\n\n"
        "Excel 파일의 컬럼명과 data_loader.COLUMN_MAP 정의를 비교하여 수정하세요."
    )


# ════════════════════════════════════════════════════════════════════════════
#  시트별 전처리 (로드 직후 Phase 0 클렌징 호출)
# ════════════════════════════════════════════════════════════════════════════

def _normalize_amount_col_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    AMOUNT_COLS + NATURE_COLS의 컬럼명을 대소문자 무관하게 매칭하여 코드 기준명으로 rename.

    실제 Excel 파일이 'sum of da_p'(소문자) 등으로 저장된 경우에도
    'Sum of DA_P'(코드 기준명)로 통일하여 이후 집계가 정상 동작하도록 한다.
    """
    target_cols = AMOUNT_COLS + NATURE_COLS
    lower_to_canonical = {c.lower(): c for c in target_cols}
    rename_map: dict[str, str] = {}
    for col in df.columns:
        canonical = lower_to_canonical.get(col.lower())
        if canonical and col != canonical:
            rename_map[col] = canonical
    return df.rename(columns=rename_map) if rename_map else df


def _preprocess_actual(df: pd.DataFrame) -> pd.DataFrame:
    """
    실제발생사업비 시트 전처리.

    적용 순서:
        1. Unnamed 컬럼 제거
        2. AMOUNT_COLS / NATURE_COLS 컬럼명 대소문자 정규화
        3. cleanse_types  — JOIN 키 (코스트 센터, 원가요소)
        4. cleanse_whitespace — 텍스트 컬럼 전체
        5. cleanse_nulls  — 금액·성격별 분류 컬럼
    """
    # 1. Unnamed 컬럼 제거
    unnamed = [c for c in df.columns if "Unnamed" in str(c)]
    if unnamed:
        df = df.drop(columns=unnamed)

    # 필수 컬럼 검증 — JOIN 키 누락 시 즉시 ValueError
    _validate_columns(df, _REQUIRED_TRANSACTION_COLS, _SHEET_TRANSACTION)

    # 2. 금액 컬럼명 대소문자 정규화 (예: 'sum of da_p' → 'Sum of DA_P')
    df = _normalize_amount_col_names(df)

    # JOIN 키 형변환
    df = cleanse_types(df, [COLUMN_MAP["코스트센터"], COLUMN_MAP["원가요소"]])

    # 텍스트 공백 제거
    df = cleanse_whitespace(df)

    # 숫자 결측치 처리
    _numeric_cols = AMOUNT_COLS + NATURE_COLS
    if "합계" in df.columns:
        _numeric_cols = _numeric_cols + ["합계"]
    df = cleanse_nulls(df, _numeric_cols)

    return df


def _preprocess_ccm(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cost Center Master 시트 전처리.

    적용 순서:
        1. 중복 열 리네임 (Code → 본부_Code, Code.1 → 팀_Code)
        2. cleanse_types  — JOIN 키 (Cost Center Code)
        3. cleanse_whitespace
        4. 필요 컬럼만 선택 및 키 결측 행 제거
    """
    df = _rename_duplicate_code_cols(df)

    # 필수 컬럼 검증 — JOIN 키 누락 시 즉시 ValueError
    _validate_columns(df, _REQUIRED_CCM_COLS, _SHEET_CCM)

    # JOIN 키 형변환
    df = cleanse_types(df, [COLUMN_MAP["cc_code"]])

    # 텍스트 공백 제거
    df = cleanse_whitespace(df)

    # 필요 컬럼만 선택
    available = [c for c in _get_ccm_cols_needed() if c in df.columns]
    df = df[available].copy()

    # 키가 없는 행 및 중복 키 행 제거 (JOIN fan-out 방지)
    _cc = COLUMN_MAP["cc_code"]
    df = df.dropna(subset=[_cc])
    df = df[df[_cc] != ""]
    df = df.drop_duplicates(subset=[_cc])

    return df


def _preprocess_account(df: pd.DataFrame) -> pd.DataFrame:
    """
    계정정보 시트 전처리.

    적용 순서:
        1. cleanse_types  — JOIN 키 (계정번호)
        2. cleanse_whitespace
        3. 필요 컬럼만 선택
    """
    # 필수 컬럼 검증 — JOIN 키 누락 시 즉시 ValueError
    _validate_columns(df, _REQUIRED_ACCOUNT_COLS, _SHEET_ACCOUNT)

    # JOIN 키 형변환
    df = cleanse_types(df, [COLUMN_MAP["계정번호"]])

    # 텍스트 공백 제거
    df = cleanse_whitespace(df)

    available = [c for c in _get_account_cols_needed() if c in df.columns]
    df = df[available].copy()

    # 중복 키 행 제거 (JOIN fan-out 방지)
    _acct = COLUMN_MAP["계정번호"]
    if _acct in df.columns:
        df = df.drop_duplicates(subset=[_acct])

    return df


_CODE_RESERVED: frozenset[str] = frozenset({"", "nan", "None", "NaN"})


def _parse_output_code(val: object, extra_reserved: frozenset[str] = frozenset()) -> str | None:
    """셀 값을 출력전표 코드 문자열로 변환한다.

    변환 우선순위:
        1. int(float(val)) 성공 → 정수 문자열 반환 (예: "5001.0" → "5001")
        2. 실패 → 영문·혼합 코드로 간주하고 공백 제거 후 그대로 반환 (예: "A001")
    None을 반환하는 경우:
        - val이 NaN/None
        - 변환 결과가 _CODE_RESERVED 또는 extra_reserved에 속하는 경우
        - 공백 제거 후 빈 문자열
    """
    import pandas as _pd  # noqa: PLC0415
    if _pd.isna(val):
        return None
    raw = str(val).strip()
    if raw in _CODE_RESERVED or raw in extra_reserved:
        return None
    try:
        return str(int(float(raw)))
    except (ValueError, OverflowError):
        cleaned = re.sub(r"\s+", "", raw)
        return raw if cleaned else None


def _preprocess_output(
    df_raw: pd.DataFrame,
    label_코드: str = "출력전표",
    label_주관: str = "주관부서",
    label_사용: str = "사용부서",
) -> pd.DataFrame:
    """
    출력> 시트 전처리 — 출력전표·주관부서·사용부서 열을 독립적으로 수집한다.

    각 열은 자신의 헤더 행 다음 행부터 독립적으로 수집되므로
    열 간 행 위치는 무관하다. ffill 없음.
    반환: [출력전표, 귀속_주관부서, 귀속_사용부서] (짧은 열은 NaN 패딩)
    """
    col_코드_out = COLUMN_MAP["출력전표"]
    col_주관_out = COLUMN_MAP["귀속_주관부서"]
    col_사용_out = COLUMN_MAP["귀속_사용부서"]
    _EMPTY = pd.DataFrame(columns=[col_코드_out, col_주관_out, col_사용_out])

    if df_raw.empty:
        return _EMPTY

    # ── 1. 각 헤더의 열 위치와 헤더 행 번호를 독립적으로 탐색 ────────────────────
    c_코드: int | None = None
    c_주관: int | None = None
    c_사용: int | None = None
    hr_코드: int | None = None
    hr_주관: int | None = None
    hr_사용: int | None = None

    for row_i in range(min(_HEADER_SCAN_ROWS, len(df_raw))):
        for col_i, val in enumerate(df_raw.iloc[row_i]):
            v = str(val).strip()
            if v == label_코드 and c_코드 is None:
                c_코드 = col_i
                hr_코드 = row_i
            if v == label_주관 and c_주관 is None:
                c_주관 = col_i
                hr_주관 = row_i
            if v == label_사용 and c_사용 is None:
                c_사용 = col_i
                hr_사용 = row_i

    # 주관부서 열은 필수
    if c_주관 is None or hr_주관 is None:
        return _EMPTY

    # ── 2. 각 열을 독립적으로 수집 (헤더 다음 행부터) ────────────────────────────
    _EXCLUDED = {"", "nan", label_코드, label_주관, label_사용}

    # 출력전표
    if c_코드 is not None and hr_코드 is not None:
        코드_vals = (
            df_raw.iloc[hr_코드 + 1:, c_코드]
            .apply(_parse_output_code)
            .dropna()
            .reset_index(drop=True)
        )
    else:
        코드_vals = pd.Series(dtype=object)

    # 주관부서
    주관_raw = df_raw.iloc[hr_주관 + 1:, c_주관].astype(str).str.strip()
    주관_vals = 주관_raw[~주관_raw.isin(_EXCLUDED)].dropna().reset_index(drop=True)

    # 사용부서
    if c_사용 is not None and hr_사용 is not None:
        사용_raw = df_raw.iloc[hr_사용 + 1:, c_사용].astype(str).str.strip()
        사용_vals = 사용_raw[~사용_raw.isin(_EXCLUDED)].dropna().reset_index(drop=True)
    else:
        사용_vals = pd.Series(dtype=str)

    # ── 3. 가장 긴 열 기준 DataFrame (짧은 열은 NaN 패딩) ─────────────────────────
    result = pd.DataFrame({
        col_코드_out: 코드_vals,
        col_주관_out: 주관_vals,
        col_사용_out: 사용_vals,
    })

    return result


# ════════════════════════════════════════════════════════════════════════════
#  헬퍼 — 시트 탐지
# ════════════════════════════════════════════════════════════════════════════

def _find_source_sheet(xl: pd.ExcelFile, target: str = _SHEET_TRANSACTION) -> str:
    """
    원천 트랜잭션 시트를 자동 탐지한다.

    정확한 시트명이 없을 경우 컬럼명으로 폴백 탐색하여
    파일 버전 변경에 대응한다.

    Args:
        xl: pd.ExcelFile 객체

    Returns:
        탐지된 시트명

    Raises:
        ValueError: 트랜잭션 시트를 찾을 수 없을 때
    """
    if target in xl.sheet_names:
        return target
    for name in xl.sheet_names:
        try:
            df_peek = pd.read_excel(xl, sheet_name=name, nrows=1)
            if COLUMN_MAP["원가요소"] in df_peek.columns and COLUMN_MAP["코스트센터"] in df_peek.columns:
                return name
        except Exception:
            continue
    raise ValueError(
        f"원천 트랜잭션 시트를 찾을 수 없습니다.\n"
        f"시트 목록: {xl.sheet_names}"
    )


def _find_actual_header_row(wb_path: str, sheet_name: str = _SHEET_TRANSACTION) -> int:
    """
    실제발생사업비 시트의 헤더 행 번호를 동적으로 탐지한다.

    '원가요소' 또는 '코스트 센터'가 포함된 행을 헤더로 간주한다.
    상단에 제목·메타 행이 추가된 경우에도 정상 파싱되도록 한다.

    Returns:
        헤더 행 번호 (0-based). 탐지 실패 시 0 반환 (기존 동작 유지).
    """
    try:
        df_raw = pd.read_excel(wb_path, sheet_name=sheet_name, header=None, nrows=_HEADER_SCAN_ROWS)
        for i, row in df_raw.iterrows():
            vals = [str(v).strip() for v in row.values if pd.notna(v)]
            if COLUMN_MAP["원가요소"] in vals or COLUMN_MAP["코스트센터"] in vals:
                return int(i)
    except Exception:
        pass
    return 0


def _find_account_header_row(wb_path: str, sheet_name: str = _SHEET_ACCOUNT) -> int:
    """
    계정정보 시트의 헤더 행 번호를 동적으로 탐지한다.

    '계정번호'가 포함된 행을 헤더로 간주한다.

    Returns:
        헤더 행 번호 (0-based). 탐지 실패 시 0 반환 (기존 동작 유지).
    """
    try:
        df_raw = pd.read_excel(wb_path, sheet_name=sheet_name, header=None, nrows=_HEADER_SCAN_ROWS)
        for i, row in df_raw.iterrows():
            vals = [str(v).strip() for v in row.values if pd.notna(v)]
            if COLUMN_MAP["계정번호"] in vals:
                return int(i)
    except Exception:
        pass
    return 0


def _find_ccm_header_row(wb_path: str, sheet_name: str = _SHEET_CCM) -> int:
    """
    Cost Center Master 시트의 헤더 행 번호를 동적으로 탐지한다.

    파일 버전에 따라 상단에 메타 행이 추가될 수 있으므로
    'Cost Center Code'와 'Cost Center name'이 동시에 존재하는
    행을 스캔하여 실제 헤더 행 번호를 반환한다.

    Args:
        wb_path: Excel 파일 경로

    Returns:
        헤더로 사용할 행 번호 (0-based, pd.read_excel의 header= 인수에 직접 사용)

    Raises:
        ValueError: 헤더 행을 찾을 수 없을 때
    """
    df_raw = pd.read_excel(wb_path, sheet_name=sheet_name, header=None, nrows=_HEADER_SCAN_ROWS)
    for i, row in df_raw.iterrows():
        values = [str(v).strip() for v in row.values if pd.notna(v)]
        if COLUMN_MAP["cc_code"] in values and COLUMN_MAP["cc_name"] in values:
            return int(i)
    raise ValueError(
        "Cost Center Master 헤더 행을 찾을 수 없습니다.\n"
        f"'{COLUMN_MAP['cc_code']}'와 '{COLUMN_MAP['cc_name']}' 컬럼이 파일에 존재하는지 확인하세요."
    )


# ════════════════════════════════════════════════════════════════════════════
#  Fuzzy Resolver — config 기반 컬럼·시트명 유연 탐색 (Phase 2 신규)
# ════════════════════════════════════════════════════════════════════════════

def resolve_col(df: pd.DataFrame, logical_name: str) -> str | None:
    """
    DataFrame에서 논리명에 해당하는 실제 컬럼명을 탐색한다.

    탐색 우선순위:
        1. COLUMN_MAP[logical_name] exact match (기존 동작 유지)
        2. _COLUMN_PATTERNS[logical_name] regex scan over df.columns
        3. difflib.get_close_matches(COLUMN_MAP[logical_name], df.columns, cutoff=0.8)
        4. None 반환

    Args:
        df:           탐색 대상 DataFrame
        logical_name: COLUMN_MAP 키 (예: "대상정의", "cc_code")

    Returns:
        찾은 실제 컬럼명. 모든 단계 실패 시 None.
    """
    expected = COLUMN_MAP.get(logical_name)
    if expected is None:
        return None

    # 1. Exact match
    if expected in df.columns:
        return expected

    # 2. Regex scan
    pattern = _COLUMN_PATTERNS.get(logical_name)
    if pattern:
        for col in df.columns:
            if re.search(pattern, col):
                return col

    # 3. Fuzzy match
    candidates = list(df.columns)
    matches = difflib.get_close_matches(expected, candidates, n=1, cutoff=0.8)
    if matches:
        return matches[0]

    return None


def _resolve_sheet_name(sheet_names: list[str], target: str) -> str | None:
    """
    Excel 시트 목록에서 target과 가장 유사한 시트명을 반환한다.

    탐색 우선순위:
        1. Exact match
        2. difflib.get_close_matches(target, sheet_names, cutoff=0.7)
        3. None 반환

    Args:
        sheet_names: pd.ExcelFile.sheet_names
        target:      탐색할 시트명

    Returns:
        찾은 시트명. 실패 시 None.
    """
    if target in sheet_names:
        return target
    matches = difflib.get_close_matches(target, sheet_names, n=1, cutoff=0.7)
    return matches[0] if matches else None


def _rename_duplicate_code_cols(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cost Center Master에서 pandas가 자동 생성한 'Code', 'Code.1', 'Code.2'...를
    위치 순서대로 의미 있는 이름('본부_Code', '팀_Code')으로 교체한다.

    기존 _CCM_RENAME_MAP 방식은 pandas 버전 따라 자동 접미사(Code.1→Code.2 등)가
    달라지면 조용히 실패하므로, 정규식으로 패턴을 탐지하여 위치 기반으로 처리한다.
    """
    labels = ["본부_Code", "팀_Code"]
    counter = 0
    rename: dict[str, str] = {}
    for col in df.columns:
        if re.fullmatch(r"Code(\.\d+)?", col):
            rename[col] = labels[counter] if counter < len(labels) else f"기타_Code_{counter}"
            counter += 1
    return df.rename(columns=rename) if rename else df
