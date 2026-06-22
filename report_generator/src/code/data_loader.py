"""
data_loader.py
역할: 4개 CSV 파일 로드 + Phase 0 클렌징(형변환·공백제거·결측치 처리).
      클렌징이 완료된 DataFrame만 processor.py 파이프라인에 전달된다.

입력 데이터는 단일 Excel(다중 시트)이 아니라 4개의 평면 CSV 파일이다:
    사업비정보.csv / 계정정보.csv / 부서정보.csv / 출력.csv
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

# ── 논리 키 상수 ─────────────────────────────────────────────────────────────
# load_all_csvs() 반환 dict의 키이자, processor·main에서 시트(데이터셋) 식별에 사용.
# (Excel 시트 시절 명칭을 유지하여 import 호환을 보존)
_SHEET_TRANSACTION = "사업비정보"
_SHEET_ACCOUNT     = "계정정보"
_SHEET_CCM         = "부서정보"
_SHEET_OUTPUT      = "출력"

# JOIN 미매칭 시 채워지는 표시 텍스트
FALLBACK_TEXT = "(미매칭)"

# 파이프라인 전반에서 사용되는 파생 컬럼명 상수
COL_TOTAL_AMOUNT = "합계"   # 사업비정보.csv의 행별 합계 컬럼 = 대상금액 소스
COL_SUBTOTAL     = "소계"   # nature 집계 결과 df의 부서별 소계 컬럼 (계산 결과, CSV에 없음)
COL_TOTAL        = "총계"   # 집계 마지막 행의 부서명 값
COL_TEAM         = "팀"     # 부서정보.csv의 팀 컬럼 (JOIN으로 트랜잭션에 결합)

# ── 컬럼명 중앙 관리 (논리명 → 실제 CSV 컬럼명) ──────────────────────────────
# 여기서 수정하면 전체 파이프라인에 반영됨
COLUMN_MAP: dict[str, str] = {
    # 사업비정보.csv 컬럼
    "원가요소":           "원가요소",
    "코스트센터":         "코스트센터",
    # 계정정보.csv 컬럼
    "계정번호":           "계정번호",
    "계정명":             "계정명",
    "계정그룹ID":         "계정그룹ID",
    "계정그룹명":         "계정그룹명",
    "대상정의":           "대상정의",
    "범위":               "사용 부서",
    "지급대상":           "비용 지급 범위",
    "산출기준":           "산출기준",
    # 부서정보.csv 컬럼
    "cc_code":            "Cost Center Code",
    "직간접구분":         "직간접구분",
    "팀":                 "팀",
    # 출력.csv 컬럼
    "출력전표":           "출력전표",
    "귀속_주관부서":      "주관부서",
    "귀속_사용부서":      "사용부서",
}

# ── 성격별 분류 컬럼 (사업비정보.csv의 실제 숫자 값; fillna(0) 대상) ──────────
NATURE_COLS: list[str] = [
    "계약비", "유지비", "손해조사비", "투자관리비", "간접비", "공통비",
]

# ── 출력.csv의 분류 근거 텍스트 컬럼 (출력전표 코드별로 행이 다름) ───────────
# 각 컬럼 → processor.build_classification_basis()가 "{컬럼}_근거" 키로 변환
BASIS_TEXT_COLS: list[str] = [
    "직접비", "간접비", "공통비", "계약비", "유지비", "손해조사비", "투자관리비",
]


# ── 데이터셋별 필수 컬럼: 없으면 JOIN이 전량 NaN으로 오염되므로 즉시 에러 ──────
_REQUIRED_TRANSACTION_COLS: list[str] = [
    COLUMN_MAP["원가요소"],    # 필터 키 + JOIN 2 키
    COLUMN_MAP["코스트센터"],  # JOIN 1 키
]

_REQUIRED_CCM_COLS: list[str] = [
    COLUMN_MAP["cc_code"],     # JOIN 1 키
    COLUMN_MAP["직간접구분"],  # 직간접 분류 기준
]

_REQUIRED_ACCOUNT_COLS: list[str] = [
    COLUMN_MAP["계정번호"],    # JOIN 2 키
]


# ── 데이터셋에서 실제로 필요한 컬럼 목록 ──────────────────────────────────────
def _get_ccm_cols_needed() -> list[str]:
    return [COLUMN_MAP["cc_code"], COLUMN_MAP["팀"], COLUMN_MAP["직간접구분"]]


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
    ]


def _get_output_cols_needed() -> list[str]:
    return (
        [COLUMN_MAP["출력전표"], COLUMN_MAP["귀속_주관부서"], COLUMN_MAP["귀속_사용부서"]]
        + BASIS_TEXT_COLS
    )


# ════════════════════════════════════════════════════════════════════════════
#  Public API
# ════════════════════════════════════════════════════════════════════════════

def _apply_config_overrides(config: dict | None) -> None:
    """config의 columns / nature_cols / col_team 설정을 전역 상수에 반영한다."""
    global NATURE_COLS, COL_TEAM
    global _REQUIRED_TRANSACTION_COLS, _REQUIRED_CCM_COLS, _REQUIRED_ACCOUNT_COLS

    if not config:
        return

    for k, v in config.get("columns", {}).items():
        if k in COLUMN_MAP and v:
            COLUMN_MAP[k] = v
    if "nature_cols" in config and isinstance(config["nature_cols"], list):
        NATURE_COLS = list(config["nature_cols"])
    if "col_team" in config and config["col_team"]:
        COL_TEAM = str(config["col_team"])
        COLUMN_MAP["팀"] = COL_TEAM

    # COLUMN_MAP 변경 반영 — 필수 컬럼 목록 재구성
    _REQUIRED_TRANSACTION_COLS = [COLUMN_MAP["원가요소"], COLUMN_MAP["코스트센터"]]
    _REQUIRED_CCM_COLS         = [COLUMN_MAP["cc_code"],  COLUMN_MAP["직간접구분"]]
    _REQUIRED_ACCOUNT_COLS     = [COLUMN_MAP["계정번호"]]


def _read_csv(path: Path) -> pd.DataFrame:
    """CSV를 모든 값 str로 읽는다. utf-8-sig 우선, 실패 시 cp949 폴백."""
    for enc in ("utf-8-sig", "cp949"):
        try:
            return pd.read_csv(path, dtype=str, keep_default_na=True, encoding=enc)
        except UnicodeDecodeError:
            continue
    # 마지막 폴백 — 오류 무시하고 utf-8
    return pd.read_csv(path, dtype=str, keep_default_na=True, encoding="utf-8", encoding_errors="replace")


def check_unregistered_teams(file_paths: dict, config: dict | None = None) -> list[str]:
    """
    실제 발생 팀 목록(Set A)과 출력.csv 등록 부서(Set B)를 비교하여
    Set A - Set B (미등록 팀)를 반환한다.

    Set A: 사업비정보의 코스트센터 → 부서정보 JOIN으로 얻은 팀.
           코스트센터가 부서정보에 없으면 "(미매칭:{코스트센터코드})"로 포함하여
           등록되지 않은 코스트센터도 미등록 경고에 잡히도록 한다.
    Set B: 출력.csv의 주관부서 ∪ 사용부서.

    load_all_csvs() 이전에 호출되므로, 함수 진입 시 _apply_config_overrides(config)를
    먼저 호출하여 config의 컬럼명 오버라이드(COLUMN_MAP)가 반영된 상태에서 검증한다.

    Args:
        file_paths: {"transaction":…, "account":…, "ccm":…, "output":…} 경로 dict
        config:     column_config.yaml 설정 (None이면 기본값 사용)

    Returns:
        미등록 팀 이름 목록 (정렬됨). 없으면 [].
    """
    _apply_config_overrides(config)   # COLUMN_MAP 오버라이드 선반영 (load_all_csvs보다 먼저 호출됨)

    col_cc   = COLUMN_MAP["코스트센터"]
    col_code = COLUMN_MAP["cc_code"]
    col_team = COLUMN_MAP["팀"]
    col_주관 = COLUMN_MAP["귀속_주관부서"]
    col_사용 = COLUMN_MAP["귀속_사용부서"]

    # ── Set A: 사업비정보 코스트센터 → 부서정보 팀 매핑 ─────────────────────
    try:
        df_tx  = _read_csv(Path(file_paths["transaction"]))
        df_ccm = _read_csv(Path(file_paths["ccm"]))
    except Exception:
        return []

    if col_cc not in df_tx.columns:
        return []

    # 부서정보: 코스트센터 코드 → 팀 lookup
    team_lookup: dict[str, str] = {}
    if col_code in df_ccm.columns and col_team in df_ccm.columns:
        for _, row in df_ccm[[col_code, col_team]].iterrows():
            code = str(row[col_code]).strip()
            team = str(row[col_team]).strip()
            if code and code.lower() != "nan":
                team_lookup[code] = team

    set_a: set[str] = set()
    for raw_cc in df_tx[col_cc].dropna():
        cc = str(raw_cc).strip().replace(".0", "") if str(raw_cc).strip().endswith(".0") else str(raw_cc).strip()
        if not cc or cc.lower() == "nan":
            continue
        team = team_lookup.get(cc, "")
        if team and team.lower() != "nan" and team != FALLBACK_TEXT:
            set_a.add(team)
        else:
            # 부서정보에 없는 코스트센터 — 코드 자체를 미매칭 표시로 포함
            set_a.add(f"(미매칭:{cc})")

    if not set_a:
        return []

    # ── Set B: 출력.csv 등록 부서 합집합 ─────────────────────────────────────
    try:
        df_out = _read_csv(Path(file_paths["output"]))
    except Exception:
        return []

    set_b: set[str] = set()
    for col in [col_주관, col_사용]:
        if col in df_out.columns:
            vals = df_out[col].dropna().astype(str).str.strip()
            set_b.update(vals[~vals.isin({"", "nan", FALLBACK_TEXT})].tolist())

    return sorted(set_a - set_b)


def load_all_csvs(file_paths: dict, config: dict | None = None) -> dict[str, pd.DataFrame]:
    """
    4개 CSV를 로드하고 Phase 0 클렌징까지 완료한 DataFrame dict를 반환한다.

    반환 키:
        "사업비정보", "계정정보", "부서정보", "출력"

    Args:
        file_paths: {"transaction":…, "account":…, "ccm":…, "output":…} 경로 dict
        config:     column_config.yaml 설정 (None이면 기본값 사용)

    Returns:
        클렌징이 완료된 DataFrame을 담은 dict
    """
    _apply_config_overrides(config)

    df_tx      = _read_csv(Path(file_paths["transaction"]))
    df_account = _read_csv(Path(file_paths["account"]))
    df_ccm     = _read_csv(Path(file_paths["ccm"]))
    df_output  = _read_csv(Path(file_paths["output"]))

    return {
        _SHEET_TRANSACTION: _preprocess_actual(df_tx),
        _SHEET_ACCOUNT:     _preprocess_account(df_account),
        _SHEET_CCM:         _preprocess_ccm(df_ccm),
        _SHEET_OUTPUT:      _preprocess_output(df_output),
    }


# ════════════════════════════════════════════════════════════════════════════
#  Phase 0 — 클렌징 핵심 함수
# ════════════════════════════════════════════════════════════════════════════

def cleanse_types(df: pd.DataFrame, join_key_cols: list[str]) -> pd.DataFrame:
    """
    JOIN 키 컬럼을 str 타입으로 강제 형변환한다.

    CSV에서 숫자로 저장된 코드(예: 1234)가 한쪽은 "1234", 다른 쪽은 "1234.0"으로
    로드되면 JOIN 결과가 전량 NaN으로 처리되는 조용한 버그가 발생한다.
    형변환과 공백 제거를 동시에 적용하고, 변환 후 생성된 "nan" 문자열을 제거한다.
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
    """object dtype 컬럼 전체에 .str.strip()을 적용한다."""
    str_cols = df.select_dtypes(include="object").columns
    for col in str_cols:
        df[col] = df[col].str.strip()
    return df


def cleanse_nulls(df: pd.DataFrame, numeric_cols: list[str]) -> pd.DataFrame:
    """숫자 컬럼의 NaN을 fillna(0)으로 처리한다."""
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def _validate_columns(
    df: pd.DataFrame,
    required_cols: list[str],
    dataset_name: str,
) -> None:
    """
    필수 컬럼이 DataFrame에 모두 존재하는지 검증한다.

    JOIN 키 컬럼이 없으면 merge() 결과가 전량 NaN이 되어 '(미매칭)'으로
    조용히 오염되는 문제를 사전에 차단한다.

    Raises:
        ValueError: 필수 컬럼 중 하나라도 누락된 경우
    """
    missing = [c for c in required_cols if c not in df.columns]
    if not missing:
        return
    missing_lines = "\n".join(f"  - {c}" for c in missing)
    actual_lines  = "\n".join(f"  - {c}" for c in df.columns.tolist())
    raise ValueError(
        f"[{dataset_name}] 파일에 필수 컬럼이 없습니다.\n\n"
        f"누락된 컬럼:\n{missing_lines}\n\n"
        f"실제 컬럼 목록:\n{actual_lines}\n\n"
        "CSV 파일의 컬럼명과 column_config.yaml의 columns 정의를 비교하여 수정하세요."
    )


# ════════════════════════════════════════════════════════════════════════════
#  데이터셋별 전처리 (로드 직후 Phase 0 클렌징 호출)
# ════════════════════════════════════════════════════════════════════════════

def _drop_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    """pandas가 생성한 'Unnamed: N' 컬럼을 제거한다."""
    unnamed = [c for c in df.columns if str(c).startswith("Unnamed")]
    return df.drop(columns=unnamed) if unnamed else df


def _preprocess_actual(df: pd.DataFrame) -> pd.DataFrame:
    """
    사업비정보.csv 전처리.

    적용 순서:
        1. Unnamed 컬럼 제거
        2. 필수 컬럼 검증 (JOIN 키)
        3. cleanse_types  — JOIN 키 (코스트센터, 원가요소)
        4. cleanse_whitespace — 텍스트 컬럼 전체
        5. cleanse_nulls  — 성격별 분류 컬럼 + 합계
    """
    df = _drop_unnamed(df)

    _validate_columns(df, _REQUIRED_TRANSACTION_COLS, _SHEET_TRANSACTION)

    df = cleanse_types(df, [COLUMN_MAP["코스트센터"], COLUMN_MAP["원가요소"]])
    df = cleanse_whitespace(df)

    _numeric_cols = NATURE_COLS + [COL_TOTAL_AMOUNT]
    df = cleanse_nulls(df, _numeric_cols)

    return df


def _preprocess_ccm(df: pd.DataFrame) -> pd.DataFrame:
    """
    부서정보.csv 전처리.

    적용 순서:
        1. 필수 컬럼 검증 (JOIN 키)
        2. cleanse_types  — JOIN 키 (Cost Center Code)
        3. cleanse_whitespace
        4. 필요 컬럼만 선택 및 키 결측/중복 행 제거 (JOIN fan-out 방지)
    """
    df = _drop_unnamed(df)

    _validate_columns(df, _REQUIRED_CCM_COLS, _SHEET_CCM)

    df = cleanse_types(df, [COLUMN_MAP["cc_code"]])
    df = cleanse_whitespace(df)

    available = [c for c in _get_ccm_cols_needed() if c in df.columns]
    df = df[available].copy()

    _cc = COLUMN_MAP["cc_code"]
    df = df.dropna(subset=[_cc])
    df = df[df[_cc] != ""]
    df = df.drop_duplicates(subset=[_cc])

    return df


def _preprocess_account(df: pd.DataFrame) -> pd.DataFrame:
    """
    계정정보.csv 전처리.

    적용 순서:
        1. 필수 컬럼 검증 (JOIN 키)
        2. cleanse_types  — JOIN 키 (계정번호)
        3. cleanse_whitespace
        4. 필요 컬럼만 선택 및 중복 키 제거
    """
    df = _drop_unnamed(df)

    _validate_columns(df, _REQUIRED_ACCOUNT_COLS, _SHEET_ACCOUNT)

    df = cleanse_types(df, [COLUMN_MAP["계정번호"]])
    df = cleanse_whitespace(df)

    available = [c for c in _get_account_cols_needed() if c in df.columns]
    df = df[available].copy()

    _acct = COLUMN_MAP["계정번호"]
    if _acct in df.columns:
        df = df.drop_duplicates(subset=[_acct])

    return df


def _preprocess_output(df: pd.DataFrame) -> pd.DataFrame:
    """
    출력.csv 전처리 — 평면 테이블이므로 헤더 위치 탐색 불필요.

    출력전표/주관부서/사용부서 + 분류 근거 7컬럼을 그대로 보존하고
    JOIN 키(출력전표) 형변환 및 텍스트 공백 제거만 수행한다.
    """
    if df.empty:
        return df

    df = _drop_unnamed(df)

    # 출력전표를 JOIN 키(원가요소)와 동일 규칙으로 형변환 (예: "5001.0" → "5001")
    df = cleanse_types(df, [COLUMN_MAP["출력전표"]])
    df = cleanse_whitespace(df)

    # 필요 컬럼만 선택 (존재하는 것만)
    available = [c for c in _get_output_cols_needed() if c in df.columns]
    if available:
        df = df[available].copy()

    return df
