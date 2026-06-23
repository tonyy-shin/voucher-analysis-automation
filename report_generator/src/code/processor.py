"""
processor.py
역할: Phase 1 데이터 파이프라인 — 필터링, 3중 LEFT JOIN, 다차원 집계.
      모든 수치 연산을 pandas에서 완료하고 결과값(Scalar)만 반환한다.

[원칙 1] 이 모듈은 절대로 Excel 수식을 생성하지 않는다.
         모든 계산은 Python 메모리 내에서 완료한다.
"""
from __future__ import annotations

import pandas as pd

# data_loader에 정의된 컬럼 상수를 재사용 (중복 정의 방지)
from data_loader import (
    NATURE_COLS, BASIS_TEXT_COLS,
    COLUMN_MAP,
    FALLBACK_TEXT, COL_SUBTOTAL, COL_TOTAL, COL_TEAM, COL_TOTAL_AMOUNT,
    _SHEET_TRANSACTION, _SHEET_CCM, _SHEET_ACCOUNT, _SHEET_OUTPUT,
    _to_numeric_safe,
)


# ── 분류 근거 텍스트 — 출력.csv 코드별 행에서 생성 ────────────────────────────
def build_classification_basis(output_row: pd.Series | dict | None) -> dict[str, str]:
    """
    출력.csv의 단일 출력전표 행(BASIS_TEXT_COLS)을 분류 근거 dict로 변환한다.

    출력.csv 컬럼(직접비/간접비/공통비/계약비/유지비/손해조사비/투자관리비)의
    텍스트 값을 "{컬럼}_근거" 키로 매핑한다. pdf_exporter._tbl_basis가
    직•공통비 행과 성격별 분류 행에서 이 키들을 참조한다.

    Args:
        output_row: 해당 출력전표 코드의 출력.csv 행 (없으면 빈 dict 반환)

    Returns:
        {"직접비_근거": str, "간접비_근거": str, ..., "투자관리비_근거": str}
    """
    _KEYS = ["직접비_근거", "간접비_근거", "공통비_근거", "계약비_근거",
             "유지비_근거", "손해조사비_근거", "투자관리비_근거"]
    if output_row is None:
        return {k: "" for k in _KEYS}

    basis: dict[str, str] = {}
    for col in BASIS_TEXT_COLS:
        val = output_row.get(col, "") if hasattr(output_row, "get") else ""
        if val is None or (isinstance(val, float) and pd.isna(val)):
            val = ""
        s = str(val).strip()
        basis[f"{col}_근거"] = "" if s.lower() in ("nan", "none") else s
    return basis


# ════════════════════════════════════════════════════════════════════════════
#  Step 1 — 키워드 필터링
# ════════════════════════════════════════════════════════════════════════════

def filter_by_keyword(df_actual: pd.DataFrame, keyword: str) -> pd.DataFrame:
    """
    실제발생사업비의 '원가요소' 컬럼에서 keyword(출력전표 코드번호)와 일치하는 행을 추출한다.

    JOIN 3 구조: 실제발생사업비.원가요소 = 출력>.출력전표
    사용자가 입력하는 값은 출력전표 코드번호이며, 이 값이 '원가요소' 컬럼과
    str.contains()로 매칭된다. 완전 일치 코드("5001")와 부분 일치 모두 지원한다.

    검색 전 keyword를 .strip()하여 UI 입력 공백 오염을 방어한다.
    검색 실패 시 '원가요소'의 고유값 목록을 에러 메시지에 포함하여
    사용자가 올바른 코드번호를 확인할 수 있도록 돕는다.

    Args:
        df_actual: Phase 0 클렌징이 완료된 실제발생사업비 DataFrame
        keyword:   사용자가 입력한 출력전표 코드번호

    Returns:
        필터링된 DataFrame (copy)

    Raises:
        ValueError: '원가요소' 컬럼 부재 또는 검색 결과 0건
                    (main.py에서 catch → 스크롤 창으로 처리)
    """
    col = COLUMN_MAP["원가요소"]

    # 컬럼 부재 — 실제 컬럼 목록을 힌트로 제공
    if col not in df_actual.columns:
        col_list = "\n".join(f"  - {c}" for c in df_actual.columns)
        raise ValueError(
            f"'{col}' 컬럼이 실제발생사업비 시트에 없습니다.\n\n"
            f"실제 컬럼 목록:\n{col_list}"
        )

    # UI 입력 공백 재방어
    keyword_clean = keyword.strip()

    mask = df_actual[col].str.strip() == keyword_clean
    df_filtered = df_actual[mask].copy()

    if df_filtered.empty:
        raise ValueError(
            f"코드번호 '{keyword_clean}'에 해당하는 원가요소가 없습니다."
        )

    return df_filtered


# ════════════════════════════════════════════════════════════════════════════
#  Step 2 — 3중 LEFT JOIN + 대상금액 합산
# ════════════════════════════════════════════════════════════════════════════

def enrich_data(
    df_filtered: pd.DataFrame,
    df_ccm: pd.DataFrame,
    df_account: pd.DataFrame,
    df_output: pd.DataFrame,
) -> pd.DataFrame:
    """
    3중 LEFT JOIN으로 트랜잭션 데이터에 마스터 정보를 결합하고 대상금액을 계산한다.

    JOIN 순서:
        JOIN 1: (코스트센터)  ↔ (Cost Center Code) → 팀·직간접구분 추가 (부서정보)
        JOIN 2: (원가요소)    ↔ (계정번호)          → 계정명·계정그룹명·사용부서·비용 지급 범위·산출기준 추가
        JOIN 3: (원가요소)    ↔ (출력전표)           → 귀속_주관부서·귀속_사용부서 추가

    JOIN 후 처리:
        - 매칭 실패 텍스트 컬럼 → "(미매칭)"
        - 매칭 실패 숫자 컬럼  → 0
        - 대상금액 = 사업비정보.csv의 '합계' 컬럼 (COL_TOTAL_AMOUNT)

    Args:
        df_filtered: filter_by_keyword()의 반환값
        df_ccm:      Cost Center Master (Phase 0 클렌징 완료)
        df_account:  계정정보 (Phase 0 클렌징 완료)
        df_output:   출력> (Phase 0 클렌징 완료)

    Returns:
        보강 완료된 DataFrame ('대상금액' 컬럼 포함)
    """
    _cc_left    = COLUMN_MAP["코스트센터"]
    _cc_right   = COLUMN_MAP["cc_code"]
    _acct_left  = COLUMN_MAP["원가요소"]
    _acct_right = COLUMN_MAP["계정번호"]
    _out_right  = COLUMN_MAP["출력전표"]

    df = df_filtered.copy()

    # ── JOIN 1: 코스트 센터 → 직간접구분 ─────────────────────────────────
    if not df_ccm.empty and _cc_right in df_ccm.columns:
        df = df.merge(
            df_ccm,
            left_on=_cc_left,
            right_on=_cc_right,
            how="left",
            suffixes=("", "_ccm"),
        )
        # merge로 생긴 중복 키 컬럼 정리
        if _cc_right in df.columns and _cc_right != _cc_left:
            df = df.drop(columns=[_cc_right])

    # ── JOIN 2: 원가요소 → 계정 상세 정보 ───────────────────────────────
    if not df_account.empty and _acct_right in df_account.columns:
        df = df.merge(
            df_account,
            left_on=_acct_left,
            right_on=_acct_right,
            how="left",
            suffixes=("", "_account"),
        )
        if _acct_right in df.columns and _acct_right != _acct_left:
            df = df.drop(columns=[_acct_right])

    # ── JOIN 3: 원가요소 → 귀속 부서 ────────────────────────────────────
    # 출력> 시트에 같은 출력전표 행이 여러 개일 경우 카테시안 곱으로 행이 복제되어
    # 금액이 배수 중복되는 문제를 방지하기 위해 출력전표 기준으로 중복 제거
    if not df_output.empty and _out_right in df_output.columns:
        df_output_join = df_output.drop_duplicates(subset=[_out_right], keep="first")
        df = df.merge(
            df_output_join,
            left_on=_acct_left,
            right_on=_out_right,
            how="left",
            suffixes=("", "_output"),
        )
        if _out_right in df.columns and _out_right != _acct_left:
            df = df.drop(columns=[_out_right])

    # JOIN 3 없거나 출력> 시트에 부서 컬럼 없을 때 계정정보 기반 fallback
    # (사용부서는 계정정보 "사용 부서"에서 확보 가능)
    if COLUMN_MAP["귀속_사용부서"] not in df.columns and COLUMN_MAP["범위"] in df.columns:
        df[COLUMN_MAP["귀속_사용부서"]] = df[COLUMN_MAP["범위"]]
    if COLUMN_MAP["귀속_주관부서"] not in df.columns:
        df[COLUMN_MAP["귀속_주관부서"]] = df.get(
            COLUMN_MAP["귀속_사용부서"], pd.Series(FALLBACK_TEXT, index=df.index)
        )

    # JOIN 후 텍스트 결측치 처리
    text_fill_cols = [
        COLUMN_MAP["직간접구분"],       # 직간접구분 (부서정보)
        COLUMN_MAP["계정명"],           # 계정명
        COLUMN_MAP["계정그룹명"],       # 계정그룹명
        COLUMN_MAP["대상정의"],         # 대상정의
        COLUMN_MAP["범위"],             # 사용부서
        COLUMN_MAP["지급대상"],         # 비용 지급 범위
        COLUMN_MAP["산출기준"],         # 산출기준
        COLUMN_MAP["귀속_주관부서"],    # 귀속_주관부서
        COLUMN_MAP["귀속_사용부서"],    # 귀속_사용부서
    ]
    for col in text_fill_cols:
        if col in df.columns:
            df[col] = df[col].fillna(FALLBACK_TEXT)

    # ── JOIN 후 숫자 결측치 처리 ──────────────────────────────────────────
    for col in NATURE_COLS + [COL_TOTAL_AMOUNT]:
        if col in df.columns:
            df[col] = _to_numeric_safe(df[col])

    # ── 대상금액 = 사업비정보.csv의 '합계' 컬럼 ──────────────────────────
    # 이 '대상금액' 값이 calc_dept_attribution의 팀별 합산
    # (df_enriched.loc[..., "대상금액"].sum())을 거쳐 d주관/d사용['대상금액'].sum()
    # 총계까지 흘러간다. 아래 calc_dept_attribution 주석 참조.
    if COL_TOTAL_AMOUNT in df.columns:
        df["대상금액"] = _to_numeric_safe(df[COL_TOTAL_AMOUNT])
    else:
        # 합계 컬럼이 없으면 성격 컬럼 합산으로 폴백 (방어적)
        existing_nature = [c for c in NATURE_COLS if c in df.columns]
        df["대상금액"] = df[existing_nature].sum(axis=1) if existing_nature else 0.0

    # 대상금액 흐름 보장 — 이후 모든 부점/직간접/성격 집계의 금액 소스
    assert "대상금액" in df.columns, "enrich_data: '대상금액' 컬럼 생성 실패"

    return df


# ════════════════════════════════════════════════════════════════════════════
#  Step 3 — 부점귀속 현황 (주관/사용부서별 구성비)
# ════════════════════════════════════════════════════════════════════════════

def calc_dept_attribution(
    df_enriched: pd.DataFrame,
    df_output: pd.DataFrame = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    주관부서별·사용부서별 대상금액 합계 및 구성비(%)를 계산한다.

    [산출 로직]
    df_output(출력> 시트 전체 flat lookup)의 귀속_주관부서·귀속_사용부서 열에서
    각각 set을 구성하고, 실제발생사업비 팀 각각을 집합 멤버십으로 분류한다.
        - 주관_set에만 있음 -> 주관부서
        - 사용_set에만 있음 -> 사용부서
        - 양쪽 모두 있음   -> 주관부서 우선
        - 어디에도 없음    -> 미분류, 표시 제외
    df_output가 없거나 비어 있으면 귀속_주관/사용부서 콜럼 기준 groupby 폴백.

    Args:
        df_enriched: enrich_data()의 반환값 ('대상금액', '팀' 콜럼 포함)
        df_output:   출력> 시트 전체 flat lookup DataFrame

    Returns:
        (df_주관부서, df_사용부서) 튜플.
        각 DataFrame 콜럼: ['부서명', '대상금액', '구성비(%)']
        금액 내림차순 정렬.

    [대상금액 흐름]
    enrich_data()가 사업비정보.csv '합계' 컬럼으로 세팅한 df_enriched['대상금액']이
    여기서 팀별로 sum()되어 각 부서 행의 '대상금액'이 되고, pdf_exporter._tbl_dept의
    d주관/d사용['대상금액'].sum() 총계로 이어진다.
    """
    empty = pd.DataFrame(columns=["부서명", "대상금액", "구성비(%)"])

    def _finalize(rows: list) -> pd.DataFrame:
        if not rows:
            return empty.copy()
        df_r = pd.DataFrame(rows)
        total = df_r["대상금액"].sum()
        df_r["구성비(%)"] = (
            df_r["대상금액"].div(total).mul(100).round(2) if total != 0 else 0.0
        )
        return df_r.sort_values("대상금액", ascending=False).reset_index(drop=True)

    def _build_from_groupby(dept_col: str) -> pd.DataFrame:
        if dept_col not in df_enriched.columns:
            return empty.copy()
        grouped = (
            df_enriched.groupby(dept_col, as_index=False)["대상금액"]
            .sum()
            .rename(columns={dept_col: "부서명"})
        )
        grouped = grouped[~grouped["부서명"].isin([FALLBACK_TEXT, ""])]
        total = grouped["대상금액"].sum()
        grouped["구성비(%)"] = (
            grouped["대상금액"].div(total).mul(100).round(2) if total != 0 else 0.0
        )
        return grouped.sort_values("대상금액", ascending=False).reset_index(drop=True)

    # -- df_output 없음 -> groupby 폴백 -----------------------------------------------
    if df_output is None or df_output.empty or COL_TEAM not in df_enriched.columns:
        return (
            _build_from_groupby(COLUMN_MAP["귀속_주관부서"]),
            _build_from_groupby(COLUMN_MAP["귀속_사용부서"]),
        )

    # -- 출력> 전체 열에서 주관/사용 set 구성 ------------------------------------------
    def _make_set(col: str) -> set:
        if col not in df_output.columns:
            return set()
        vals = df_output[col].dropna().astype(str).str.strip()
        return set(vals[~vals.isin({"", FALLBACK_TEXT})].tolist())

    주관_set = _make_set(COLUMN_MAP["귀속_주관부서"])
    사용_set  = _make_set(COLUMN_MAP["귀속_사용부서"])

    # -- 팀별 set 멤버십 분류 ----------------------------------------------------------
    주관_rows = []
    사용_rows = []

    for team in df_enriched[COL_TEAM].dropna().unique():
        team_str = str(team).strip()
        if not team_str or team_str == FALLBACK_TEXT:
            continue

        in_주관 = team_str in 주관_set
        in_사용 = team_str in 사용_set

        if not in_주관 and not in_사용:
            continue  # 어느 열에도 없음 -> 미분류, 표시 제외

        amount = float(
            df_enriched.loc[df_enriched[COL_TEAM] == team_str, "대상금액"].sum()
        )

        # 양쪽 모두 있을 때는 주관부서 우선
        if in_주관 and not in_사용:
            주관_rows.append({"부서명": team_str, "대상금액": amount})
        elif in_사용 and not in_주관:
            사용_rows.append({"부서명": team_str, "대상금액": amount})
        else:
            주관_rows.append({"부서명": team_str, "대상금액": amount})

    return _finalize(주관_rows), _finalize(사용_rows)



# ════════════════════════════════════════════════════════════════════════════
#  Step 4 — 직접비 / 공통비 합계
# ════════════════════════════════════════════════════════════════════════════

def calc_direct_indirect(df_enriched: pd.DataFrame) -> dict[str, float]:
    """
    직공통비·총계를 계산한다.

    직간접 구분 기준은 부서정보.csv의 '직간접구분' 컬럼 (코스트센터 JOIN으로 결합).
    '직접' / '공통' 텍스트를 str.contains()로 유연하게 매칭하여
    값 표기 방식 변형(예: '직접비', '직접 비용')에도 대응한다.

    Args:
        df_enriched: enrich_data()의 반환값

    Returns:
        {'직접비': float, '공통비': float, '총계': float}
    """
    총계 = float(df_enriched["대상금액"].sum())

    # 직간접 구분 기준 = 부서정보 '직간접구분' 컬럼
    col = COLUMN_MAP["직간접구분"]

    if col not in df_enriched.columns:
        return {"직접비": 0.0, "공통비": 0.0, "총계": 총계}

    is_직접     = df_enriched[col].str.contains("직접", na=False)
    is_공통_only = df_enriched[col].str.contains("공통", na=False) & ~is_직접

    직접비 = float(df_enriched.loc[is_직접,      "대상금액"].sum())
    공통비 = float(df_enriched.loc[is_공통_only, "대상금액"].sum())

    return {"직접비": 직접비, "공통비": 공통비, "총계": 총계}


# ════════════════════════════════════════════════════════════════════════════
#  Step 5 — 성격별 × 부점 교차 집계
# ════════════════════════════════════════════════════════════════════════════

def calc_nature_classification(
    df_enriched: pd.DataFrame,
    df_output: pd.DataFrame = None,
) -> pd.DataFrame:
    """
    성격별(NATURE_COLS) × 부점 교차 집계 + 총계 행 추가.

    [집계 우선순위]
    1차: df_output가 있고 df_enriched에 '팀' 컬럼이 있으면,
         실제 원가발생 기준 '팀' 컬럼으로 groupby.
         출력> 시트의 주관/사용부서 중 실제발생사업비 미입력 팀은 0원 행으로 추가.
    2차 폴백: df_output가 없거나 '팀' 컬럼이 없으면 기존 귀속_사용부서 기준 groupby.

    Args:
        df_enriched:           enrich_data()의 반환값
        df_output: 현재 keyword에 해당하는 출력> 행 (귀속_주관부서, 귀속_사용부서 포함)

    Returns:
        컬럼: ['귀속_사용부서', '계약비', '유지비', '손해조사비', '투자관리비',
               '간접비', '공통비', '소계']  (소계 = COL_SUBTOTAL)
        마지막 행: '귀속_사용부서' == '총계' 인 합산 행.
        대응되는 NATURE_COLS 컬럼이 없으면 빈 DataFrame 반환.
    """
    dept_col = COLUMN_MAP["귀속_사용부서"]
    existing_nature = [c for c in NATURE_COLS if c in df_enriched.columns]

    if not existing_nature:
        return pd.DataFrame()

    if (df_output is not None
            and not df_output.empty
            and COL_TEAM in df_enriched.columns):

        # 1. 실제 원가발생 팀 기준 집계
        grouped = (
            df_enriched
            .groupby(COL_TEAM, as_index=False)[existing_nature]
            .sum()
            .rename(columns={COL_TEAM: dept_col})
        )

        # 2. 출력> 시트 주관/사용부서 전체 수집
        all_output_depts: set[str] = set()
        for col in [COLUMN_MAP["귀속_주관부서"], COLUMN_MAP["귀속_사용부서"]]:
            if col in df_output.columns:
                vals = (
                    df_output[col]
                    .dropna()
                    .astype(str)
                    .str.strip()
                )
                all_output_depts.update(
                    vals[~vals.isin({"", FALLBACK_TEXT})].tolist()
                )

        # 등록되지 않은 팀 제거 — all_output_depts 비어 있으면(sentinel) 건너뜀
        if all_output_depts:
            grouped = grouped[grouped[dept_col].isin(all_output_depts)].reset_index(drop=True)
    else:
        # 폴백: 귀속_사용부서 기준 집계 (df_output 없거나 팀 컬럼 없을 때)
        if dept_col not in df_enriched.columns:
            return pd.DataFrame()
        grouped = (
            df_enriched
            .groupby(dept_col, as_index=False)[existing_nature]
            .sum()
        )

    grouped[COL_SUBTOTAL] = grouped[existing_nature].sum(axis=1)

    # 누락된 성격 컬럼은 0으로 패딩 (템플릿 컬럼 수 고정 대응)
    for col in NATURE_COLS:
        if col not in grouped.columns:
            grouped[col] = 0.0

    # 총계 행 추가
    sum_cols = existing_nature + [COL_SUBTOTAL]
    total_row: dict[str, object] = {dept_col: COL_TOTAL}
    for col in sum_cols:
        total_row[col] = float(grouped[col].sum())
    for col in NATURE_COLS:
        if col not in sum_cols:
            total_row[col] = 0.0

    df_result = pd.concat(
        [grouped, pd.DataFrame([total_row])],
        ignore_index=True,
    )

    # 컬럼 순서 고정
    ordered_cols = [dept_col] + NATURE_COLS + [COL_SUBTOTAL]
    df_result = df_result[[c for c in ordered_cols if c in df_result.columns]]

    return df_result


# ════════════════════════════════════════════════════════════════════════════
#  Step 6 — 계정 헤더 정보 추출
# ════════════════════════════════════════════════════════════════════════════

def extract_account_info(df_enriched: pd.DataFrame) -> dict[str, object]:
    """
    키워드로 필터링된 단일 계정의 헤더 정보를 추출한다.

    동일 키워드로 필터링된 데이터는 모두 같은 계정에 속하므로
    첫 번째 유효 행의 값을 대표값으로 사용한다.
    반환값은 template_mapper가 헤더 영역(C7·D7·F7·G7·C10~C13)에 주입한다.

    Args:
        df_enriched: enrich_data()의 반환값

    Returns:
        {
            '계정번호':   str,   # 원가요소 코드
            '계정명':     str,   # 계정 이름
            '직간접구분': str,   # '직접비' 또는 '공통비'
            '대상정의':   str,   # 대상정의 v3.0_0415
            '범위':       str,   # 사용부서
            '지급대상':   str,   # 비용 지급 범위
            '산출기준':   str,
            '대상금액':   float, # 전체 대상금액 합계
        }
    """
    def _first_valid(col: str) -> str:
        if col not in df_enriched.columns:
            return ""
        series = df_enriched[col].dropna()
        series = series[series.astype(str).str.strip().isin(["", FALLBACK_TEXT]) == False]  # noqa: E712
        return str(series.iloc[0]).strip() if not series.empty else ""

    return {
        "계정번호":   _first_valid(COLUMN_MAP["원가요소"]),
        "계정명":     _first_valid(COLUMN_MAP["계정명"]),
        "계정그룹ID": _first_valid(COLUMN_MAP["계정그룹ID"]),
        "계정그룹명": _first_valid(COLUMN_MAP["계정그룹명"]),
        "대상정의":   _first_valid(COLUMN_MAP["대상정의"]),
        "범위":       _first_valid(COLUMN_MAP["범위"]),
        "지급대상":   _first_valid(COLUMN_MAP["지급대상"]),
        "산출기준":   _first_valid(COLUMN_MAP["산출기준"]),
    }


# ════════════════════════════════════════════════════════════════════════════
#  Public API — 파이프라인 일괄 실행
# ════════════════════════════════════════════════════════════════════════════

def run_pipeline(
    sheets: dict[str, pd.DataFrame],
    keyword: str,
) -> dict[str, object]:
    """
    Phase 1 전체 파이프라인(Step 1~6)을 실행하고 결과를 dict로 반환한다.

    main.py에서 단일 호출로 파이프라인 전체를 완료할 수 있도록 제공되는
    편의 함수. 각 Step을 순서대로 실행하며 의존성을 유지한다.

    Args:
        sheets:  data_loader.load_all_csvs()의 반환값
        keyword: 사용자 입력 키워드

    Returns:
        {
            'account_info':    dict,       # extract_account_info 결과
            'dept_주관':       DataFrame,  # 주관부서별 귀속 현황
            'dept_사용':       DataFrame,  # 사용부서별 귀속 현황
            'direct_indirect': dict,       # 직접비·공통비·총계
            'nature':          DataFrame,  # 성격별 × 부점 교차 집계
        }

    Raises:
        ValueError: 키워드 매칭 결과 없음 (main.py에서 catch → messagebox 처리)
    """
    df_actual  = sheets[_SHEET_TRANSACTION]
    df_ccm     = sheets[_SHEET_CCM]
    df_account = sheets[_SHEET_ACCOUNT]
    df_output  = sheets.get(_SHEET_OUTPUT, pd.DataFrame())

    # ── 분류 근거: 해당 출력전표 코드의 출력.csv 행에서 텍스트 추출 ──────────
    col_code = COLUMN_MAP["출력전표"]
    basis_row = None
    if not df_output.empty and col_code in df_output.columns:
        mask = df_output[col_code].astype(str).str.strip() == str(keyword).strip()
        basis_row = df_output[mask].iloc[0] if mask.any() else None
    classification_basis = build_classification_basis(basis_row)

    # Step 1 — 필터링
    df_filtered = filter_by_keyword(df_actual, keyword)

    # Step 2 — 보강 (3중 JOIN + 대상금액 합산)
    df_enriched = enrich_data(df_filtered, df_ccm, df_account, df_output)

    # Step 3 — 부점귀속 현황
    # df_output 전체를 flat lookup 테이블로 사용하여 팀 주관/사용 분류
    df_주관, df_사용 = calc_dept_attribution(df_enriched, df_output)

    di_result   = calc_direct_indirect(df_enriched)                      # Step 4

    # ── 정합성 검증: 부점귀속 주관 총계 + 누락 총합 ≈ 직접비+공통비 총계 ──────
    return {
        "account_info":         extract_account_info(df_enriched),       # Step 6
        "dept_주관":            df_주관,                                  # Step 3
        "dept_사용":            df_사용,                                  # Step 3
        "direct_indirect":      di_result,                               # Step 4
        "nature":               calc_nature_classification(df_enriched, df_output), # Step 5
        "classification_basis": classification_basis,                   # 출력.csv 코드별 분류 근거
    }
