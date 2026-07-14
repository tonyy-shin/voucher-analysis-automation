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


# ── 분류 근거 텍스트 — 출력.csv 전체에서 컬럼별 첫 non-empty 행으로 생성 ──────
def _first_nonempty(df_output: pd.DataFrame, col: str) -> str:
    """df_output[col] 전체에서 non-empty(NaN/""/nan/none/None 제외) 첫 값을 반환한다."""
    if col not in df_output.columns:
        return ""
    series = df_output[col].dropna().astype(str).str.strip()
    series = series[~series.isin(["", "nan", "none", "None"])]
    return series.iloc[0] if not series.empty else ""


def build_classification_basis(
    df_output: pd.DataFrame | None,
    basis_rows: list[dict],
) -> dict[str, str]:
    """
    섹션4 분류 근거 dict를 생성한다 — legacy 키 + custom 키를 단일 dict로 병합 반환.

    legacy(직공통비/성격별분류) 행은 BASIS_TEXT_COLS 컬럼별 첫 non-empty 텍스트를
    "{컬럼}_근거" 키로 매핑한다(pdf_exporter._tbl_basis가 참조). 단, basis_rows에
    type이 "직공통비" 또는 "성격별분류"인 행이 하나라도 있을 때만 legacy 키를 생성한다.

    custom 행은 row["csv_column"]이 비어 있지 않고 df_output에 존재하면, 해당 컬럼
    전체의 첫 non-empty 값을 cb[csv_column] 키로 저장한다.

    Args:
        df_output:  출력.csv 전체 DataFrame (없거나 비어 있으면 빈/빈값 dict 반환)
        basis_rows: config display_labels.section4.rows (타입별 행 설정 리스트)

    Returns:
        {"직접비_근거": str, ...(legacy), <csv_column>: str, ...(custom)}
    """
    _LEGACY_KEYS = ["직접비_근거", "간접비_근거", "공통비_근거", "계약체결비_근거",
                    "계약유지비_근거", "손해조사비_근거", "투자관리비_근거"]
    rows = basis_rows or []
    has_legacy = any(
        r.get("type") in ("직공통비", "성격별분류") for r in rows
    )

    empty_df = df_output is None or df_output.empty

    basis: dict[str, str] = {}

    # legacy 키 — 해당 타입 행이 존재할 때만 생성
    if has_legacy:
        if empty_df:
            basis.update({k: "" for k in _LEGACY_KEYS})
        else:
            for col in BASIS_TEXT_COLS:
                basis[f"{col}_근거"] = _first_nonempty(df_output, col)

    # custom 키 — csv_column 으로 지정한 컬럼의 첫 non-empty 값
    if not empty_df:
        for r in rows:
            if r.get("type") != "custom":
                continue
            csv_col = r.get("csv_column", "")
            if csv_col and csv_col in df_output.columns:
                basis[csv_col] = _first_nonempty(df_output, csv_col)

    # 참조 키 — 모든 행 공통. 참조_csv_column 으로 지정한 컬럼의 첫 non-empty 값.
    # content 키(csv_column / {col}_근거)와 충돌하지 않도록 "참조__" 접두사를 사용한다.
    if not empty_df:
        for r in rows:
            ref_col = r.get("참조_csv_column", "")
            if ref_col and ref_col in df_output.columns:
                basis["참조__" + ref_col] = _first_nonempty(df_output, ref_col)

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
    group_cols: list[str] | None = None,
) -> list[pd.DataFrame]:
    """
    귀속 그룹별(기본: 주관부서/사용부서) 대상금액 합계 및 구성비(%)를 계산한다.

    [산출 로직]
    df_output(출력> 시트 전체 flat lookup)의 각 그룹 컬럼(group_cols)에서 set을
    구성하고, 실제발생사업비 팀 각각을 집합 멤버십으로 분류한다.
        - 리스트 순서 = 우선순위: 여러 set에 속한 팀은 첫 번째 그룹으로 귀속
          (기본 [주관, 사용] 순서에서 기존 '양쪽 모두 -> 주관 우선' 규칙과 동일)
        - 어디에도 없음 -> 미분류, 표시 제외
    df_output가 없거나 비어 있으면 각 그룹 컬럼 기준 groupby 폴백.

    Args:
        df_enriched: enrich_data()의 반환값 ('대상금액', '팀' 콜럼 포함)
        df_output:   출력> 시트 전체 flat lookup DataFrame
        group_cols:  출력.csv의 부서 목록 컬럼명 리스트
                     (config display_labels.section2.groups 의 csv_column 값들;
                     None/빈 리스트면 기존 주관/사용 쌍으로 폴백)

    Returns:
        group_cols 와 같은 순서의 DataFrame 리스트.
        각 DataFrame 콜럼: ['부서명', '대상금액', '구성비(%)']
        금액 내림차순 정렬.

    [대상금액 흐름]
    enrich_data()가 사업비정보.csv '합계' 컬럼으로 세팅한 df_enriched['대상금액']이
    여기서 팀별로 sum()되어 각 부서 행의 '대상금액'이 되고, pdf_exporter._tbl_dept의
    그룹별 ['대상금액'].sum() 총계로 이어진다.
    """
    if not group_cols:
        group_cols = [COLUMN_MAP["귀속_주관부서"], COLUMN_MAP["귀속_사용부서"]]
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
        return [_build_from_groupby(col) for col in group_cols]

    # -- 출력> 전체 열에서 그룹별 set 구성 ---------------------------------------------
    def _make_set(col: str) -> set:
        if col not in df_output.columns:
            return set()
        vals = df_output[col].dropna().astype(str).str.strip()
        return set(vals[~vals.isin({"", FALLBACK_TEXT})].tolist())

    group_sets = [_make_set(col) for col in group_cols]

    # -- 팀별 set 멤버십 분류 (첫 번째로 일치한 그룹으로 귀속) --------------------------
    group_rows: list[list[dict]] = [[] for _ in group_cols]

    for team in df_enriched[COL_TEAM].dropna().unique():
        team_str = str(team).strip()
        if not team_str or team_str == FALLBACK_TEXT:
            continue

        for gi, gset in enumerate(group_sets):
            if team_str in gset:
                amount = float(
                    df_enriched.loc[df_enriched[COL_TEAM] == team_str, "대상금액"].sum()
                )
                group_rows[gi].append({"부서명": team_str, "대상금액": amount})
                break
        # 어느 열에도 없음 -> 미분류, 표시 제외

    return [_finalize(rows) for rows in group_rows]



# ════════════════════════════════════════════════════════════════════════════
#  Step 4 — 직접비 / 공통비 합계
# ════════════════════════════════════════════════════════════════════════════

def calc_direct_indirect(
    df_enriched: pd.DataFrame,
    category_rows: list[dict] | None = None,
) -> dict[str, object]:
    """
    직간접구분 키워드 분류별 합계 및 총계를 계산한다.

    분류 기준은 부서정보.csv의 '직간접구분' 컬럼 (코스트센터 JOIN으로 결합).
    각 분류의 keyword 를 str.contains(regex=False)로 부분 일치 매칭하여
    값 표기 방식 변형(예: '직접비', '직접 비용')에도 대응한다.
    리스트 순서 = first-match-wins: 이미 위쪽 분류에 잡힌 행은 제외되므로,
    기본 [직접, 공통] 순서에서 기존 '공통 = contains("공통") & ~직접' 규칙과 동일하다.

    Args:
        df_enriched:   enrich_data()의 반환값
        category_rows: config display_labels.section3_1.rows
                       ([{"label": str, "keyword": str}, ...];
                       None/빈 리스트면 기존 직접비/공통비 기본 분류 사용)

    Returns:
        {'rows': [{'label': str, 'amount': float}, ...], '총계': float}
    """
    if not category_rows:
        category_rows = [
            {"label": "직접비", "keyword": "직접"},
            {"label": "공통비", "keyword": "공통"},
        ]

    총계 = float(df_enriched["대상금액"].sum())

    # 직간접 구분 기준 = 부서정보 '직간접구분' 컬럼
    col = COLUMN_MAP["직간접구분"]

    if col not in df_enriched.columns:
        return {
            "rows": [{"label": r.get("label", ""), "amount": 0.0} for r in category_rows],
            "총계": 총계,
        }

    rows: list[dict] = []
    matched = pd.Series(False, index=df_enriched.index)
    for r in category_rows:
        keyword = str(r.get("keyword", ""))
        if keyword:
            m = df_enriched[col].str.contains(keyword, na=False, regex=False) & ~matched
        else:
            m = pd.Series(False, index=df_enriched.index)
        rows.append({
            "label": r.get("label", ""),
            "amount": float(df_enriched.loc[m, "대상금액"].sum()),
        })
        matched |= m

    return {"rows": rows, "총계": 총계}


# ════════════════════════════════════════════════════════════════════════════
#  Step 5 — 성격별 × 부점 교차 집계
# ════════════════════════════════════════════════════════════════════════════

def calc_nature_classification(
    df_enriched: pd.DataFrame,
    df_output: pd.DataFrame = None,
    dept_cols: list[str] | None = None,
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
        dept_cols: 출력.csv의 부서 목록 컬럼 (section2.groups 의 csv_column 값들;
                   None/빈 리스트면 기존 주관/사용 쌍 사용) — 섹션2와 동일한 기준으로
                   등록 부서를 필터링하기 위해 공유한다.

    Returns:
        컬럼: ['귀속_사용부서', '계약체결비', '계약유지비', '손해조사비', '투자관리비',
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

        # 2. 출력> 시트 등록 부서 전체 수집 (section2 그룹 컬럼과 동일 기준)
        all_output_depts: set[str] = set()
        for col in (dept_cols or [COLUMN_MAP["귀속_주관부서"], COLUMN_MAP["귀속_사용부서"]]):
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

def extract_account_info(
    df_enriched: pd.DataFrame,
    field_cols: list[str] | None = None,
) -> dict[str, str]:
    """
    키워드로 필터링된 단일 계정의 헤더 정보를 추출한다.

    동일 키워드로 필터링된 데이터는 모두 같은 계정에 속하므로
    첫 번째 유효 행의 값을 대표값으로 사용한다.
    반환값은 pdf_exporter 가 계정 테이블·섹션1을 csv_column 키로 조회한다.

    Args:
        df_enriched: enrich_data()의 반환값
        field_cols:  추출할 계정정보.csv 컬럼명 리스트
                     (config account_table.columns + section1.rows 의 csv_column 값들;
                     None/빈 리스트면 기존 8개 기본 컬럼)

    Returns:
        {csv_column: 첫 번째 유효 값(str)} — field_cols 의 각 컬럼이 키가 된다.

    [주의 — 계정번호 별칭]
    enrich_data()의 JOIN 2 가 우측 키 컬럼(계정번호)을 drop 하므로 df_enriched 에는
    계정번호 컬럼이 없다. 설정된 컬럼이 계정번호(COLUMN_MAP)와 같으면
    JOIN 키인 원가요소 컬럼에서 값을 읽는다 (동일 값).
    """
    if not field_cols:
        field_cols = [
            COLUMN_MAP["계정번호"],
            COLUMN_MAP["계정명"],
            COLUMN_MAP["계정그룹ID"],
            COLUMN_MAP["계정그룹명"],
            COLUMN_MAP["대상정의"],
            COLUMN_MAP["범위"],
            COLUMN_MAP["지급대상"],
            COLUMN_MAP["산출기준"],
        ]

    def _first_valid(col: str) -> str:
        if col not in df_enriched.columns:
            return ""
        series = df_enriched[col].dropna()
        series = series[series.astype(str).str.strip().isin(["", FALLBACK_TEXT]) == False]  # noqa: E712
        return str(series.iloc[0]).strip() if not series.empty else ""

    def _resolve(col: str) -> str:
        # 계정번호는 JOIN 2 에서 drop 되므로 원가요소(JOIN 키)로 별칭 처리
        if col not in df_enriched.columns and col == COLUMN_MAP["계정번호"]:
            return COLUMN_MAP["원가요소"]
        return col

    return {col: _first_valid(_resolve(col)) for col in field_cols}


# ════════════════════════════════════════════════════════════════════════════
#  Public API — 파이프라인 일괄 실행
# ════════════════════════════════════════════════════════════════════════════

def run_pipeline(
    sheets: dict[str, pd.DataFrame],
    keyword: str,
    labels: dict,
) -> dict[str, object]:
    """
    Phase 1 전체 파이프라인(Step 1~6)을 실행하고 결과를 dict로 반환한다.

    main.py에서 단일 호출로 파이프라인 전체를 완료할 수 있도록 제공되는
    편의 함수. 각 Step을 순서대로 실행하며 의존성을 유지한다.

    레이어링 주의: processor 는 config_manager 를 import 하지 않는다.
    main.py 가 config["display_labels"] 를 명시적으로 넘겨 연결하며,
    여기서는 labels.get(...) 으로 누락 키를 방어한다 (누락 시 빈 결과로 강등).

    Args:
        sheets:  data_loader.load_all_csvs()의 반환값
        keyword: 사용자 입력 키워드
        labels:  config["display_labels"] 전체 dict — 섹션별 동적 리스트
                 (account_table.columns / section1.rows / section2.groups /
                 section3_1.rows / section4.rows)를 여기서 꺼내 각 Step 에 전달한다.

    Returns:
        {
            'account_info':    dict,             # extract_account_info 결과 (csv_column 키)
            'dept_groups':     list[DataFrame],  # section2.groups 순서의 귀속 현황
            'direct_indirect': dict,             # {'rows': [...], '총계': float}
            'nature':          DataFrame,        # 성격별 × 부점 교차 집계
        }

    Raises:
        ValueError: 키워드 매칭 결과 없음 (main.py에서 catch → messagebox 처리)
    """
    df_actual  = sheets[_SHEET_TRANSACTION]
    df_ccm     = sheets[_SHEET_CCM]
    df_account = sheets[_SHEET_ACCOUNT]
    df_output  = sheets.get(_SHEET_OUTPUT, pd.DataFrame())

    # ── display_labels 동적 리스트 추출 (누락/비정상 키는 빈 리스트로 방어) ──
    def _entries(sec_key: str, list_key: str) -> list[dict]:
        sec = labels.get(sec_key) if isinstance(labels, dict) else None
        entries = sec.get(list_key) if isinstance(sec, dict) else None
        if not isinstance(entries, list):
            return []
        return [e for e in entries if isinstance(e, dict)]

    basis_rows = _entries("section4", "rows")
    group_cols = [g["csv_column"] for g in _entries("section2", "groups") if g.get("csv_column")]
    di_rows    = _entries("section3_1", "rows")
    field_cols = list(dict.fromkeys(
        e["csv_column"]
        for e in _entries("account_table", "columns") + _entries("section1", "rows")
        if e.get("csv_column")
    ))

    # ── 분류 근거: 섹션4 행 설정(basis_rows) 기준 legacy + custom 텍스트 추출 ──
    classification_basis = build_classification_basis(df_output, basis_rows)

    # Step 1 — 필터링
    df_filtered = filter_by_keyword(df_actual, keyword)

    # Step 2 — 보강 (3중 JOIN + 대상금액 합산)
    df_enriched = enrich_data(df_filtered, df_ccm, df_account, df_output)

    # Step 3 — 부점귀속 현황
    # df_output 전체를 flat lookup 테이블로 사용하여 그룹별 팀 귀속 분류
    dept_groups = calc_dept_attribution(df_enriched, df_output, group_cols)

    di_result = calc_direct_indirect(df_enriched, di_rows)               # Step 4

    return {
        "account_info":         extract_account_info(df_enriched, field_cols),  # Step 6
        "dept_groups":          dept_groups,                             # Step 3
        "direct_indirect":      di_result,                               # Step 4
        "nature":               calc_nature_classification(df_enriched, df_output, group_cols),  # Step 5
        "classification_basis": classification_basis,                   # 출력.csv 코드별 분류 근거
    }
