"""
csv_exporter.py
역할: template.csv 레이아웃을 processor 결과로 채워 출력전표 코드별 CSV로 저장.
      template_mapper.py + pdf_exporter.py를 완전 대체한다.
"""
from __future__ import annotations

import csv
import logging
import math
import os

import pandas as pd

logger = logging.getLogger(__name__)


# ── 고정 셀 매핑: (0-indexed row, col) → account_info 키 ────────────────────
# template.csv를 csv.reader로 읽으면 헤더 행(Unnamed:0...)도 rows[0]으로 포함된다.
# Excel C7 = csv.reader rows[6][2], Excel C10 = rows[9][2] (row = Excel행번호 - 1)
_FIXED_CELLS: dict[tuple[int, int], str] = {
    (6, 2): "계정번호",
    (6, 3): "계정명",
    (6, 4): "계정그룹ID",
    (6, 5): "계정그룹명",
    (9, 2): "대상정의",
    (10, 2): "범위",
    (11, 2): "지급대상",
    (12, 2): "산출기준",
}

_NUM_COLS = 8  # template 열 수 (Unnamed:0 ~ Unnamed:7)

from data_loader import NATURE_COLS as _NATURE_COLS_BASE, COL_SUBTOTAL
NATURE_COLS = _NATURE_COLS_BASE + [COL_SUBTOTAL]


# ── 공개 함수 ────────────────────────────────────────────────────────────────

def export(results: dict, template_csv_path: str, out_dir: str, code: str) -> str:
    """
    template.csv 레이아웃을 processor 결과로 채워 코드별 CSV 파일로 저장한다.

    Args:
        results:           processor.run_pipeline() 반환 dict
        template_csv_path: template.csv 절대 경로
        out_dir:           출력 폴더 경로
        code:              출력전표 코드 (파일명에 사용)

    Returns:
        생성된 CSV 파일의 절대 경로.
    """
    rows = _load_template(template_csv_path)
    rows = _inject_fixed(rows, results["account_info"])
    rows = _inject_dept(rows, results["dept_주관"], results["dept_사용"])
    rows = _inject_direct_indirect(rows, results["direct_indirect"])
    rows = _inject_nature(rows, results["nature"])
    rows = _inject_classification_basis(rows, results.get("classification_basis", {}))

    out_path = os.path.join(out_dir, f"전표분석_{code}.csv")
    _write_csv(rows, out_path)
    return out_path


# ── 내부 헬퍼 ────────────────────────────────────────────────────────────────

def _load_template(path: str) -> list[list[str]]:
    """template.csv를 2D 리스트로 읽는다. 각 행을 _NUM_COLS열로 정규화."""
    for enc in ("utf-8-sig", "utf-8", "cp949"):
        try:
            with open(path, newline="", encoding=enc) as f:
                raw = list(csv.reader(f))
            break
        except UnicodeDecodeError:
            continue
    result: list[list[str]] = []
    for row in raw:
        padded = (row + [""] * _NUM_COLS)[:_NUM_COLS]
        result.append(padded)
    _validate_template_structure(result)
    return result


def _validate_template_structure(rows: list[list[str]]) -> None:
    """
    template.csv 고정 셀 좌표에 접근할 수 있는지 사전 검증한다.

    행 수가 부족하면 ValueError를 발생시키고,
    이미 값이 있는 셀은 경고 로그를 남겨 의도치 않은 덮어쓰기를 감지한다.
    """
    min_rows = max(r for r, _ in _FIXED_CELLS) + 1
    if len(rows) < min_rows:
        raise ValueError(
            f"template.csv 행 수 부족: 고정 셀 좌표에 접근하려면 최소 {min_rows}행이 필요하지만 "
            f"{len(rows)}행만 있습니다.\n"
            "template.csv의 구조가 변경되었는지 확인하세요."
        )
    for (r, c), key in _FIXED_CELLS.items():
        if len(rows[r]) <= c:
            logger.warning(
                "[template 검증] rows[%d][%d] 셀 접근 불가 — 해당 행 열 수가 %d개입니다. "
                "(예상 키: %s)",
                r, c, len(rows[r]), key,
            )
            continue
        existing = rows[r][c].strip()
        if existing not in ("", "0"):
            logger.warning(
                "[template 검증] rows[%d][%d] 셀에 이미 값 '%s'이 있습니다. "
                "(예상 키: %s) — 덮어쓰기됩니다.",
                r, c, existing, key,
            )


def _inject_fixed(rows: list[list[str]], account_info: dict) -> list[list[str]]:
    """account_info 값을 고정 셀 좌표에 기입한다."""
    rows = [r[:] for r in rows]
    for (r, c), key in _FIXED_CELLS.items():
        rows[r][c] = str(account_info.get(key, ""))
    return rows


def _inject_dept(
    rows: list[list[str]],
    dept_주관: pd.DataFrame,
    dept_사용: pd.DataFrame,
) -> list[list[str]]:
    """
    부점귀속 섹션: 플레이스홀더 행을 부서별 데이터 행으로 교체하고
    총계 행에 합계를 기입한다.

    [변경] 고정 인덱스(19, 20) 제거 — 텍스트 스캔으로 위치를 동적으로 탐색한다.
    탐색 실패 시 ValueError를 발생시켜 template.csv 구조 변경을 명확히 감지한다.

    탐색 조건:
      - 헤더 행: col[2] == '부점명'
      - 총계 행: 헤더 이후에서 col[2]=='총계' AND col[5]=='총계'
      - 플레이스홀더: 헤더 행 + 1
    """
    rows = [r[:] for r in rows]

    # ── 부점명 헤더 행 스캔 → 플레이스홀더 = 헤더 + 1 ─────────────────────
    header_idx: int | None = None
    for i, row in enumerate(rows):
        if len(row) > 2 and row[2] == "부점명":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(
            "template.csv 구조 오류: 부점귀속 섹션 헤더 행을 찾을 수 없습니다.\n"
            "조건: col[2] == '부점명' 인 행이 존재해야 합니다.\n"
            "template.csv의 구조가 변경되었는지 확인하세요."
        )
    logger.debug("[_inject_dept] 부점명 헤더 행: rows[%d]", header_idx)
    placeholder_idx = header_idx + 1

    # ── 총계 행 스캔: 헤더 이후 col[2]='총계' AND col[5]='총계' ────────────
    # col[5]='총계' 조건으로 성격별 분류 섹션의 총계 행과 구별한다.
    total_idx: int | None = None
    for i in range(placeholder_idx, len(rows)):
        if len(rows[i]) > 5 and rows[i][2] == "총계" and rows[i][5] == "총계":
            total_idx = i
            break
    if total_idx is None:
        raise ValueError(
            "template.csv 구조 오류: 부점귀속 섹션 총계 행을 찾을 수 없습니다.\n"
            "조건: col[2]='총계' AND col[5]='총계' 인 행이 존재해야 합니다.\n"
            "template.csv의 구조가 변경되었는지 확인하세요."
        )
    logger.debug("[_inject_dept] 총계 행: rows[%d]", total_idx)

    # 데이터 행 빌드 (주관 cols 2-4, 사용 cols 5-7)
    n = max(len(dept_주관), len(dept_사용))
    dept_rows: list[list[str]] = []
    for i in range(n):
        new_row = [""] * _NUM_COLS
        if i < len(dept_주관):
            r = dept_주관.iloc[i]
            new_row[2] = str(r["부서명"])
            new_row[3] = _fmt(r["대상금액"])
            new_row[4] = _fmt(r["구성비(%)"])
        if i < len(dept_사용):
            r = dept_사용.iloc[i]
            new_row[5] = str(r["부서명"])
            new_row[6] = _fmt(r["대상금액"])
            new_row[7] = _fmt(r["구성비(%)"])
        dept_rows.append(new_row)

    # 플레이스홀더 제거 후 데이터 행 삽입, 총계 행 유지
    rows = rows[:placeholder_idx] + dept_rows + rows[total_idx:]

    # splice 후 총계 행 인덱스
    new_total = placeholder_idx + len(dept_rows)
    rows[new_total][3] = _fmt(dept_주관["대상금액"].sum())
    rows[new_total][6] = _fmt(dept_사용["대상금액"].sum())

    return rows


def _inject_direct_indirect(
    rows: list[list[str]], direct_indirect: dict
) -> list[list[str]]:
    """
    직접비·공통비 금액을 텍스트 스캔으로 찾아 기입한다.
    (dept splice 후 절대 행 인덱스가 바뀌므로 스캔 방식 사용)
    """
    rows = [r[:] for r in rows]
    for i, row in enumerate(rows):
        if row[1] == "직접비":
            rows[i][2] = _fmt(direct_indirect.get("직접비", 0))
        elif row[1] == "공통비":
            rows[i][2] = _fmt(direct_indirect.get("공통비", 0))
    return rows


def _inject_nature(
    rows: list[list[str]],
    nature: pd.DataFrame,
) -> list[list[str]]:
    """
    성격별 분류 섹션: 플레이스홀더 행을 부서별 행으로 교체하고
    총계 행에 합계를 기입한다.

    섹션 구조 (절대 인덱스는 dept splice 후 이동하므로 텍스트 스캔):
      ,,계약비,유지비,...        ← 성격 컬럼 헤더 행
      ,,,,,,,                   ← 빈 플레이스홀더 행
      ,총계,,,,,,               ← 총계 행 (col1="총계")
    """
    rows = [r[:] for r in rows]

    if nature.empty:
        return rows  # 성격별 데이터 없음 → 템플릿 구조 그대로 유지

    data_df = nature.iloc[:-1]        # 마지막 행(총계) 제외
    total_series = nature.iloc[-1]    # 총계 행

    # 성격별 데이터 행 빌드
    nature_rows: list[list[str]] = []
    for _, r in data_df.iterrows():
        new_row = [""] * _NUM_COLS
        new_row[1] = str(r["귀속_사용부서"])
        for j, col in enumerate(NATURE_COLS):
            new_row[2 + j] = _fmt(r[col])
        nature_rows.append(new_row)

    # ── "3-2" 섹션 이후에서 col[1]=="총계" 행을 탐색 ──────────────────────
    # [변경] len(rows)-5 폴백 제거 — 탐색 실패 시 ValueError로 명확히 감지한다.
    found_32 = False
    total_idx: int | None = None
    for i, row in enumerate(rows):
        if len(row) > 1 and "3-2" in row[1]:
            found_32 = True
        if found_32 and len(row) > 1 and row[1] == "총계":
            total_idx = i
            break

    if not found_32:
        raise ValueError(
            "template.csv 구조 오류: 성격별 분류 섹션 표시자를 찾을 수 없습니다.\n"
            "조건: col[1]에 '3-2'가 포함된 행이 존재해야 합니다.\n"
            "template.csv의 구조가 변경되었는지 확인하세요."
        )
    if total_idx is None:
        raise ValueError(
            "template.csv 구조 오류: 성격별 분류 섹션의 총계 행을 찾을 수 없습니다.\n"
            "조건: '3-2' 행 이후 col[1]=='총계' 인 행이 존재해야 합니다.\n"
            "template.csv의 구조가 변경되었는지 확인하세요."
        )
    logger.debug("[_inject_nature] 성격별 분류 총계 행: rows[%d]", total_idx)

    placeholder_idx = total_idx - 1  # 총계 바로 위 빈 행

    # 플레이스홀더 제거 후 데이터 행 삽입
    rows = rows[:placeholder_idx] + nature_rows + rows[total_idx:]

    # 새 총계 행 인덱스
    new_total = placeholder_idx + len(nature_rows)
    for j, col in enumerate(NATURE_COLS):
        rows[new_total][2 + j] = _fmt(total_series[col])

    return rows


def _fmt(value) -> str:
    """숫자를 문자열로 변환. 정수이면 소수점 제거, NaN은 빈 문자열."""
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isnan(f):
        return ""
    if f == int(f):
        return str(int(f))
    return f"{f:.2f}"


def _inject_classification_basis(
    rows: list[list[str]],
    classification_basis: dict,
) -> list[list[str]]:
    """
    4. 분류 근거 섹션에 CLASSIFICATION_BASIS 상수 텍스트를 기입한다.

    "4." + "근거" 포함 행 이후를 분류 근거 섹션으로 간주.
    col 1에 "직" + "공통" 포함 행 → col 2에 직공통비_근거 텍스트
    col 1에 "성격별" + "분류" 포함 행 → col 2에 성격별_근거 텍스트
    """
    rows = [r[:] for r in rows]
    in_근거 = False
    for i, row in enumerate(rows):
        if "4." in row[1] and "근거" in row[1]:
            in_근거 = True
        if not in_근거:
            continue
        if "직" in row[1] and "공통" in row[1]:
            rows[i][2] = classification_basis.get("직공통비_근거", "")
        elif "성격별" in row[1] and "분류" in row[1]:
            rows[i][2] = classification_basis.get("성격별_근거", "")
    return rows


def _write_csv(rows: list[list[str]], out_path: str) -> None:
    """utf-8-sig 인코딩으로 CSV 저장 (Excel 한글 깨짐 방지)."""
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerows(rows)
