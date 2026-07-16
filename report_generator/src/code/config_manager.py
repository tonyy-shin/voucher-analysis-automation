"""
config_manager.py
역할: column_config.yaml 로드/저장/생성 + CSV 폴더 대상 사전 검증.

저장 경로 규칙:
    - EXE 배포 환경 (PyInstaller frozen): sys.executable 옆 폴더
      → sys._MEIPASS 는 임시 폴더이므로 재시작 시 삭제됨, 절대 사용 금지
    - 개발 환경: 프로젝트 루트 (Path(__file__).parents[3])
"""
from __future__ import annotations

import copy
import sys
from pathlib import Path
from typing import Any

import yaml

# ── 설정 파일 경로 ─────────────────────────────────────────────────────────────
# PyInstaller frozen 환경에서는 sys.executable(EXE 파일) 옆에 저장한다.
# sys._MEIPASS 는 프로그램 종료 시 삭제되는 임시 폴더이므로 사용하면 안 된다.
if getattr(sys, "frozen", False):
    _CONFIG_DIR = Path(sys.executable).parent
else:
    _CONFIG_DIR = Path(__file__).parents[3]

_CONFIG_PATH = _CONFIG_DIR / "column_config.yaml"

# ── 기본 설정값 — data_loader.COLUMN_MAP 의 단일 원본 ─────────────────────────
_DEFAULT_CONFIG: dict[str, Any] = {
    "columns": {
        # 사업비정보.csv
        "원가요소":    "원가요소",
        "코스트센터":  "코스트센터",
        "주관부서":    "주관부서",   # 라인별 주관부서 (섹션2 주관부서 귀속 롤업 소스)
        # 계정정보.csv
        "계정번호":    "계정번호",
        "계정명":      "계정명",
        "계정그룹ID":  "계정그룹ID",
        "계정그룹명":  "계정그룹명",
        "대상정의":    "대상정의",
        "범위":        "사용 부서",
        "지급대상":    "비용 지급 범위",
        "산출기준":    "산출기준",
        "직간접구분":  "직간접구분",   # 계정정보.csv (3-1 직간접 분류 기준)
        # 부서정보.csv
        "cc_code":     "Cost Center Code",
        "팀":          "팀",
        # 출력.csv
        "출력전표":     "출력전표",
        "귀속_주관부서": "주관부서",
        "귀속_사용부서": "사용부서",
    },
    "nature_cols": [
        "계약체결비", "계약유지비", "손해조사비", "투자관리비", "간접비", "공통비",
    ],
    "col_team": "팀",
    "output_labels": {
        "코드": "출력전표",
        "주관": "주관부서",
        "사용": "사용부서",
    },
    # ── PDF 표시 문구 (pdf_exporter.py가 출력하는 모든 라벨/제목/설명) ──────────
    # 주의: 여기 값은 'PDF에 인쇄될 텍스트'이며, columns/nature_cols(=CSV 컬럼명)와
    #       별개입니다. 섹션 키(header, section1 ...)와 항목 키는 수정하지 마세요.
    "display_labels": {
        "header": "[ 사업비 전표 분석 ]",
        # 계정 테이블 컬럼은 동적 리스트 — csv_column 은 계정정보.csv 의 실제 헤더명.
        # width 는 컬럼 폭 비율 (기본 1.0, YAML 에서만 수정 — UI 미노출).
        "account_table": {
            "계정": "계정",
            "columns": [
                {"label": "계정코드",   "csv_column": "계정번호",   "width": 1.0},
                {"label": "계정명",     "csv_column": "계정명",     "width": 1.5},
                {"label": "사업비 코드", "csv_column": "계정그룹ID", "width": 1.0},
                {"label": "채널구분",   "csv_column": "계정그룹명", "width": 2.5},
            ],
        },
        # 섹션1 행도 동적 리스트 — csv_column 은 계정정보.csv 의 실제 헤더명.
        "section1": {
            "제목": "1. 대상정의",
            "rows": [
                {"label": "대상정의", "csv_column": "대상정의"},
                {"label": "범위",     "csv_column": "사용 부서"},
                {"label": "지급대상", "csv_column": "비용 지급 범위"},
                {"label": "산출기준", "csv_column": "산출기준"},
            ],
        },
        # 섹션2 귀속 그룹은 동적 리스트 — csv_column 은 집계 기준이 되는 df_enriched 컬럼명.
        #   주관부서 = 사업비정보.csv 라인별 주관부서 / 팀 = 코스트센터→부서정보 매핑.
        # 각 그룹은 해당 컬럼 기준 전체 롤업(총계 = 전체 대상금액)으로 계산된다.
        "section2": {
            "제목": "2. 부점귀속 현황",
            "귀속현황": "귀속현황",
            "groups": [
                {"label": "주관부서 귀속",     "csv_column": "주관부서"},
                {"label": "실제 사용부서 귀속", "csv_column": "팀"},
            ],
            "부점명": "부점명",
            "대상금액": "대상금액(단위: 원)",
            "구성비": "구성비(%)",
            "총계": "총계",
            "설명": ": 주관부서 귀속은 전표 발의 및 예산 통제 조직 기준이며, 실제 사용부서 귀속은 "
                    "업무 수행 및 비용 효익이 실질적으로 귀속되는 조직 기준.",
        },
        # 섹션3-1 분류 행은 동적 리스트 — keyword 는 부서정보.csv 직간접구분 값 부분일치.
        # 리스트 순서 = first-match-wins (위에서부터 먼저 일치한 분류로 집계).
        "section3_1": {
            "제목": "3-1. 직접비 공통비 분류 결과",
            "구분": "구분",
            "분류금액": "분류금액",
            "rows": [
                {"label": "직접비", "keyword": "직접"},
                {"label": "공통비", "keyword": "공통"},
            ],
        },
        "section3_2": {
            "제목": "3-2. 성격별 분류 결과",
            "귀속현황": "귀속현황",
            "합계": "합계",
            "설명": ": 성격 분류는 직접비로 구분된 금액을 계정 성격에 따라 분류하며, "
                    "공통비로 구분된 금액은 귀속 부서 성격에 따라 공통부서에서 직접 부서로 순차적 배부됨.",
            # 성격별 컬럼 표시명 — CSV 컬럼명(nature_cols)과 별개의 인쇄 텍스트.
            # CSV nature_cols와 같은 순서로 매칭되며, 부족분은 CSV 컬럼명으로 대체됩니다.
            "nature_cols": [
                "계약체결비", "계약유지비", "손해조사비", "투자관리비", "간접비", "공통비",
            ],
        },
        # 섹션4 분류 근거는 타입별 동적 행 시스템.
        #   - 직공통비   : cb의 직접비/공통비 근거로 렌더 (legacy 로직)
        #   - 성격별분류 : nat 절댓값 합 필터링으로 렌더 (legacy 로직)
        #   - custom     : csv_column 으로 지정한 출력.csv 컬럼의 첫 non-empty 값
        # rows 는 "문구 설정" 탭에서 추가/삭제/편집되며, 행 삭제 시 PDF에서도 제거된다.
        "section4": {
            "제목": "4. 분류 근거",
            "rows": [
                {"type": "직공통비",   "label": "직 • 공통비", "근거_csv_columns": ["직접비", "공통비"], "참조_csv_column": ""},
                {"type": "성격별분류", "label": "성격별 분류", "참조_csv_column": ""},
            ],
        },
    },
}

# ── CSV별 검증 대상 컬럼 키 목록 ──────────────────────────────────────────────
# JOIN 키 및 집계 필수 컬럼만 — 선택적 컬럼 누락은 오탐이므로 제외
_CSV_REQUIRED_COLS: dict[str, list[str]] = {
    "transaction": ["원가요소", "코스트센터"],
    "account":     ["계정번호", "직간접구분"],
    "ccm":         ["cc_code", "팀"],
    "output":      ["출력전표"],
}

# ── CSV 논리 키 → 표시용 한국어 레이블 ───────────────────────────────────────
CSV_LABELS: dict[str, str] = {
    "transaction": "사업비정보.csv",
    "account":     "계정정보.csv",
    "ccm":         "부서정보.csv",
    "output":      "출력.csv",
}


# ════════════════════════════════════════════════════════════════════════════
#  Public API
# ════════════════════════════════════════════════════════════════════════════

def default_display_labels() -> dict[str, Any]:
    """기본 PDF 표시 문구(display_labels)의 깊은 복사본을 반환한다.

    labels 인자가 주어지지 않은 호출부(예: pdf_exporter.export(labels=None))의
    안전한 폴백 값으로 사용된다.
    """
    return copy.deepcopy(_DEFAULT_CONFIG["display_labels"])


def load_config() -> dict[str, Any]:
    """column_config.yaml을 로드하고 기본값과 deep-merge하여 반환한다.

    파일이 없으면 주석 포함 YAML을 자동 생성한 뒤 기본값을 반환한다.
    파일이 있으나 파싱 실패 또는 비정상 구조 시 기본값으로 폴백한다.
    누락된 키는 기본값으로 채운다.
    """
    if not _CONFIG_PATH.exists():
        generate_default_config_file()
        return copy.deepcopy(_DEFAULT_CONFIG)

    try:
        raw = _CONFIG_PATH.read_text(encoding="utf-8")
        loaded = yaml.safe_load(raw)
        if not isinstance(loaded, dict):
            return copy.deepcopy(_DEFAULT_CONFIG)
        _discard_legacy_section4(loaded)
        _normalize_section4_rows(loaded)
        _normalize_account_table(loaded)
        _normalize_section1(loaded)
        _normalize_section2(loaded)
        _normalize_section3_1(loaded)
        return _deep_merge(_DEFAULT_CONFIG, loaded)
    except Exception:
        return copy.deepcopy(_DEFAULT_CONFIG)


def save_config(config: dict[str, Any]) -> bool:
    """설정을 column_config.yaml에 한국어 주석을 보존하여 저장한다.

    plain yaml.dump 대신 _render_annotated_yaml() 로 섹션별 주석을 재생성하여,
    저장 후에도 사람이 읽을 수 있는 안내 주석이 사라지지 않도록 한다.

    Returns:
        True: 저장 성공, False: OS 오류로 저장 실패 (쓰기 권한 부족 등)
    """
    try:
        merged = _deep_merge(_DEFAULT_CONFIG, config)
        _CONFIG_PATH.write_text(_render_annotated_yaml(merged), encoding="utf-8")
        return True
    except OSError:
        return False


def generate_default_config_file() -> None:
    """기본값으로 column_config.yaml을 생성한다. 파일이 이미 있으면 덮어쓰지 않는다."""
    if _CONFIG_PATH.exists():
        return
    try:
        _CONFIG_PATH.write_text(_render_annotated_yaml(_DEFAULT_CONFIG), encoding="utf-8")
    except OSError:
        pass


def validate_csv_columns(file_paths: dict, config: dict[str, Any]) -> list[str]:
    """선택된 4개 CSV 파일의 존재 여부 및 필수 컬럼 보유 여부를 검증한다.

    Args:
        file_paths: {"transaction":…, "account":…, "ccm":…, "output":…} 경로 dict
        config:     load_config()의 반환값

    Returns:
        문제 메시지 리스트 (사람이 읽을 수 있는 형태). 모두 정상이면 빈 리스트.
    """
    import pandas as pd  # noqa: PLC0415

    cols_cfg = config.get("columns", {})
    problems: list[str] = []

    for csv_key, required_col_keys in _CSV_REQUIRED_COLS.items():
        label = CSV_LABELS.get(csv_key, csv_key)
        path_str = file_paths.get(csv_key, "")
        if not path_str:
            problems.append(f"• {label}: 파일이 선택되지 않았습니다.")
            continue

        fpath = Path(path_str)
        disp = fpath.name or label

        if not fpath.exists():
            problems.append(f"• {label}({disp}) 파일을 찾을 수 없습니다.")
            continue

        try:
            head = pd.read_csv(fpath, dtype=str, nrows=0, encoding="utf-8-sig")
            actual_cols = set(head.columns)
        except Exception:
            try:
                head = pd.read_csv(fpath, dtype=str, nrows=0, encoding="cp949")
                actual_cols = set(head.columns)
            except Exception as exc:
                problems.append(f"• {label}({disp}) 파일을 읽을 수 없습니다: {exc}")
                continue

        missing = [
            cols_cfg.get(ck, ck)
            for ck in required_col_keys
            if cols_cfg.get(ck, ck) not in actual_cols
        ]
        if missing:
            problems.append(
                f"• {label}({disp}): 필수 컬럼 누락 → {', '.join(missing)}"
            )

    return problems


# ════════════════════════════════════════════════════════════════════════════
#  Internal helpers
# ════════════════════════════════════════════════════════════════════════════

def _discard_legacy_section4(loaded: dict) -> None:
    """구버전(플랫) section4 설정을 제거하여 기본 rows 구조로 폴백시킨다.

    구버전 section4 는 "직공통비"/"성격별분류"/"참조" 같은 플랫 키만 가지며 "rows" 키가
    없다. 이 경우 loaded에서 section4 를 통째로 삭제하여 _deep_merge가 기본값(_DEFAULT_CONFIG)의
    rows 리스트를 그대로 사용하도록 한다. 병합(_deep_merge) 전에 호출되어야 한다.
    """
    dl = loaded.get("display_labels")
    if not isinstance(dl, dict):
        return
    sec4 = dl.get("section4")
    if isinstance(sec4, dict) and "rows" not in sec4:
        del dl["section4"]


def _normalize_section4_rows(loaded: dict) -> None:
    """구버전 행의 "참조" 키를 폐기하여 새 "참조_csv_column" 스키마로 정규화한다.

    각 section4 행은 정적 텍스트 "참조" 대신 출력.csv 컬럼명("참조_csv_column")을 갖는다.
    로드된 행이 "참조" 키만 가지고 "참조_csv_column" 키가 없으면, "참조" 값을 조용히 버리고
    "참조_csv_column"은 빈 문자열("")로 둔다(_deep_merge 시 기본값 폴백). 병합 전에 호출한다.
    """
    dl = loaded.get("display_labels")
    if not isinstance(dl, dict):
        return
    sec4 = dl.get("section4")
    if not isinstance(sec4, dict):
        return
    rows = sec4.get("rows")
    if not isinstance(rows, list):
        return
    for row in rows:
        if isinstance(row, dict) and "참조" in row and "참조_csv_column" not in row:
            del row["참조"]


def _resolve_csv_col(loaded: dict, logical_key: str) -> str:
    """논리 컬럼 키를 실제 CSV 컬럼명으로 변환한다 (구버전 → 신 스키마 마이그레이션용).

    로드된 파일의 columns 맵이 우선한다 — 사용자가 예: 범위 컬럼을 리매핑했다면
    마이그레이션된 csv_column 에도 그 값이 반영되어야 한다. 없으면 기본값으로 폴백.
    """
    cols = loaded.get("columns")
    if isinstance(cols, dict):
        v = cols.get(logical_key)
        if isinstance(v, str) and v:
            return v
    return _DEFAULT_CONFIG["columns"].get(logical_key, logical_key)


def _str_or(value: Any, default: str) -> str:
    """비어 있지 않은 문자열이면 그대로, 아니면 default를 반환한다."""
    return value if isinstance(value, str) and value else default


def _sanitize_entry_list(sec: dict, list_key: str, required: tuple[str, ...]) -> None:
    """신 스키마 리스트의 비정상 항목을 제거한다 (수기 편집 방어).

    항목은 required 키를 모두 가진 dict 여야 한다. 값은 str 로 강제 변환한다.
    전부 비정상이면 키 자체를 삭제하여 _deep_merge 가 기본 리스트를 공급하게 한다.
    """
    entries = sec.get(list_key)
    if not isinstance(entries, list):
        del sec[list_key]
        return
    cleaned: list[dict] = []
    for e in entries:
        if not (isinstance(e, dict) and all(k in e for k in required)):
            continue
        for k in required:
            if not isinstance(e[k], str):
                e[k] = str(e[k])
        cleaned.append(e)
    if cleaned:
        sec[list_key] = cleaned
    else:
        del sec[list_key]


def _normalize_account_table(loaded: dict) -> None:
    """구버전(플랫) account_table 을 새 columns 리스트 스키마로 변환한다.

    섹션4의 폐기 방식(_discard_legacy_section4)과 달리 변환 방식을 쓴다 — 구버전 플랫
    dict 를 그대로 두면 _deep_merge 가 키 단위로 병합하여 잔여 키가 남고 사용자 라벨
    커스터마이징이 조용히 무시되기 때문이다. 병합(_deep_merge) 전에 호출한다.
    """
    dl = loaded.get("display_labels")
    if not isinstance(dl, dict):
        return
    sec = dl.get("account_table")
    if not isinstance(sec, dict):
        return
    if "columns" in sec:
        _sanitize_entry_list(sec, "columns", ("label", "csv_column"))
        for legacy_key in ("계정코드", "계정명", "사업비코드", "채널구분"):
            sec.pop(legacy_key, None)
        return
    # (구 라벨 키, 논리 컬럼 키, 기본 라벨, 폭 비율) — 기존 위치 매핑 그대로
    legacy = [
        ("계정코드",   "계정번호",   "계정코드",   1.0),
        ("계정명",     "계정명",     "계정명",     1.5),
        ("사업비코드", "계정그룹ID", "사업비 코드", 1.0),
        ("채널구분",   "계정그룹명", "채널구분",   2.5),
    ]
    dl["account_table"] = {
        "계정": _str_or(sec.get("계정"), "계정"),
        "columns": [
            {
                "label": _str_or(sec.get(old_key), default_label),
                "csv_column": _resolve_csv_col(loaded, logical),
                "width": width,
            }
            for old_key, logical, default_label, width in legacy
        ],
    }


def _normalize_section1(loaded: dict) -> None:
    """구버전(플랫) section1 을 새 rows 리스트 스키마로 변환한다. 병합 전에 호출한다."""
    dl = loaded.get("display_labels")
    if not isinstance(dl, dict):
        return
    sec = dl.get("section1")
    if not isinstance(sec, dict):
        return
    if "rows" in sec:
        _sanitize_entry_list(sec, "rows", ("label", "csv_column"))
        for legacy_key in ("대상정의", "범위", "지급대상", "산출기준"):
            sec.pop(legacy_key, None)
        return
    # (구 라벨 키 = 기본 라벨, 논리 컬럼 키) — 범위/지급대상은 columns 맵 경유로 해석
    legacy = [
        ("대상정의", "대상정의"),
        ("범위",     "범위"),
        ("지급대상", "지급대상"),
        ("산출기준", "산출기준"),
    ]
    dl["section1"] = {
        "제목": _str_or(sec.get("제목"), "1. 대상정의"),
        "rows": [
            {
                "label": _str_or(sec.get(old_key), old_key),
                "csv_column": _resolve_csv_col(loaded, logical),
            }
            for old_key, logical in legacy
        ],
    }


def _normalize_section2(loaded: dict) -> None:
    """구버전 section2 의 주관/사용 라벨 쌍을 새 groups 리스트 스키마로 변환한다.

    제목/귀속현황/부점명 등 공유 스칼라 키는 그대로 두고(_deep_merge 가 처리),
    groups 만 생성한 뒤 구버전 라벨 키를 제거한다. 병합 전에 호출한다.
    """
    dl = loaded.get("display_labels")
    if not isinstance(dl, dict):
        return
    sec = dl.get("section2")
    if not isinstance(sec, dict):
        return
    # 섹션2 집계 기준 컬럼: 주관부서 = 사업비정보 라인별 주관부서, 사용부서 = 코스트센터→팀.
    old_사용 = _resolve_csv_col(loaded, "귀속_사용부서")   # 구버전 출력.csv 사용부서 컬럼명
    team_col = _str_or(loaded.get("col_team"), _resolve_csv_col(loaded, "팀"))
    if "groups" in sec:
        _sanitize_entry_list(sec, "groups", ("label", "csv_column"))
        # 구버전 마이그레이션: '사용부서' 귀속 기준이 출력.csv 사용부서(원가요소 단위)에서
        # 코스트센터→부서정보 '팀' 롤업으로 바뀌었으므로 구 csv_column 을 팀 컬럼으로 이관.
        for g in sec.get("groups", []):
            if isinstance(g, dict) and g.get("csv_column") == old_사용:
                g["csv_column"] = team_col
        sec.pop("주관부서귀속", None)
        sec.pop("실제사용부서귀속", None)
        return
    sec["groups"] = [
        {
            "label": _str_or(sec.get("주관부서귀속"), "주관부서 귀속"),
            "csv_column": _resolve_csv_col(loaded, "주관부서"),
        },
        {
            "label": _str_or(sec.get("실제사용부서귀속"), "실제 사용부서 귀속"),
            "csv_column": team_col,
        },
    ]
    sec.pop("주관부서귀속", None)
    sec.pop("실제사용부서귀속", None)


def _normalize_section3_1(loaded: dict) -> None:
    """구버전 section3_1 의 직접비/공통비 라벨 쌍을 새 rows 리스트 스키마로 변환한다.

    keyword 는 기존 로직의 부분일치 문자열("직접"/"공통")을 그대로 부여하여
    분류 결과가 변하지 않도록 한다. 병합 전에 호출한다.
    """
    dl = loaded.get("display_labels")
    if not isinstance(dl, dict):
        return
    sec = dl.get("section3_1")
    if not isinstance(sec, dict):
        return
    if "rows" in sec:
        _sanitize_entry_list(sec, "rows", ("label", "keyword"))
        sec.pop("직접비", None)
        sec.pop("공통비", None)
        return
    sec["rows"] = [
        {"label": _str_or(sec.get("직접비"), "직접비"), "keyword": "직접"},
        {"label": _str_or(sec.get("공통비"), "공통비"), "keyword": "공통"},
    ]
    sec.pop("직접비", None)
    sec.pop("공통비", None)


def _deep_merge(base: dict, override: dict) -> dict:
    """base에 override를 재귀적으로 병합한다. override 값이 우선한다.

    리스트(예: section4.rows, nature_cols)는 dict가 아니므로 else 분기로 빠져 override가
    base를 통째로 대체한다. 인덱스 단위 병합은 하지 않는다.
    """
    result = copy.deepcopy(base)
    for key, val in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(val, dict):
            result[key] = _deep_merge(result[key], val)
        else:
            result[key] = val
    return result


# ════════════════════════════════════════════════════════════════════════════
#  YAML 파일 문자열 상수
# ════════════════════════════════════════════════════════════════════════════

_YAML_HEADER = """\
# =====================================================================
# 전표분석서 자동화 — 컬럼·파일명 설정 파일
# =====================================================================
# ⚠️  주의: CSV 파일의 컬럼명이나 파일명이 바뀌었을 때만 수정하세요.
#       수정 후 저장하면 다음 실행부터 새 값이 자동 적용됩니다.
# =====================================================================
# 형식: YAML (텍스트 파일, 메모장으로 열 수 있습니다)
#   각 항목:  논리명: "CSV에서 사용하는 실제 이름"
#   큰따옴표 안의 내용만 수정하고, 논리명(왼쪽)은 수정하지 마세요.
# =====================================================================

"""

def _yaml_dq(value: Any) -> str:
    """YAML 큰따옴표 스칼라로 직렬화한다 (백슬래시·큰따옴표만 이스케이프).

    라벨 값에는 ':' '%' '•' '(' 등이 들어갈 수 있어 항상 인용 처리한다.
    """
    s = str(value).replace("\\", "\\\\").replace('"', '\\"')
    return f'"{s}"'


def _render_annotated_yaml(config: dict[str, Any]) -> str:
    """config 값으로 주석이 포함된 column_config.yaml 본문 전체를 생성한다.

    columns / nature_cols / col_team / output_labels 블록은 섹션 주석과 함께
    config 값으로 재구성하고, display_labels 블록은 주석 헤더 + yaml.dump 로 덧붙인다.
    save_config / generate_default_config_file 양쪽이 공유한다.
    """
    cols = config.get("columns", {})
    out_lbl = config.get("output_labels", {})
    nat = config.get("nature_cols", [])

    def col(key: str) -> str:
        return _yaml_dq(cols.get(key, key))

    lines: list[str] = [_YAML_HEADER.rstrip("\n"), ""]

    lines += [
        "columns:",
        "  # ── 사업비정보.csv ──────────────────────────────────────────────────────",
        f"  원가요소: {col('원가요소')}",
        f"  코스트센터: {col('코스트센터')}",
        f"  주관부서: {col('주관부서')}   # 라인별 주관부서 (섹션2 주관부서 귀속 롤업 소스)",
        "",
        "  # ── 계정정보.csv ───────────────────────────────────────────────────────",
        f"  계정번호: {col('계정번호')}",
        f"  계정명: {col('계정명')}",
        f"  계정그룹ID: {col('계정그룹ID')}",
        f"  계정그룹명: {col('계정그룹명')}",
        f"  대상정의: {col('대상정의')}",
        f"  범위: {col('범위')}",
        f"  지급대상: {col('지급대상')}",
        f"  산출기준: {col('산출기준')}",
        "",
        "  # ── 부서정보.csv ───────────────────────────────────────────────────────",
        f"  cc_code: {col('cc_code')}",
        f"  직간접구분: {col('직간접구분')}",
        f"  팀: {col('팀')}",
        "",
        "  # ── 출력.csv ───────────────────────────────────────────────────────────",
        f"  출력전표: {col('출력전표')}",
        f"  귀속_주관부서: {col('귀속_주관부서')}",
        f"  귀속_사용부서: {col('귀속_사용부서')}",
        "",
        "nature_cols:",
        "  # ⚠️  성격별 분류 컬럼명(사업비정보.csv의 실제 숫자 값)이 바뀌면 아래 목록을 수정하세요.",
        "  # 순서도 보고서 테이블의 컬럼 순서에 영향을 줍니다.",
    ]
    lines += [f"  - {_yaml_dq(v)}" for v in nat]
    lines += [
        "",
        f"col_team: {_yaml_dq(config.get('col_team', '팀'))}   # 부서정보.csv의 팀 컬럼명. 이름이 바뀌면 여기를 수정하세요.",
        "",
        "output_labels:",
        "  # 출력.csv 안에서 사용하는 컬럼 레이블입니다.",
        f"  코드: {_yaml_dq(out_lbl.get('코드', '출력전표'))}",
        f"  주관: {_yaml_dq(out_lbl.get('주관', '주관부서'))}",
        f"  사용: {_yaml_dq(out_lbl.get('사용', '사용부서'))}",
        "",
        "display_labels:",
        "  # PDF에 인쇄되는 모든 표시 문구입니다. (CSV 컬럼명과 별개)",
        "  # 큰따옴표 안의 텍스트만 수정하고, 왼쪽 항목 키는 바꾸지 마세요.",
    ]

    dl_dump = yaml.dump(
        {"display_labels": config.get("display_labels", {})},
        allow_unicode=True,
        default_flow_style=False,
        indent=2,
        sort_keys=False,
    )
    # 최상위 'display_labels:' 줄은 위에서 주석과 함께 직접 출력했으므로 제거하고 본문만 덧붙인다.
    dl_body = "\n".join(dl_dump.splitlines()[1:])
    lines.append(dl_body)

    return "\n".join(lines).rstrip("\n") + "\n"
