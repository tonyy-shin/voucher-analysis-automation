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

# ── 기본 설정값 — data_loader.COLUMN_MAP + _DEFAULT_CSV_FILES 의 단일 원본 ─────
_DEFAULT_CONFIG: dict[str, Any] = {
    "csv_files": {
        "transaction": "사업비정보.csv",
        "account":     "계정정보.csv",
        "ccm":         "부서정보.csv",
        "output":      "출력.csv",
    },
    "columns": {
        # 사업비정보.csv
        "원가요소":    "원가요소",
        "코스트센터":  "코스트센터",
        # 계정정보.csv
        "계정번호":    "계정번호",
        "계정명":      "계정명",
        "계정그룹ID":  "계정그룹ID",
        "계정그룹명":  "계정그룹명",
        "대상정의":    "대상정의",
        "범위":        "사용 부서",
        "지급대상":    "비용 지급 범위",
        "산출기준":    "산출기준",
        # 부서정보.csv
        "cc_code":     "Cost Center Code",
        "직간접구분":  "직간접구분",
        "팀":          "팀",
        # 출력.csv
        "출력전표":     "출력전표",
        "귀속_주관부서": "주관부서",
        "귀속_사용부서": "사용부서",
    },
    "nature_cols": [
        "계약비", "유지비", "손해조사비", "투자관리비", "간접비", "공통비",
    ],
    "col_team": "팀",
    "output_labels": {
        "코드": "출력전표",
        "주관": "주관부서",
        "사용": "사용부서",
    },
}

# ── CSV별 검증 대상 컬럼 키 목록 ──────────────────────────────────────────────
# JOIN 키 및 집계 필수 컬럼만 — 선택적 컬럼 누락은 오탐이므로 제외
_CSV_REQUIRED_COLS: dict[str, list[str]] = {
    "transaction": ["원가요소", "코스트센터"],
    "account":     ["계정번호"],
    "ccm":         ["cc_code", "직간접구분", "팀"],
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
        return _deep_merge(_DEFAULT_CONFIG, loaded)
    except Exception:
        return copy.deepcopy(_DEFAULT_CONFIG)


def save_config(config: dict[str, Any]) -> bool:
    """설정을 column_config.yaml에 한국어 주석 헤더와 함께 저장한다.

    Returns:
        True: 저장 성공, False: OS 오류로 저장 실패 (쓰기 권한 부족 등)
    """
    try:
        body = yaml.dump(
            config,
            allow_unicode=True,
            default_flow_style=False,
            indent=2,
            sort_keys=False,
        )
        _CONFIG_PATH.write_text(_YAML_HEADER + body, encoding="utf-8")
        return True
    except OSError:
        return False


def generate_default_config_file() -> None:
    """기본값으로 column_config.yaml을 생성한다. 파일이 이미 있으면 덮어쓰지 않는다."""
    if _CONFIG_PATH.exists():
        return
    try:
        _CONFIG_PATH.write_text(_YAML_HEADER + _ANNOTATED_YAML_BODY, encoding="utf-8")
    except OSError:
        pass


def validate_csv_columns(csv_dir: str, config: dict[str, Any]) -> list[str]:
    """CSV 폴더의 4개 파일 존재 여부 및 필수 컬럼 보유 여부를 검증한다.

    Args:
        csv_dir: 4개 CSV가 들어있는 폴더 경로
        config:  load_config()의 반환값

    Returns:
        문제 메시지 리스트 (사람이 읽을 수 있는 형태). 모두 정상이면 빈 리스트.
    """
    import pandas as pd  # noqa: PLC0415

    base = Path(csv_dir)
    files_cfg = config.get("csv_files", _DEFAULT_CONFIG["csv_files"])
    cols_cfg = config.get("columns", {})
    problems: list[str] = []

    for csv_key, required_col_keys in _CSV_REQUIRED_COLS.items():
        fname = files_cfg.get(csv_key, CSV_LABELS.get(csv_key, csv_key))
        fpath = base / fname

        if not fpath.exists():
            problems.append(f"• {fname} 파일을 찾을 수 없습니다.")
            continue

        try:
            head = pd.read_csv(fpath, dtype=str, nrows=0, encoding="utf-8-sig")
            actual_cols = set(head.columns)
        except Exception:
            try:
                head = pd.read_csv(fpath, dtype=str, nrows=0, encoding="cp949")
                actual_cols = set(head.columns)
            except Exception as exc:
                problems.append(f"• {fname} 파일을 읽을 수 없습니다: {exc}")
                continue

        missing = [
            cols_cfg.get(ck, ck)
            for ck in required_col_keys
            if cols_cfg.get(ck, ck) not in actual_cols
        ]
        if missing:
            problems.append(
                f"• {fname}: 필수 컬럼 누락 → {', '.join(missing)}"
            )

    return problems


# ════════════════════════════════════════════════════════════════════════════
#  Internal helpers
# ════════════════════════════════════════════════════════════════════════════

def _deep_merge(base: dict, override: dict) -> dict:
    """base에 override를 재귀적으로 병합한다. override 값이 우선한다."""
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

_ANNOTATED_YAML_BODY = """\
csv_files:
  # 입력 폴더 안에 있어야 하는 4개 CSV 파일의 실제 파일명
  transaction: "사업비정보.csv"
  account: "계정정보.csv"
  ccm: "부서정보.csv"
  output: "출력.csv"

columns:
  # ── 사업비정보.csv ──────────────────────────────────────────────────────
  원가요소: "원가요소"
  코스트센터: "코스트센터"

  # ── 계정정보.csv ───────────────────────────────────────────────────────
  계정번호: "계정번호"
  계정명: "계정명"
  계정그룹ID: "계정그룹ID"
  계정그룹명: "계정그룹명"
  대상정의: "대상정의"
  범위: "사용 부서"
  지급대상: "비용 지급 범위"
  산출기준: "산출기준"

  # ── 부서정보.csv ───────────────────────────────────────────────────────
  cc_code: "Cost Center Code"
  직간접구분: "직간접구분"
  팀: "팀"

  # ── 출력.csv ───────────────────────────────────────────────────────────
  출력전표: "출력전표"
  귀속_주관부서: "주관부서"
  귀속_사용부서: "사용부서"

nature_cols:
  # ⚠️  성격별 분류 컬럼명(사업비정보.csv의 실제 숫자 값)이 바뀌면 아래 목록을 수정하세요.
  # 순서도 보고서 테이블의 컬럼 순서에 영향을 줍니다.
  - "계약비"
  - "유지비"
  - "손해조사비"
  - "투자관리비"
  - "간접비"
  - "공통비"

col_team: "팀"   # 부서정보.csv의 팀 컬럼명. 이름이 바뀌면 여기를 수정하세요.

output_labels:
  # 출력.csv 안에서 사용하는 컬럼 레이블입니다.
  코드: "출력전표"
  주관: "주관부서"
  사용: "사용부서"
"""
