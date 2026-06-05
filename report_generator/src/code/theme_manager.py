"""
theme_manager.py
역할: 사용자 색상 테마를 theme_config.json 에 저장·로드한다.

저장 경로 규칙:
    - EXE 배포 환경 (PyInstaller frozen): sys.executable 옆 폴더
      → sys._MEIPASS 는 임시 폴더이므로 재시작 시 삭제됨, 절대 사용 금지
    - 개발 환경: 프로젝트 루트 (Path(__file__).parents[3])

Fallback:
    파일 없음·파싱 실패·필수 키 누락·잘못된 hex 형식 — 모든 경우 GRAY_DEFAULT 반환.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from company_theme import CompanyTheme

# ── 설정 파일 경로 ─────────────────────────────────────────────────────────────
# PyInstaller frozen 환경에서는 sys.executable(EXE 파일) 옆에 저장한다.
# sys._MEIPASS 는 프로그램 종료 시 삭제되는 임시 폴더이므로 사용하면 안 된다.
if getattr(sys, "frozen", False):
    _BASE_DIR = Path(sys.executable).parent
else:
    _BASE_DIR = Path(__file__).parents[3]

_CONFIG_PATH = _BASE_DIR / "theme_config.json"

# ── 로고 상수 (색상과 무관하게 코드에서 고정 관리) ────────────────────────────
_LOGOS_DEFAULT: list[tuple[str, int]] = [
    ("우리금융그룹.png", 20),
    ("ABL.png", 26),
]

# ── 기본 회색 테마 — config 파일 없거나 손상 시 자동 적용 ──────────────────────
GRAY_DEFAULT = CompanyTheme(
    primary_hex="#808080",
    primary_light_hex="#C8C8C8",
    logos=_LOGOS_DEFAULT,
    logo_paths=[],
)

# ── hex 형식 검증 정규식 ──────────────────────────────────────────────────────
_HEX_RE = re.compile(r'^#[0-9A-Fa-f]{6}$')


# ── 공개 함수 ──────────────────────────────────────────────────────────────────

def load_theme() -> CompanyTheme:
    """theme_config.json 에서 테마를 로드한다.

    다음 모든 경우에 GRAY_DEFAULT 를 반환 (예외 발생 없음):
        - 파일이 없는 경우
        - JSON 파싱 실패 (손상)
        - 필수 키(primary_hex) 누락
        - 유효하지 않은 hex 색상 형식

    하위 호환성: 구버전 "logo_path" 단일 키가 있으면 첫 번째 원소로 자동 변환한다.
    각 경로는 파일 존재 여부를 검증하여 유효한 경로만 logo_paths 에 포함된다.
    """
    try:
        data = json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
        primary = _validate_hex(data["primary_hex"])

        if "logo_paths" in data:
            logo_paths = _validate_logo_paths(data["logo_paths"])
        elif "logo_path" in data:
            # 구버전 Migration: 단일 문자열 → 리스트
            migrated = _validate_logo_path(data["logo_path"])
            logo_paths = [migrated] if migrated else []
        else:
            logo_paths = []

        return CompanyTheme(
            primary_hex=primary,
            primary_light_hex=_derive_light(primary),
            logos=_LOGOS_DEFAULT,
            logo_paths=logo_paths,
        )
    except Exception:
        return GRAY_DEFAULT


def save_theme(theme: CompanyTheme) -> None:
    """현재 테마의 primary_hex 와 logo_paths 를 theme_config.json 에 저장한다.

    각 경로는 포워드 슬래시(/) 형식으로 정규화하여 저장한다 (JSON 파싱 안전).
    쓰기 권한 오류 등 OS 오류는 조용히 무시한다 (사용자에게 에러 노출 방지).
    """
    try:
        payload: dict = {"primary_hex": theme.primary_hex}
        if theme.logo_paths:
            payload["logo_paths"] = [Path(p).as_posix() for p in theme.logo_paths]
        _CONFIG_PATH.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        pass


# ── 내부 헬퍼 ─────────────────────────────────────────────────────────────────

def _validate_logo_path(path: str | None) -> str | None:
    """경로가 실제 파일로 존재하면 문자열로 반환, 아니면 None 반환 (Graceful Fallback)."""
    if path is None:
        return None
    try:
        return str(path) if Path(path).is_file() else None
    except Exception:
        return None


def _validate_logo_paths(raw: object) -> list[str]:
    """경로 리스트에서 실제 존재하는 파일 경로만 추려 반환한다 (최대 3개).

    raw 가 list 가 아니거나 각 항목이 유효하지 않은 경우 해당 항목만 제외한다.
    """
    if not isinstance(raw, list):
        return []
    result = []
    for item in raw[:3]:
        validated = _validate_logo_path(item if isinstance(item, str) else None)
        if validated:
            result.append(validated)
    return result


def _validate_hex(s: str) -> str:
    """'#RRGGBB' 형식인지 검증하고 대문자로 정규화한다. 실패 시 ValueError."""
    s = s.strip()
    if not _HEX_RE.match(s):
        raise ValueError(f"유효하지 않은 hex 색상: {s!r}")
    return s.upper()


def _derive_light(hex_str: str) -> str:
    """primary 색상을 흰색과 60% 블렌딩하여 밝은 변형 색상을 반환한다.

    계산식: 각 채널 = round(primary_channel * 0.4 + 255 * 0.6)
    int() 대신 round() 사용으로 반올림 오차를 방지한다.
    """
    h = hex_str.lstrip('#')
    r = round(int(h[0:2], 16) * 0.4 + 255 * 0.6)
    g = round(int(h[2:4], 16) * 0.4 + 255 * 0.6)
    b = round(int(h[4:6], 16) * 0.4 + 255 * 0.6)
    return f'#{r:02X}{g:02X}{b:02X}'