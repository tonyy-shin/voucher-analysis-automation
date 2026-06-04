"""
company_theme.py
역할: 회사별 색상 테마를 표현하는 데이터클래스.
      ReportLab 의존성은 프로퍼티 내부에만 국한하여 임포트 비용을 최소화한다.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class CompanyTheme:
    """PDF 리포트에 적용될 색상 테마와 로고 정보를 담는다.

    Attributes:
        primary_hex:       섹션 라벨·헤더바·푸터바에 사용되는 주요 색상 (#RRGGBB).
        primary_light_hex: primary_hex 의 밝은 변형 (자동 산출, 매뉴얼 서브섹션용).
        logos:             [(파일명, 높이pt), ...] — Fallback 기본 로고 목록.
        logo_paths:        사용자가 지정한 로고 절대 경로 목록 (최대 3개, 빈 리스트 = 기본 로고 사용).
                           각 경로는 load 시 파일 존재 여부가 검증된다.
    """
    primary_hex: str
    primary_light_hex: str
    logos: list[tuple[str, int]]
    logo_paths: list[str] = field(default_factory=list)

    @property
    def primary_color(self):
        from reportlab.lib import colors
        return colors.HexColor(self.primary_hex)

    @property
    def primary_light_color(self):
        from reportlab.lib import colors
        return colors.HexColor(self.primary_light_hex)