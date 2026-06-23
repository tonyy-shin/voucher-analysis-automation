"""
pdf_exporter.py
역할: ReportLab Platypus로 A4 세로 PDF를 직접 생성한다.
      template.pdf 레이아웃 스펙 기준으로 스타일 구성.
"""
from __future__ import annotations

import math
import os
import sys
from pathlib import Path
from tkinter import messagebox

import io

import pandas as pd
from PIL import Image as PILImage
from reportlab.lib import colors

from data_loader import NATURE_COLS, COL_SUBTOTAL, COL_TOTAL, COLUMN_MAP
from company_theme import CompanyTheme
from theme_manager import GRAY_DEFAULT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

# ── 경로 설정 ──────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    _BASE_DIR = Path(sys._MEIPASS)
else:
    _BASE_DIR = Path(__file__).parents[3]

# ── 폰트 경로 — 환경변수 > OS 자동 탐색 순으로 결정 ─────────────────────────
def _resolve_font_paths() -> tuple[Path, Path]:
    """한글 폰트(일반/볼드) 경로를 반환한다.

    탐색 순서:
        1. 환경변수 FONT_PATH (개별 경로 지정)
        2. OS별 기본 경로 자동 탐색
        3. 탐색 실패 시 빈 Path 반환 (호출부에서 Helvetica fallback)
    """
    import platform

    env = os.environ.get("FONT_PATH", "")
    if env:
        p = Path(env)
        return p, p

    os_name = platform.system()
    if os_name == "Windows":
        base = Path(r"C:\Windows\Fonts")
        return base / "malgun.ttf", base / "malgunbd.ttf"
    elif os_name == "Darwin":
        candidates = [
            Path("/System/Library/Fonts/AppleSDGothicNeo.ttc"),
            Path("/Library/Fonts/AppleGothic.ttf"),
        ]
        for c in candidates:
            if c.exists():
                return c, c
    else:  # Linux
        candidates = [
            Path("/usr/share/fonts/truetype/nanum/NanumGothic.ttf"),
            Path("/usr/share/fonts/nanum/NanumGothic.ttf"),
        ]
        for c in candidates:
            if c.exists():
                return c, c
    return Path(""), Path("")


_MALGUN_PATH, _MALGUNBD_PATH = _resolve_font_paths()

# ── 페이지 (A4 세로, 여백 15mm) ────────────────────────────────────────────────
_PAGE   = A4
_MARGIN = 15 * mm
_WIDTH  = _PAGE[0] - 2 * _MARGIN

# ── 색상 ──────────────────────────────────────────────────────────────────────
# 브랜드 무관 고정 색상만 유지; 포인트 색상은 CompanyTheme 을 통해 런타임에 주입된다.
_BLACK  = colors.black
_WHITE  = colors.white

# ── 헤더 테이블 컬럼 너비 상수 ───────────────────────────────────────────────
_LOGO_CELL_W       = _WIDTH * 0.25                     # 헤더 로고 셀 전체 너비
_COLS_HEADER_LOGO  = [_WIDTH * 0.125, _WIDTH * 0.125]  # 2-로고 서브테이블: 좌 | 우 (대칭)
_COLS_HEADER_LOGO_3 = [_WIDTH * 0.40 / 3] * 3          # 3-로고 서브테이블: 균등 3열 (슬롯당 ≈ 68pt)
_COLS_HEADER_TITLE = [_WIDTH * 0.75, _WIDTH * 0.25]   # 헤더 테이블: 타이틀 | 로고셀 (1~2 로고)
_COLS_HEADER_TITLE_3 = [_WIDTH * 0.60, _WIDTH * 0.40] # 헤더 테이블: 타이틀 | 로고셀 (3 로고)
_RENDER_DPI        = 150                               # 로고 사전 리사이즈 기준 DPI

# ── 섹션별 컬럼 너비 ──────────────────────────────────────────────────────────
# [C] Account 테이블: Account | Code | Name | GroupCode | GroupName (비율 1:1:1.5:1:2.5)
_COLS_ACCT = [r / 7.0 * _WIDTH for r in [1.0, 1.0, 1.5, 1.0, 2.5]]

# [D] 대상정의 2열: 라벨 22% | 데이터 78%
_COLS_2 = [0.22 * _WIDTH, 0.78 * _WIDTH]

# [E] 부점귀속 7열 테이블: 비율 1.2:1:1.2:0.8:1:1.2:0.8 (부점별 + 주관3 + 사용3)
_COLS_7 = [r / 7.2 * _WIDTH for r in [1.2, 1.0, 1.2, 0.8, 1.0, 1.2, 0.8]]

# [G] 3-2 성격별 8열 테이블: 부점별 + NATURE 6 + 합계(소계)
_COLS_8 = [r / 8.3 * _WIDTH for r in [1.3, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0]]

# [F] 직접비/공통비 2열: 구분 20% | 분류금액 80%
_COLS_DI = [0.20 * _WIDTH, 0.80 * _WIDTH]

# [H] 분류근거 3열: 라벨 18% | 내용 64% | 참조 18%
_COLS_BASIS = [0.18 * _WIDTH, 0.64 * _WIDTH, 0.18 * _WIDTH]


# ── 폰트 등록 ──────────────────────────────────────────────────────────────────

def _setup_fonts() -> tuple[str, str]:
    if 'MG' not in pdfmetrics._fonts:
        if _MALGUN_PATH and _MALGUN_PATH.exists():
            try:
                pdfmetrics.registerFont(TTFont('MG', str(_MALGUN_PATH)))
                bd = _MALGUNBD_PATH if (_MALGUNBD_PATH and _MALGUNBD_PATH.exists()) else _MALGUN_PATH
                pdfmetrics.registerFont(TTFont('MG-B', str(bd)))
                pdfmetrics.registerFontFamily('MG', normal='MG', bold='MG-B')
            except Exception:
                import warnings
                warnings.warn("한글 폰트 등록 실패 — 기본 폰트(Helvetica)로 fallback합니다.")
                return 'Helvetica', 'Helvetica-Bold'
        else:
            import warnings
            warnings.warn(
                f"한글 폰트를 찾을 수 없습니다 (경로: {_MALGUN_PATH}) — "
                "환경변수 FONT_PATH 를 설정하거나 지원 OS에서 실행하세요. "
                "기본 폰트(Helvetica)로 fallback합니다."
            )
            return 'Helvetica', 'Helvetica-Bold'
    if 'MG' in pdfmetrics._fonts:
        return 'MG', 'MG-B'
    return 'Helvetica', 'Helvetica-Bold'


# ── 공개 함수 ──────────────────────────────────────────────────────────────────

def export(results: dict, pdf_path: str, code: str,
           theme: CompanyTheme | None = None, logo_max_height: int = 48,
           logo_heights: list[int] | None = None) -> bool:
    _theme = theme or GRAY_DEFAULT
    _lh = list(logo_heights or [])
    _lh = (_lh + [_theme.logo_max_height] * (3 - len(_lh)))[:3]
    try:
        fn, fb = _setup_fonts()
        s = _styles(fn, fb, _theme)
        story = _build_story(results, s, fn, fb, logo_max_height=_theme.logo_max_height,
                             logo_heights=_lh)
        doc = SimpleDocTemplate(
            pdf_path, pagesize=_PAGE,
            leftMargin=_MARGIN, rightMargin=_MARGIN,
            topMargin=_MARGIN,  bottomMargin=_MARGIN,
        )
        doc.build(story)
        return True
    except Exception as exc:
        messagebox.showerror(
            "PDF 생성 실패",
            f"PDF 생성 중 오류가 발생했습니다.\n\n"
            f"파일: {os.path.basename(pdf_path)}\n사유: {exc}\n\n"
            "다른 출력 파일에는 영향이 없습니다.",
        )
        return False


# ── 스타일 팩토리 ──────────────────────────────────────────────────────────────

def _styles(fn: str, fb: str, theme: CompanyTheme) -> dict:
    s = {
        # 흰 배경 위 검정 글씨 타이틀
        'title_w': ParagraphStyle('title_w', fontName=fb, fontSize=12, leading=16, textColor=_BLACK),
        # 섹션 제목 (흰 배경, 검정 Bold 9pt, 좌측)
        'sec':     ParagraphStyle('sec',     fontName=fb, fontSize=9,   leading=12),
        # 본문
        'body':    ParagraphStyle('body',    fontName=fn, fontSize=8.5, leading=11),
        'bold':    ParagraphStyle('bold',    fontName=fb, fontSize=8.5, leading=11),
        'small':   ParagraphStyle('small',   fontName=fn, fontSize=7.5, leading=10),
        'right':   ParagraphStyle('right',   fontName=fn, fontSize=8.5, leading=11, alignment=2),
        'center':  ParagraphStyle('center',  fontName=fn, fontSize=8.5, leading=11, alignment=1),
        # 진회색 배경 컬럼 헤더 — 흰 글씨, 가운데 정렬, Bold
        'hdr_w':   ParagraphStyle('hdr_w',   fontName=fb, fontSize=8.5, leading=11, alignment=1, textColor=_WHITE),
        # 포인트/흰 배경 라벨 — 검정 글씨, 가운데 정렬, Bold
        'hdr':     ParagraphStyle('hdr',     fontName=fb, fontSize=8.5, leading=11, alignment=1),
    }
    # ── 테마 색상 주입 — s 경유로 전달함으로써 _tbl_*() 시그니처 변경 최소화 ──
    s['c_primary']  = theme.primary_color   # 섹션 라벨·헤더바·푸터바 색상
    s['logos']      = theme.logos           # PDF 헤더 Fallback 기본 로고 목록
    s['logo_paths'] = theme.logo_paths      # 사용자 지정 로고 경로 리스트 (최대 3개)
    return s


# ── 스토리 빌드 ────────────────────────────────────────────────────────────────

def _build_story(results: dict, s: dict, fn: str, fb: str, logo_max_height: int = 48,
                 logo_heights: list[int] | None = None) -> list:
    ai   = results.get('account_info', {})
    d주관 = results['dept_주관']
    d사용 = results['dept_사용']
    di   = results.get('direct_indirect', {})
    nat  = results['nature']
    cb   = results.get('classification_basis', {})

    gap = Spacer(1, 4)
    story: list = []
    story.append(_tbl_header_bar(s))
    story.append(_tbl_header(s, logo_max_height=logo_max_height, logo_heights=logo_heights));  story.append(gap)
    story.append(_tbl_account(ai, s));               story.append(gap)
    story.extend(_tbl_target(ai, s));                story.append(gap)
    story.extend(_tbl_dept(d주관, d사용, s, fb));    story.append(gap)
    story.extend(_tbl_nature_31(di, s));             story.append(Spacer(1, 2))
    story.extend(_tbl_nature_32(nat, s, fb));        story.append(gap)
    story.extend(_tbl_basis(cb, di, nat, s));        story.append(gap)
    story.append(_tbl_footer(s))
    return story


# ── 헬퍼 ──────────────────────────────────────────────────────────────────────

def _P(text, style) -> Paragraph:
    return Paragraph(str(text).replace('\n', '<br/><br/>'), style)


def _base_style() -> list:
    return [
        ('GRID',         (0, 0), (-1, -1), 0.5, _BLACK),
        ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING',  (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING',   (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 2),
    ]


# ── [A][B] 최상단 헤더 배너 + 안내문 ──────────────────────────────────────────

def _load_image_safe(path: str, max_height: float = 48, max_width: float = None) -> Image:
    """투명 PNG를 흰 배경 RGB로 합성하고 LANCZOS로 사전 리사이즈하여 ReportLab Image 반환.

    모든 투명도 모드('P', 'LA' 등)를 RGBA로 정규화 후 합성하여 흰 줄 아티팩트를 방지한다.
    _RENDER_DPI 기준으로 PIL LANCZOS 리사이즈 후 임베딩하여 화질을 PDF 렌더러에 의존하지 않는다.
    """
    img = PILImage.open(path)
    if img.mode not in ('RGB', 'RGBA'):
        img = img.convert('RGBA')
    if img.mode == 'RGBA':
        background = PILImage.new('RGB', img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[3])
        img = background

    iw, ih = img.size
    if max_width is not None:
        effective_w = max_width - 0.5
        scale = min(max_height / ih, effective_w / iw)
    else:
        scale = max_height / ih

    render_w = iw * scale
    render_h = ih * scale

    target_px_w = max(1, round(render_w * _RENDER_DPI / 72))
    target_px_h = max(1, round(render_h * _RENDER_DPI / 72))
    img = img.resize((target_px_w, target_px_h), PILImage.Resampling.LANCZOS)

    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return Image(buf, width=render_w, height=render_h)


def _tbl_header(s: dict, logo_max_height: int = 48,
                logo_heights: list[int] | None = None) -> Table:
    logo_imgs = []

    # Phase A — 유효 경로 수집 (is_file 검사 통과한 것만)
    valid_paths = [
        p for p in s.get('logo_paths', [])
        if p and Path(p).is_file()
    ]
    n_expected = len(valid_paths)

    # Phase B — 슬롯별 가용 너비 결정 (예상 로고 수 기준)
    if n_expected >= 3:
        slot_widths = list(_COLS_HEADER_LOGO_3)         # ≈ 68pt × 3
    elif n_expected == 2:
        slot_widths = list(_COLS_HEADER_LOGO)          # [90.82pt, 98.39pt]
    elif n_expected == 1:
        slot_widths = [_LOGO_CELL_W]                   # ≈ 189.21pt
    else:
        slot_widths = []

    # Phase C — 슬롯별 높이를 적용하여 fit-within 로딩
    # export()에서 길이 3으로 정규화된 값이 내려오므로 단순 슬라이싱으로 충분
    per_slot_heights = (logo_heights or [logo_max_height] * 3)[:n_expected]

    rendered = []
    for user_path, cell_w, slot_h in zip(valid_paths, slot_widths, per_slot_heights):
        try:
            rendered.append(_load_image_safe(user_path, max_height=slot_h, max_width=cell_w))
        except Exception:
            rendered.append(None)

    logo_imgs = [im for im in rendered if im is not None]

    # Fallback: 사용자 로고가 하나도 없을 때만 기본 로고 목록 사용
    if not logo_imgs:
        # Phase B와 동일한 슬롯 너비 로직을 fallback에도 적용하여 셀 초과 클리핑 방지
        fb_candidates = [
            (str(_BASE_DIR / 'images' / name), h)
            for name, h in s['logos']
            if (_BASE_DIR / 'images' / name).exists()
        ]
        n_fb = len(fb_candidates)
        if n_fb >= 3:
            fb_widths = [_LOGO_CELL_W / 3] * n_fb
        elif n_fb == 2:
            fb_widths = list(_COLS_HEADER_LOGO)
        elif n_fb == 1:
            fb_widths = [_LOGO_CELL_W]
        else:
            fb_widths = []

        for (path, h), cell_w in zip(fb_candidates, fb_widths):
            try:
                logo_imgs.append(_load_image_safe(path, max_height=h, max_width=cell_w))
            except Exception:
                pass

    # ── 로고 개수에 따른 동적 서브테이블 구성 ─────────────────────────────
    _logo_tbl_style = TableStyle([
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ('TOPPADDING',    (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ])
    n = len(logo_imgs)
    valid_heights = [im.drawHeight for im in logo_imgs if im is not None]
    common_h = max(valid_heights) if valid_heights else logo_max_height
    if n >= 3:
        lt = Table([logo_imgs[:3]], colWidths=_COLS_HEADER_LOGO_3, rowHeights=[common_h])
        lt.setStyle(_logo_tbl_style)
        logo_cell = lt
    elif n == 2:
        lt = Table([[logo_imgs[0], logo_imgs[1]]], colWidths=_COLS_HEADER_LOGO, rowHeights=[common_h])
        lt.setStyle(_logo_tbl_style)
        logo_cell = lt
    elif n == 1:
        logo_cell = logo_imgs[0]
    else:
        logo_cell = ''

    data = [
        [_P('[ 사업비 전표 분석 ]', s['title_w']), logo_cell],
    ]
    _h_col_w = _COLS_HEADER_TITLE_3 if n_expected >= 3 else _COLS_HEADER_TITLE
    t = Table(data, colWidths=_h_col_w)
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (0, 0), _WHITE),
        ('BACKGROUND',    (1, 0), (1, 0), _WHITE),
        ('ALIGN',         (1, 0), (1, 0), 'RIGHT'),
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING',   (0, 0), (-1, -1), 4),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
        ('TOPPADDING',    (0, 0), (-1, 0), 7),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 7),
        ('LEFTPADDING',   (1, 0), (1, 0), 0),
        ('RIGHTPADDING',  (1, 0), (1, 0), 0),
        ('TOPPADDING',    (1, 0), (1, 0), 0),
        ('BOTTOMPADDING', (1, 0), (1, 0), 0),
    ]))
    return t


# ── [C] Account 헤더 테이블 ────────────────────────────────────────────────────

def _tbl_account(ai: dict, s: dict) -> Table:
    data = [
        [_P('Account', s['hdr']),      _P('Account Code', s['hdr']),
         _P('Account Name', s['hdr']), _P('Group Code', s['hdr']),
         _P('Account Group Name', s['hdr'])],
        ['',
         _P(ai.get('계정번호', ''), s['body']),
         _P(ai.get('계정명', ''),   s['body']),
         _P(ai.get('계정그룹ID', ''), s['body']),
         _P(ai.get('계정그룹명', ''), s['body'])],
    ]
    cmds = _base_style() + [
        ('BACKGROUND', (0, 0), (-1, 0), s['c_primary']),        # 헤더행
        ('BACKGROUND', (0, 1), (-1, 1), colors.white),          # 데이터행: 흰색
        ('SPAN',       (0, 0), (0, 1)),                          # Account 셀 세로 병합
        ('BACKGROUND', (0, 0), (0, 1), s['c_primary']),          # SPAN 셀 전체 (덮어쓰기)
    ]
    return Table(data, colWidths=_COLS_ACCT, style=TableStyle(cmds))


# ── [D] 1. 대상정의 ────────────────────────────────────────────────────────────

def _tbl_target(ai: dict, s: dict) -> list:
    data = [
        [_P('대상정의', s['bold']), _P(ai.get('대상정의', ''), s['body'])],
        [_P('범위',     s['bold']), _P(ai.get('범위', ''),     s['body'])],
        [_P('지급대상', s['bold']), _P(ai.get('지급대상', ''), s['body'])],
        [_P('산출기준', s['bold']), _P(ai.get('산출기준', ''), s['body'])],
    ]
    cmds = _base_style() + [
        ('BACKGROUND', (0, 0), (0, 3), s['c_primary']),       # 라벨 셀
        ('BACKGROUND', (1, 0), (1, 3), colors.white),         # 입력 셀: 흰색
    ]
    return [_P('1. 대상정의', s['sec']), Table(data, colWidths=_COLS_2, style=TableStyle(cmds))]


# ── [E] 2. 부점귀속 현황 ───────────────────────────────────────────────────────

def _tbl_dept(d주관: pd.DataFrame, d사용: pd.DataFrame, s: dict, fb: str) -> list:
    n = max(len(d주관), len(d사용), 0)

    data = [
        # 행 0: 대분류 컬럼 헤더
        [_P('부점별<br/>(지역단별)<br/>귀속현황', s['hdr']),
         _P('주관부서 귀속', s['hdr']), None, None,
         _P('실제 사용부서 귀속', s['hdr']), None, None],
        # 행 1: 소분류 컬럼 헤더
        [None,
         _P('부점명', s['hdr']), _P('대상금액(단위: 원)', s['hdr']), _P('구성비(%)', s['hdr']),
         _P('부점명', s['hdr']), _P('대상금액(단위: 원)', s['hdr']), _P('구성비(%)', s['hdr'])],
    ]

    for i in range(n):
        row = ['']
        row += ([_P(str(d주관.iloc[i]['부서명']), s['body']),
                 _P(_fmt(d주관.iloc[i]['대상금액']), s['right']),
                 _P(_fmt(d주관.iloc[i]['구성비(%)']), s['right'])]
                if i < len(d주관) else ['', '', ''])
        row += ([_P(str(d사용.iloc[i]['부서명']), s['body']),
                 _P(_fmt(d사용.iloc[i]['대상금액']), s['right']),
                 _P(_fmt(d사용.iloc[i]['구성비(%)']), s['right'])]
                if i < len(d사용) else ['', '', ''])
        data.append(row)

    tot_idx = len(data)
    data.append([
        _P('총계', s['hdr']), '',
        _P(_fmt(d주관['대상금액'].sum()), s['right']), '',
        '', _P(_fmt(d사용['대상금액'].sum()), s['right']), '',
    ])

    cmds = _base_style() + [
        ('SPAN', (0, 0), (0, 1)),   # 부점별 rowspan=2
        ('SPAN', (1, 0), (3, 0)),   # 주관부서 귀속 colspan=3
        ('SPAN', (4, 0), (6, 0)),   # 실제 사용부서 귀속 colspan=3
        # 헤더 행 0-1: 포인트 색상
        ('BACKGROUND', (0, 0), (6, 1), s['c_primary']),
        # 총계 행: 라벨 흰색 / 나머지 노란색
        ('BACKGROUND', (0, tot_idx), (0, tot_idx), colors.white),
        ('BACKGROUND', (1, tot_idx), (6, tot_idx), colors.white),
        ('FONTNAME',   (0, tot_idx), (6, tot_idx), fb),
        ('ALIGN',      (0, 0), (0, 1), 'CENTER'),
    ]

    # 데이터 행: 노란색 (데이터가 있을 때만)
    if n > 0:
        cmds.append(('BACKGROUND', (0, 2), (6, tot_idx - 1), colors.white))

    title = _P('2. 부점귀속 현황', s['sec'])
    desc  = _P(': 주관부서 귀속은 전표 발의 및 예산 통제 조직 기준이며, 실제 사용부서 귀속은 '
               '업무 수행 및 비용 효익이 실질적으로 귀속되는 조직 기준.', s['small'])
    return [title, desc, Table(data, colWidths=_COLS_7, style=TableStyle(cmds))]


# ── [F] 3. 부점별 성격분류 + 3-1. 직접비/공통비 분류 결과 ────────────────────

def _tbl_nature_31(di: dict, s: dict) -> list:
    data = [
        [_P('구분', s['hdr']), _P('분류금액', s['hdr'])],
        [_P('직접비', s['bold']), _P(_fmt(di.get('직접비', 0)), s['right'])],
        [_P('공통비', s['bold']), _P(_fmt(di.get('공통비', 0)), s['right'])],
    ]
    cmds = _base_style() + [
        ('BACKGROUND', (0, 0), (1, 0), s['c_primary']),       # 헤더행
        ('BACKGROUND', (0, 1), (0, 2), s['c_primary']),       # 라벨(직접비/공통비)
        ('BACKGROUND', (1, 1), (1, 2), colors.white),         # 값 셀: 흰색
    ]
    return [
        _P('3. 부점별 성격분류', s['sec']),
        _P('3-1. 직접비 공통비 분류 결과', s['sec']),
        Table(data, colWidths=_COLS_DI, style=TableStyle(cmds)),
    ]


# ── [G] 3-2. 성격별 분류 결과 ──────────────────────────────────────────────────

def _tbl_nature_32(nat: pd.DataFrame, s: dict, fb: str) -> list:
    # 데이터 접근은 내부 컬럼명(COL_SUBTOTAL="소계"), 표시 헤더는 "합계"로 유지
    _display_cols = NATURE_COLS + [COL_SUBTOTAL]
    _header_labels = NATURE_COLS + ["합계"]
    _last_col = len(_display_cols)   # 컬럼 인덱스: 0=부점별, 1..N=데이터 → 마지막 인덱스

    data = [
        # 행 0: 컬럼 헤더 (부점별 + NATURE 6종 + 합계 = 8열)
        [_P('부점별<br/>(지역단별)<br/>귀속현황', s['hdr'])]
        + [_P(c, s['hdr']) for c in _header_labels],
    ]

    if not nat.empty:
        # 루프 전 성격 컬럼 일괄 숫자형 변환 — 매 행마다 변환 반복을 방지하고 str 혼입을 방어
        nat_numeric = nat.copy()
        for col in _display_cols:
            if col in nat_numeric.columns:
                nat_numeric[col] = pd.to_numeric(nat_numeric[col], errors='coerce').fillna(0)
        for _, r in nat_numeric.iloc[:-1].iterrows():
            data.append(
                [_P(str(r.get(COLUMN_MAP['귀속_사용부서'], '')), s['body'])]
                + [_P(_fmt(r.get(c, 0)), s['right']) for c in _display_cols]
            )
        tot = nat_numeric.iloc[-1]
        tot_row = [_P(COL_TOTAL, s['hdr'])] + [_P(_fmt(tot.get(c, 0)), s['right']) for c in _display_cols]
    else:
        tot_row = [_P(COL_TOTAL, s['hdr'])] + ['' for _ in _display_cols]

    tot_idx = len(data)
    data.append(tot_row)

    first_data = 1   # 데이터 행 시작 인덱스 (헤더 행 1개 제거 후)

    cmds = _base_style() + [
        # 헤더 행 0: 포인트 색상
        ('BACKGROUND', (0, 0), (_last_col, 0), s['c_primary']),
        ('ALIGN',      (0, 0), (0, 0), 'CENTER'),
        # 총계 행: 흰색
        ('BACKGROUND', (0, tot_idx), (_last_col, tot_idx), colors.white),
        ('FONTNAME',   (0, tot_idx), (_last_col, tot_idx), fb),
    ]

    # 데이터 행: 흰색
    last_data = tot_idx - 1
    if last_data >= first_data:
        cmds.append(('BACKGROUND', (0, first_data), (_last_col, last_data), colors.white))

    title = _P('3-2. 성격별 분류 결과', s['sec'])
    desc  = _P(': 성격 분류는 직접비 및 공통비로 구분된 금액을 대상으로 '
               '실제 사용부서의 업무 성격에 따라 분류한 결과.', s['small'])
    return [title, desc, Table(data, colWidths=_COLS_8, style=TableStyle(cmds))]


# ── [H] 4. 분류 근거 ───────────────────────────────────────────────────────────

def _tbl_basis(cb: dict, di: dict, nat: pd.DataFrame, s: dict) -> list:
    # 직·공통비: 출력.csv의 직접비/공통비 근거 중 텍스트가 채워진 항목만 "{구분}: {텍스트}" 형식으로 포함
    di_parts = []
    for k in ('직접비', '공통비'):
        t = cb.get(f'{k}_근거', '').strip()
        if t:
            di_parts.append(f'{k}: {t}')
    직공통비_text = '<br/><br/>'.join(di_parts)

    # 성격별 분류: 해당 성격 컬럼(총계 행 포함 전체)에 0이 아닌 값이 하나라도 있으면 근거 포함
    nat_parts = []
    if not nat.empty:
        for col in NATURE_COLS:
            if col not in nat.columns:
                continue
            abs_sum = pd.to_numeric(dept_rows[col], errors='coerce').abs().sum()
            txt = cb.get(f'{col}_근거', '').strip()
            if abs_sum > 0 and txt:
                nat_parts.append(f'{col}: {txt}')
    성격별_text = '<br/><br/>'.join(nat_parts)

    data = [
        [_P('직 • 공통비', s['bold']),
         _P(직공통비_text, s['body']),
         _P('총괄문서 참조', s['center'])],
        [_P('성격별 분류', s['bold']),
         _P(성격별_text, s['body']),
         _P('총괄문서 참조', s['center'])],
    ]
    cmds = _base_style() + [
        ('BACKGROUND', (0, 0), (0, 1), s['c_primary']),   # 좌측 라벨
        ('BACKGROUND', (1, 0), (1, 1), colors.white),     # 중간 입력 셀: 흰색
        ('BACKGROUND', (2, 0), (2, 1), colors.white),  # 우측 참조: 흰색
        ('ALIGN',      (0, 0), (0, 1), 'CENTER'),
        ('ALIGN',      (2, 0), (2, 1), 'CENTER'),
        ('ROWHEIGHT',  (0, 0), (0, 1), 30),            # 입력 행 여유 높이
    ]
    return [_P('4. 분류 근거', s['sec']), Table(data, colWidths=_COLS_BASIS, style=TableStyle(cmds))]


# ── [0] 상단 헤더 바 ───────────────────────────────────────────────────────────

def _tbl_header_bar(s: dict) -> Table:
    t = Table([['']], colWidths=[_WIDTH], rowHeights=[8])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), s['c_primary']),
        ('LINEABOVE',  (0, 0), (-1, -1), 0, s['c_primary']),
    ]))
    return t


# ── [I] 하단 푸터 바 ───────────────────────────────────────────────────────────

def _tbl_footer(s: dict) -> Table:
    t = Table([['']], colWidths=[_WIDTH], rowHeights=[8])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), s['c_primary']),
        ('GRID',       (0, 0), (-1, -1), 0, s['c_primary']),
    ]))
    return t


# ── 숫자 포매터 ────────────────────────────────────────────────────────────────

def _fmt(value) -> str:
    """천 단위 쉼표 포함 문자열 변환. NaN은 빈 문자열."""
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value) if value is not None else ''
    if math.isnan(f):
        return ''
    if f == int(f):
        return f'{int(f):,}'       # 예: 1000000 → "1,000,000"
    return f'{f:,.2f}'             # 예: 33.33 → "33.33"
