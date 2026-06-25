"""
generate_manual.py
역할: 지급사업비 분석 자동화 프로그램 사용설명서(v2.0.0) PDF를 생성한다.

기존 pdf_exporter.py의 PDF 생성 패턴(ReportLab Platypus, A4 세로, NanumGothic TTF 등록,
상/하단 색상 바)을 재사용하되, EY Korea 브랜드 색상과 20mm 여백을 적용한다.

폰트: NanumGothic (Regular / Bold). 탐색 순서 — 환경변수 FONT_PATH > ~/.fonts >
      /usr/share/fonts/truetype/nanum. 모두 실패 시 Helvetica 폴백(한글 깨짐 경고).

실행:  python3 generate_manual.py
출력:  release/사용설명서_v2.0.0.pdf
"""
from __future__ import annotations

import os
import sys
import warnings
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Frame, Image, NextPageTemplate, PageBreak, PageTemplate,
    Paragraph, Spacer, Table, TableStyle,
)

from PIL import Image as PILImage

# ── 출력 경로 ─────────────────────────────────────────────────────────────────
_HERE = Path(__file__).resolve().parent
_OUT_PDF = _HERE / "사용설명서_v2.1.0.pdf"
_SHOTS = _HERE / "screenshots"          # capture_screenshots.py 출력 폴더

# ── EY Korea 브랜드 색상 ──────────────────────────────────────────────────────
EY_YELLOW = colors.HexColor("#FFE600")   # 메인 (포인트)
EY_DARK   = colors.HexColor("#2E2E38")   # 텍스트/헤더 바
EY_GRAY   = colors.HexColor("#747480")   # 서브 (표 헤더, 보조선)
_WHITE    = colors.white
_LIGHTBG  = colors.HexColor("#F4F4F6")   # 표 라벨 셀 배경(연회색)

# ── 경고 박스 색상 ────────────────────────────────────────────────────────────
_WARN_BG     = colors.HexColor("#FFF8E1")   # 연한 앰버 배경
_WARN_BORDER = colors.HexColor("#E6A800")   # 경고 박스 테두리
_WARN_RED    = colors.HexColor("#C0392B")   # 경고 제목(앱 경고 문구와 동일 톤)

# ── 페이지 (A4 세로, 여백 20mm) ───────────────────────────────────────────────
_PAGE   = A4
_MARGIN = 20 * mm
_WIDTH  = _PAGE[0] - 2 * _MARGIN


# ════════════════════════════════════════════════════════════════════════════
#  폰트 등록 (pdf_exporter._setup_fonts 패턴 차용)
# ════════════════════════════════════════════════════════════════════════════

def _resolve_font_paths() -> tuple[Path | None, Path | None]:
    env = os.environ.get("FONT_PATH", "")
    if env and Path(env).exists():
        p = Path(env)
        return p, p
    home = Path.home()
    _win_fonts = Path(r"C:\Windows\Fonts")
    regular_candidates = [
        home / ".fonts" / "NanumGothic-Regular.ttf",
        home / ".local/share/fonts" / "NanumGothic-Regular.ttf",
        Path("/usr/share/fonts/truetype/nanum/NanumGothic.ttf"),
        Path("/usr/share/fonts/nanum/NanumGothic.ttf"),
        _win_fonts / "malgun.ttf",          # Windows 한글(맑은 고딕) 폴백
    ]
    bold_candidates = [
        home / ".fonts" / "NanumGothic-Bold.ttf",
        home / ".local/share/fonts" / "NanumGothic-Bold.ttf",
        Path("/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"),
        Path("/usr/share/fonts/nanum/NanumGothicBold.ttf"),
        _win_fonts / "malgunbd.ttf",        # Windows 한글 볼드 폴백
    ]
    reg = next((c for c in regular_candidates if c.exists()), None)
    bold = next((c for c in bold_candidates if c.exists()), None)
    return reg, bold


def setup_fonts() -> tuple[str, str]:
    """NanumGothic을 NG / NG-B 로 등록. 실패 시 Helvetica 폴백."""
    reg, bold = _resolve_font_paths()
    if reg is None:
        warnings.warn(
            "NanumGothic 폰트를 찾을 수 없습니다 — Helvetica로 폴백합니다(한글 깨짐). "
            "FONT_PATH 환경변수를 설정하거나 ~/.fonts에 NanumGothic-Regular.ttf를 두세요."
        )
        return "Helvetica", "Helvetica-Bold"
    try:
        pdfmetrics.registerFont(TTFont("NG", str(reg)))
        bold_path = bold if bold is not None else reg
        pdfmetrics.registerFont(TTFont("NG-B", str(bold_path)))
        pdfmetrics.registerFontFamily("NG", normal="NG", bold="NG-B",
                                      italic="NG", boldItalic="NG-B")
        print(f"[font] NanumGothic 등록 성공: regular={reg}, bold={bold_path}")
        return "NG", "NG-B"
    except Exception as exc:  # noqa: BLE001
        warnings.warn(f"폰트 등록 실패({exc}) — Helvetica로 폴백합니다.")
        return "Helvetica", "Helvetica-Bold"


# ════════════════════════════════════════════════════════════════════════════
#  스타일
# ════════════════════════════════════════════════════════════════════════════

def build_styles(fn: str, fb: str) -> dict:
    return {
        "cover_title": ParagraphStyle("cover_title", fontName=fb, fontSize=24,
                                       leading=32, textColor=EY_DARK, alignment=TA_LEFT),
        "cover_sub":   ParagraphStyle("cover_sub", fontName=fn, fontSize=12,
                                      leading=18, textColor=EY_GRAY, alignment=TA_LEFT),
        "cover_logo":  ParagraphStyle("cover_logo", fontName=fb, fontSize=18,
                                      leading=22, textColor=EY_DARK, alignment=TA_LEFT),
        "h1":   ParagraphStyle("h1", fontName=fb, fontSize=14, leading=18, textColor=_WHITE),
        "h2":   ParagraphStyle("h2", fontName=fb, fontSize=11, leading=15, textColor=EY_DARK),
        "body": ParagraphStyle("body", fontName=fn, fontSize=9.5, leading=15, textColor=EY_DARK),
        "bullet": ParagraphStyle("bullet", fontName=fn, fontSize=9.5, leading=14,
                                 textColor=EY_DARK, leftIndent=12, bulletIndent=2),
        "th":   ParagraphStyle("th", fontName=fb, fontSize=9, leading=12,
                               textColor=_WHITE, alignment=TA_CENTER),
        "td":   ParagraphStyle("td", fontName=fn, fontSize=9, leading=12, textColor=EY_DARK),
        "td_c": ParagraphStyle("td_c", fontName=fn, fontSize=9, leading=12,
                               textColor=EY_DARK, alignment=TA_CENTER),
        "td_b": ParagraphStyle("td_b", fontName=fb, fontSize=9, leading=12, textColor=EY_DARK),
        "faq_q": ParagraphStyle("faq_q", fontName=fb, fontSize=10, leading=14, textColor=EY_DARK),
        "faq_a": ParagraphStyle("faq_a", fontName=fn, fontSize=9.5, leading=15,
                                textColor=EY_DARK, leftIndent=10),
        "note": ParagraphStyle("note", fontName=fn, fontSize=8.5, leading=12, textColor=EY_GRAY),
        "caption": ParagraphStyle("caption", fontName=fn, fontSize=8, leading=11,
                                  textColor=EY_GRAY, alignment=TA_CENTER),
        "warn_title": ParagraphStyle("warn_title", fontName=fb, fontSize=10, leading=14,
                                     textColor=_WARN_RED),
        "warn_body": ParagraphStyle("warn_body", fontName=fn, fontSize=9, leading=14,
                                    textColor=EY_DARK),
    }


# ════════════════════════════════════════════════════════════════════════════
#  플로어블 헬퍼
# ════════════════════════════════════════════════════════════════════════════

def section_title(num: str, text: str, s: dict) -> Table:
    """Yellow 좌측 액센트 + Dark 바 + 흰 글씨 섹션 제목."""
    label = Paragraph(f"{num}. {text}", s["h1"])
    t = Table([["", label]], colWidths=[5, _WIDTH - 5])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), EY_YELLOW),
        ("BACKGROUND", (1, 0), (1, 0), EY_DARK),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (1, 0), (1, 0), 10),
        ("TOPPADDING",  (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (0, 0), 0),
        ("RIGHTPADDING", (0, 0), (0, 0), 0),
    ]))
    return t


def subhead(text: str, s: dict) -> Paragraph:
    return Paragraph(f'<font color="#FFB300">▍</font> {text}', s["h2"])


def body(text: str, s: dict) -> Paragraph:
    return Paragraph(text, s["body"])


def bullets(items: list[str], s: dict) -> list:
    out = []
    for it in items:
        out.append(Paragraph(it, s["bullet"], bulletText="•"))
    return out


def info_table(rows: list[list], s: dict, col_widths: list[float],
               header: list[str] | None = None) -> Table:
    data = []
    if header:
        data.append([Paragraph(h, s["th"]) for h in header])
    for r in rows:
        data.append(r)
    t = Table(data, colWidths=col_widths, hAlign="LEFT")
    style = [
        ("GRID",        (0, 0), (-1, -1), 0.5, EY_GRAY),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",(0, 0), (-1, -1), 6),
        ("TOPPADDING",  (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
    ]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), EY_GRAY))
    t.setStyle(TableStyle(style))
    return t


def label_table(pairs: list[tuple[str, str]], s: dict,
                label_w: float = 0.28) -> Table:
    """좌측 라벨(연회색) + 우측 값 2열 테이블."""
    data = [[Paragraph(k, s["td_b"]), Paragraph(v, s["td"])] for k, v in pairs]
    t = Table(data, colWidths=[label_w * _WIDTH, (1 - label_w) * _WIDTH], hAlign="LEFT")
    t.setStyle(TableStyle([
        ("GRID",        (0, 0), (-1, -1), 0.5, EY_GRAY),
        ("BACKGROUND",  (0, 0), (0, -1), _LIGHTBG),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",(0, 0), (-1, -1), 6),
        ("TOPPADDING",  (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
    ]))
    return t


def gap(h: float = 6):
    return Spacer(1, h)


def screenshot(name: str, caption: str, s: dict,
               max_w: float = _WIDTH * 0.82, max_h: float = 92 * mm) -> list:
    """release/screenshots/<name> 을 비율 유지로 축소하여 테두리 박스 + 캡션으로 반환한다.

    파일이 없으면 빈 리스트를 반환하여, 스크린샷이 아직 생성되지 않아도 매뉴얼은
    정상 빌드되도록 한다(capture_screenshots.py 미실행 시 안전).
    """
    p = _SHOTS / name
    if not p.exists():
        return []
    try:
        iw, ih = PILImage.open(p).size
    except Exception:  # noqa: BLE001
        return []
    scale = min(max_w / iw, max_h / ih)
    w, h = iw * scale, ih * scale
    img = Image(str(p), width=w, height=h)
    box = Table([[img]], colWidths=[w], hAlign="CENTER")
    box.setStyle(TableStyle([
        ("BOX",          (0, 0), (-1, -1), 0.5, EY_GRAY),
        ("LEFTPADDING",  (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING",   (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 2),
    ]))
    out = [box]
    if caption:
        out.append(Paragraph(caption, s["caption"]))
    return out


def warn_box(title: str, body: str, s: dict) -> Table:
    """노란 액센트 바 + 앰버 배경의 강조 경고 박스 (제목 빨강, 본문 다크)."""
    cell = [Paragraph(f"[주의] {title}", s["warn_title"])]
    if body:
        cell.append(Spacer(1, 3))
        cell.append(Paragraph(body, s["warn_body"]))
    t = Table([["", cell]], colWidths=[5, _WIDTH - 5])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (0, 0), EY_YELLOW),   # 좌측 액센트 바
        ("BACKGROUND",   (1, 0), (1, 0), _WARN_BG),     # 본문 배경
        ("BOX",          (1, 0), (1, 0), 0.6, _WARN_BORDER),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",  (0, 0), (0, 0), 0),
        ("RIGHTPADDING", (0, 0), (0, 0), 0),
        ("LEFTPADDING",  (1, 0), (1, 0), 10),
        ("RIGHTPADDING", (1, 0), (1, 0), 10),
        ("TOPPADDING",   (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 8),
    ]))
    return t


# ════════════════════════════════════════════════════════════════════════════
#  페이지 데코레이션 (상/하단 바 — pdf_exporter 헤더/푸터 바 패턴)
# ════════════════════════════════════════════════════════════════════════════

_FOOTER_LEFT = "지급사업비 분석 자동화 프로그램 사용설명서 v2.1.0"
_FOOTER_RIGHT = "© EY Korea"


def _draw_frame_bars(canvas, doc, cover: bool = False):
    canvas.saveState()
    w, h = _PAGE
    if cover:
        # 표지: 상단 큰 Dark 블록 + Yellow 액센트 띠
        canvas.setFillColor(EY_DARK)
        canvas.rect(0, h - 70 * mm, w, 70 * mm, stroke=0, fill=1)
        canvas.setFillColor(EY_YELLOW)
        canvas.rect(0, h - 74 * mm, w, 4 * mm, stroke=0, fill=1)
    else:
        # 본문: 상단 얇은 Yellow 바 + 그 아래 Dark 라인
        canvas.setFillColor(EY_YELLOW)
        canvas.rect(0, h - 6 * mm, w, 6 * mm, stroke=0, fill=1)
        canvas.setFillColor(EY_DARK)
        canvas.rect(0, h - 7 * mm, w, 1 * mm, stroke=0, fill=1)
    # 하단 푸터(공통)
    canvas.setStrokeColor(EY_GRAY)
    canvas.setLineWidth(0.5)
    canvas.line(_MARGIN, 12 * mm, w - _MARGIN, 12 * mm)
    canvas.setFont("NG" if "NG" in pdfmetrics._fonts else "Helvetica", 7.5)
    canvas.setFillColor(EY_GRAY)
    canvas.drawString(_MARGIN, 8 * mm, _FOOTER_LEFT)
    canvas.drawRightString(w - _MARGIN, 8 * mm, f"{_FOOTER_RIGHT}  |  {doc.page}")
    canvas.restoreState()


def on_cover(canvas, doc):
    _draw_frame_bars(canvas, doc, cover=True)


def on_page(canvas, doc):
    _draw_frame_bars(canvas, doc, cover=False)


# ════════════════════════════════════════════════════════════════════════════
#  본문 콘텐츠
# ════════════════════════════════════════════════════════════════════════════

def build_story(s: dict) -> list:
    story: list = []

    # ── 표지 ──────────────────────────────────────────────────────────────
    story.append(Spacer(1, 78 * mm))   # Dark 헤더 블록 아래로 내림
    story.append(Paragraph("지급사업비 분석 자동화 프로그램", s["cover_title"]))
    story.append(Paragraph("사용설명서", s["cover_title"]))
    story.append(gap(10))
    story.append(Paragraph("Business Expense Statement Analysis — User Manual", s["cover_sub"]))
    story.append(gap(28))
    cover_meta = [
        [Paragraph("버전", s["td_b"]), Paragraph("v2.1.0", s["td"])],
        [Paragraph("배포일", s["td_b"]), Paragraph("2026년 06월", s["td"])],
        [Paragraph("운영 체제", s["td_b"]), Paragraph("Windows 전용", s["td"])],
        [Paragraph("입력 형식", s["td_b"]), Paragraph("CSV 4개 파일 (UTF-8 / CP949)", s["td"])],
        [Paragraph("출력 형식", s["td_b"]), Paragraph("PDF (A4 세로)", s["td"])],
    ]
    ct = Table(cover_meta, colWidths=[35 * mm, 90 * mm], hAlign="LEFT")
    ct.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, EY_GRAY),
        ("BACKGROUND", (0, 0), (0, -1), _LIGHTBG),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(ct)

    story.append(NextPageTemplate("body"))
    story.append(PageBreak())

    # ── 1. 프로그램 개요 ─────────────────────────────────────────────────
    story.append(section_title("1", "프로그램 개요", s))
    story.append(gap(6))
    story.append(body(
        "본 프로그램은 사업비 전표 데이터가 담긴 <b>4개의 CSV 파일</b>(사업비정보·계정정보·"
        "부서정보·출력)을 입력받아, <b>출력전표 코드별로</b> 표준 양식의 전표 분석서(PDF)를 "
        "자동 생성합니다. 모든 집계는 Python(pandas) 메모리 내에서 수행되며 결과는 PDF로만 "
        "출력됩니다.", s))
    story.append(gap(6))
    story.append(subhead("주요 기능", s))
    story.extend(bullets([
        "<b>출력전표 코드별 일괄 처리</b> — 출력.csv에 등록된 코드마다 PDF 1개를 생성합니다.",
        "<b>3중 LEFT JOIN</b> — 사업비정보 ↔ 부서정보(코스트센터) ↔ 계정정보(원가요소) ↔ 출력(출력전표).",
        "<b>부점귀속 현황</b> — 출력.csv의 주관부서·사용부서 기준으로 금액을 귀속 분류합니다.",
        "<b>직접비/공통비 분류</b> — 부서정보.csv의 직간접구분('직접'/'공통') 기준으로 집계합니다.",
        "<b>성격별 분류(6종)</b> — 계약체결비·계약유지비·손해조사비·투자관리비·간접비·공통비를 부점별로 교차 집계합니다.",
        "<b>색상·로고 설정</b> — 포인트 색상 및 로고(최대 3슬롯, 슬롯별 높이 조절·실시간 미리보기).",
        "<b>문구 설정 탭</b> — 테마 대화상자의 [문구 설정] 탭에서 PDF에 인쇄되는 모든 표시 문구(제목·라벨·헤더)를 직접 편집할 수 있습니다.",
        "<b>성격 컬럼 추가/삭제</b> — [문구 설정] 탭에서 성격 분류 컬럼을 추가·삭제하며, 입력한 이름이 CSV 컬럼명과 PDF 표시명에 동시에 적용됩니다.",
        "<b>미등록 부서 사전 감지</b> — 출력.csv에 없는 부서를 처리 전에 경고합니다.",
        "<b>미등록 성격 컬럼 경고</b> — 설정한 성격 컬럼이 사업비정보.csv에 없으면 최종 요약 창에 경고로 표시됩니다.",
        "<b>유연한 컬럼 설정</b> — column_config.yaml로 컬럼명을 외부에서 변경할 수 있습니다.",
    ], s))

    story.append(gap(10))

    # ── 2. 시작 전 준비사항 ──────────────────────────────────────────────
    story.append(section_title("2", "시작 전 준비사항", s))
    story.append(gap(6))
    story.append(warn_box(
        "CSV 파일을 닫은 뒤 실행하세요",
        "입력 CSV 4개 파일 중 하나라도 Excel 등 다른 프로그램에서 <b>열려 있는 상태</b>로 "
        "프로그램을 실행하면 파일 접근 오류가 발생합니다. 실행 전 모든 CSV 파일을 닫아 주세요.", s))
    story.append(gap(6))
    story.append(subhead("2-1. 필요한 입력 파일 (CSV 4개)", s))
    story.append(body(
        "프로그램 실행 시 아래 <b>4개의 CSV 파일</b>을 각각 선택해야 합니다. 동봉된 "
        "<b>templates/</b> 폴더의 양식 파일을 복사하여 데이터를 채워 사용하시기 바랍니다.", s))
    story.append(gap(4))

    # 파일 개요 표
    file_rows = [
        [Paragraph("사업비정보.csv", s["td_b"]), Paragraph("전표 원본 데이터 + 성격별 금액·합계", s["td"])],
        [Paragraph("계정정보.csv", s["td_b"]), Paragraph("계정번호별 분류 기준 정보(계정명·대상정의 등)", s["td"])],
        [Paragraph("부서정보.csv", s["td_b"]), Paragraph("코스트센터별 팀·직간접구분 정보", s["td"])],
        [Paragraph("출력.csv", s["td_b"]), Paragraph("처리할 출력전표 코드 + 귀속 부서 + 분류 근거", s["td"])],
    ]
    story.append(info_table(file_rows, s, [40 * mm, _WIDTH - 40 * mm],
                            header=None))
    story.append(gap(6))
    story.append(body("<b>파일별 필수 컬럼</b> (대소문자·띄어쓰기 포함하여 정확히 일치해야 함):", s))
    story.append(gap(3))

    req_rows = [
        [Paragraph("사업비정보.csv", s["td_b"]),
         Paragraph("원가요소, 코스트센터", s["td"]),
         Paragraph("계약체결비·계약유지비·손해조사비·투자관리비·간접비·공통비·합계", s["td"])],
        [Paragraph("계정정보.csv", s["td_b"]),
         Paragraph("계정번호", s["td"]),
         Paragraph("계정명·계정그룹ID·계정그룹명·대상정의·사용 부서·비용 지급 범위·산출기준", s["td"])],
        [Paragraph("부서정보.csv", s["td_b"]),
         Paragraph("Cost Center Code, 직간접구분, 팀", s["td"]),
         Paragraph("—", s["td_c"])],
        [Paragraph("출력.csv", s["td_b"]),
         Paragraph("출력전표", s["td"]),
         Paragraph("주관부서·사용부서·(분류 근거 7컬럼)", s["td"])],
    ]
    story.append(info_table(
        req_rows, s,
        [34 * mm, 42 * mm, _WIDTH - 76 * mm],
        header=["파일", "필수 컬럼", "권장/선택 컬럼"]))
    story.append(gap(6))
    story.append(subhead("2-2. 각 CSV 파일 조건", s))
    story.extend(bullets([
        "<b>인코딩</b>: UTF-8(BOM 포함 권장) 또는 CP949. 프로그램은 utf-8-sig 우선, 실패 시 CP949로 자동 재시도합니다.",
        "<b>필수 컬럼 누락 시</b>: 즉시 오류 메시지를 표시하고 중단됩니다(조용한 0원 처리 없음).",
        "<b>코드 형식</b>: 원가요소·출력전표·코스트센터 코드는 양쪽 파일에서 동일 표기여야 합니다(예: \"5001\").",
        "<b>금액 컬럼</b>: 쉼표·단독 대시('-')·빈칸은 0으로 안전 변환됩니다.",
    ], s))

    story.append(gap(10))

    # ── 3. 실행 방법 ─────────────────────────────────────────────────────
    story.append(section_title("3", "실행 방법", s))
    story.append(gap(4))
    story.append(warn_box(
        "내려받은 직후에는 몇 초 기다린 뒤 실행하세요",
        "Windows에서 EXE를 내려받은 직후 곧바로 실행하면 백신·운영체제의 보안 검사 때문에 "
        "실행 오류가 발생할 수 있습니다. 다운로드 후 <b>몇 초 정도 기다린 뒤</b> 실행하면 안정적입니다.", s))
    story.append(gap(6))
    story.append(body("<b>전표분석서_자동화.exe</b>를 더블클릭하여 실행합니다. 진행 순서는 다음과 같습니다.", s))
    story.append(gap(4))

    step_rows = [
        [Paragraph("Step 1", s["td_b"]),
         Paragraph("<b>색상·로고·문구 설정</b> — 포인트 색상과 로고(최대 3슬롯, 높이 슬라이더 16~96pt·실시간 "
                   "미리보기)를 설정하고, [문구 설정] 탭에서 PDF 표시 문구 편집 및 성격 컬럼 추가/삭제까지 "
                   "할 수 있습니다. ('아니오' 시 저장된 설정으로 진행)", s["td"])],
        [Paragraph("Step 2", s["td_b"]),
         Paragraph("<b>CSV 파일 4개 선택</b> — 사업비정보 → 계정정보 → 부서정보 → 출력 순으로 "
                   "각각 지정합니다. 직전 실행 경로가 기본값으로 기억됩니다.", s["td"])],
        [Paragraph("Step 3", s["td_b"]),
         Paragraph("<b>필수 컬럼 검증</b> — 4개 파일의 필수 컬럼을 확인합니다. 누락 시 어떤 "
                   "파일의 어떤 컬럼이 없는지 오류 메시지로 안내하고 중단됩니다.", s["td"])],
        [Paragraph("Step 4", s["td_b"]),
         Paragraph("<b>미등록 부서 감지</b> — 사업비정보의 코스트센터로 매핑되는 팀이 출력.csv의 "
                   "주관/사용부서에 없으면 경고합니다. [무시하고 진행] 또는 [종료 후 수정]을 선택합니다.", s["td"])],
        [Paragraph("Step 5", s["td_b"]),
         Paragraph("<b>저장 폴더 선택</b> — 생성된 PDF를 저장할 폴더를 지정합니다.", s["td"])],
        [Paragraph("Step 6", s["td_b"]),
         Paragraph("<b>PDF 생성 완료</b> — 코드별 PDF가 생성되고, 성공/경고/실패 건수 요약 창이 표시됩니다. "
                   "미등록 팀과 함께 <b>미등록 성격 컬럼</b>(사업비정보.csv에 없는 성격 컬럼)도 경고로 안내됩니다.", s["td"])],
    ]
    story.append(info_table(step_rows, s, [20 * mm, _WIDTH - 20 * mm], header=None))
    story.append(gap(4))
    story.append(Paragraph(
        "※ 참고: 미등록 부서를 무시하고 진행하면 해당 부서의 실적은 부점귀속·성격별 집계에서 "
        "제외됩니다. 정확한 결과가 필요하면 출력.csv에 부서를 추가한 뒤 재실행하십시오.", s["note"]))

    story.append(gap(10))

    # ── 3-7. 화면별 안내 (스크린샷) ───────────────────────────────────────
    story.append(subhead("화면별 안내", s))
    story.append(gap(4))

    # Step 1 — 색상·로고·문구 설정
    story.append(body("<b>Step 1. 색상·로고·문구 설정</b>", s))
    story.append(gap(3))
    story.extend(screenshot("01_brand_tab.png",
                            "[브랜드 설정] 탭 — 포인트 색상 및 로고 슬롯(최대 3개) 설정", s))
    story.append(gap(4))
    story.extend(screenshot("02_labels_tab.png",
                            "[문구 설정] 탭 — PDF에 인쇄되는 모든 표시 문구를 편집", s))
    story.append(gap(4))
    story.extend(screenshot("03_nature_cols.png",
                            "[문구 설정] 탭 하단 — 성격 컬럼 추가/삭제 영역(CSV 컬럼명 = 표시명)", s))
    story.append(gap(5))
    story.append(warn_box(
        "문구 설정의 성격 컬럼명은 CSV 컬럼명과 정확히 일치해야 합니다",
        "[문구 설정] 탭에서 입력한 성격 컬럼명은 <b>사업비정보.csv의 컬럼 헤더와 글자·띄어쓰기까지 "
        "동일</b>해야 합니다. 이름이 다르면 오류 없이 해당 컬럼이 <b>빈 칸으로 출력</b>되며, 최종 요약 "
        "창에 '미등록 성격 컬럼' 경고로만 표시됩니다.", s))
    story.append(gap(8))

    # Step 2 — CSV 파일 선택
    story.append(body("<b>Step 2. CSV 파일 4개 선택</b>", s))
    story.append(gap(3))
    story.extend(screenshot("04_file_select.png",
                            "입력 파일 선택 — 사업비정보·계정정보·부서정보·출력 4개 CSV 지정", s))
    story.append(gap(8))

    # Step 4 — 미등록 부서 감지
    story.append(body("<b>Step 4. 미등록 부서 감지</b>", s))
    story.append(gap(3))
    story.extend(screenshot("05_unregistered.png",
                            "미등록 부서 감지 — [무시하고 진행] 또는 [종료 후 수정] 선택", s))
    story.append(gap(8))

    # Step 6 — 최종 요약
    story.append(body("<b>Step 6. 최종 요약</b>", s))
    story.append(gap(3))
    story.extend(screenshot("06_summary.png",
                            "PDF 일괄 생성 결과 — 성공/경고/실패 및 미등록 성격 컬럼 경고", s))

    story.append(gap(10))

    # ── 4. 출력 결과물 ───────────────────────────────────────────────────
    story.append(section_title("4", "출력 결과물", s))
    story.append(gap(4))
    story.append(body(
        "출력전표 코드 1개당 PDF 1개가 생성됩니다. 파일명 형식은 "
        "<b>사업비_전표_분석_{코드}.pdf</b> (예: 사업비_전표_분석_5001.pdf)이며, A4 세로로 구성됩니다.", s))
    story.append(gap(4))
    pdf_rows = [
        [Paragraph("헤더", s["td_b"]),
         Paragraph("로고 + 포인트 색상 바, 계정 정보(계정코드/계정명/사업비 코드/채널구분)", s["td"])],
        [Paragraph("1. 대상정의", s["td_b"]),
         Paragraph("대상정의·범위(사용 부서)·지급대상(비용 지급 범위)·산출기준", s["td"])],
        [Paragraph("2. 부점귀속", s["td_b"]),
         Paragraph("주관부서 귀속 / 실제 사용부서 귀속을 나란히 표시(대상금액·구성비)", s["td"])],
        [Paragraph("3-1. 직접·공통비", s["td_b"]),
         Paragraph("부서정보 직간접구분 기준 직접비·공통비 분류 금액", s["td"])],
        [Paragraph("3-2. 성격별", s["td_b"]),
         Paragraph("성격 분류는 직접비로 구분된 금액을 계정 성격에 따라 분류하며, "
                   "공통비로 구분된 금액은 귀속 부서 성격에 따라 공통부서에서 직접 부서로 "
                   "순차적 배부됨.", s["td"])],
        [Paragraph("4. 분류 근거", s["td_b"]),
         Paragraph("출력.csv의 직접비/공통비·성격별 근거 텍스트(금액 0원 항목은 제외)", s["td"])],
    ]
    story.append(info_table(pdf_rows, s, [32 * mm, _WIDTH - 32 * mm], header=None))

    story.append(gap(10))

    # ── 5. 절대 변경하면 안 되는 것들 ────────────────────────────────────
    story.append(section_title("5", "절대 변경하면 안 되는 것들", s))
    story.append(gap(4))
    story.extend(bullets([
        "<b>각 CSV 파일의 필수 컬럼명</b> — 대소문자·띄어쓰기까지 정확히 유지해야 합니다. "
        "특히 계정정보.csv의 <b>\"사용 부서\"</b>, 부서정보.csv의 <b>\"Cost Center Code\"</b>는 "
        "공백·대문자 표기를 그대로 두십시오.",
        "<b>직간접구분 값</b> — 부서정보.csv의 직간접구분 컬럼 값은 반드시 <b>\"직접\"</b> 또는 "
        "<b>\"공통\"</b> 중 하나여야 합니다. 다른 표기(예: 간접, 직접비용)는 분류되지 않습니다.",
        "<b>코드 일관성</b> — 원가요소·출력전표·코스트센터 코드는 파일 간 동일 표기를 유지하십시오.",
        "<b>문구 설정의 성격 컬럼명</b> — [문구 설정] 탭에서 입력한 성격 컬럼명은 사업비정보.csv의 "
        "컬럼 헤더와 정확히 일치시키십시오. 불일치 시 오류 없이 빈 칸으로 출력되고 요약 창에 경고만 "
        "표시됩니다(3장 참고).",
        "<b>프로그램 구성 파일</b>(.exe, column_config.yaml의 논리명 등)을 임의로 수정·삭제하지 마십시오.",
    ], s))
    story.append(gap(3))
    story.append(Paragraph(
        "※ 컬럼명이 부득이 바뀐 경우에는 column_config.yaml의 큰따옴표 안 값만 수정하면 "
        "다음 실행부터 자동 반영됩니다(논리명/왼쪽 키는 변경 금지).", s["note"]))

    story.append(gap(10))

    # ── 6. 문제 해결 (FAQ) ───────────────────────────────────────────────
    story.append(section_title("6", "문제 해결 (FAQ)", s))
    story.append(gap(4))
    faqs = [
        ("Q1. CSV 파일 선택 후 오류가 발생해요",
         "필수 컬럼 누락 또는 파일 손상 가능성이 높습니다. 오류 메시지에 표시된 <b>파일명과 누락 컬럼</b>을 "
         "확인하고, templates/ 양식과 컬럼명을 대조하십시오. 파일이 Excel 등에서 열려 있으면 닫은 뒤 다시 "
         "시도하십시오. 인코딩이 깨졌다면 UTF-8(BOM) 또는 CP949로 저장하십시오."),
        ("Q2. 미등록 부서 경고가 떠요",
         "사업비정보의 코스트센터로 매핑되는 팀이 출력.csv의 주관부서/사용부서에 없을 때 발생합니다. "
         "정확한 집계를 위해서는 [종료 후 수정]을 누르고 출력.csv에 해당 부서를 추가한 뒤 재실행하십시오. "
         "코스트센터 자체가 부서정보에 없으면 \"(미매칭:코드)\"로 표시되니 부서정보.csv도 확인하십시오."),
        ("Q3. PDF가 생성되지 않아요",
         "① 출력.csv의 출력전표 코드가 사업비정보의 원가요소에 존재하는지, ② 저장 폴더에 쓰기 권한이 "
         "있는지, ③ 동일 파일이 열려 있지 않은지 확인하십시오. 일부 코드만 실패한 경우 요약 창에 코드별 "
         "사유가 표시되며, 다른 코드의 PDF에는 영향이 없습니다."),
        ("Q4. 생성된 PDF에서 한글이 네모(□)로 깨져요",
         "한글 폰트(맑은 고딕)를 찾지 못한 경우입니다. Windows의 C:\\Windows\\Fonts에 malgun.ttf가 있는지 "
         "확인하십시오. 개발 환경에서는 FONT_PATH 환경변수로 폰트 경로를 지정할 수 있습니다."),
        ("Q5. 성격별 금액이 0으로 나와요",
         "① 사업비정보.csv의 성격 컬럼(계약체결비·계약유지비·손해조사비·투자관리비·간접비·공통비) 값이 "
         "비어 있거나, ② 직간접구분이 '직접'/'공통'이 아니어서 분류에서 빠졌거나, ③ 해당 부서가 "
         "출력.csv에 등록되지 않아 집계에서 제외된 경우입니다. 세 가지를 순서대로 점검하십시오."),
    ]
    for q, a in faqs:
        story.append(Paragraph(q, s["faq_q"]))
        story.append(Paragraph(a, s["faq_a"]))
        story.append(gap(5))

    story.append(gap(6))

    # ── 7. 변경 이력 ─────────────────────────────────────────────────────
    story.append(section_title("7", "변경 이력", s))
    story.append(gap(4))
    hist_rows = [
        [Paragraph("v2.1.0", s["td_b"]), Paragraph("2026/06", s["td_c"]),
         Paragraph("<b>문구 설정 탭</b> 추가(PDF 표시 문구 직접 편집); <b>성격 컬럼 동적 추가/삭제</b> "
                   "(CSV명=표시명 통합); <b>미등록 성격 컬럼 경고</b>를 최종 요약 창에 표시; "
                   "사용설명서에 화면 스크린샷 추가", s["td"])],
        [Paragraph("v2.0.0", s["td_b"]), Paragraph("2026/06", s["td_c"]),
         Paragraph("입력 형식을 엑셀(다중 시트)에서 <b>CSV 4개 파일</b>로 전환; 금액 소스를 "
                   "'합계' 컬럼으로 정리; 성격 분류 컬럼을 <b>6종</b>(계약체결비·계약유지비·손해조사비·"
                   "투자관리비·간접비·공통비)으로 개정; 실행 흐름(Step 1~6) 재정비", s["td"])],
        [Paragraph("v1.2.0", s["td_b"]), Paragraph("2026/06", s["td_c"]),
         Paragraph("로고 슬롯별 높이 슬라이더 및 실시간 미리보기 추가(엑셀 입력 기준)", s["td"])],
    ]
    story.append(info_table(
        hist_rows, s, [22 * mm, 22 * mm, _WIDTH - 44 * mm],
        header=["버전", "날짜", "주요 변경 내용"]))
    story.append(gap(8))
    story.append(Paragraph(
        "본 설명서는 v2.1.0(현재 코드) 기준으로 작성되었습니다. © EY Korea", s["note"]))

    return story


# ════════════════════════════════════════════════════════════════════════════
#  메인
# ════════════════════════════════════════════════════════════════════════════

def main() -> None:
    fn, fb = setup_fonts()
    s = build_styles(fn, fb)

    doc = BaseDocTemplate(
        str(_OUT_PDF), pagesize=_PAGE,
        leftMargin=_MARGIN, rightMargin=_MARGIN,
        topMargin=_MARGIN + 4 * mm, bottomMargin=_MARGIN,
        title="지급사업비 분석 자동화 프로그램 사용설명서 v2.1.0",
        author="EY Korea",
    )
    frame = Frame(_MARGIN, _MARGIN, _WIDTH, _PAGE[1] - 2 * _MARGIN - 4 * mm,
                  id="main", leftPadding=0, rightPadding=0,
                  topPadding=0, bottomPadding=0)
    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[frame], onPage=on_cover),
        PageTemplate(id="body", frames=[frame], onPage=on_page),
    ])

    story = build_story(s)
    doc.build(story)
    print(f"[done] PDF 생성 완료 → {_OUT_PDF}  ({_OUT_PDF.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
