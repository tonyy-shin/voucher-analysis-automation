"""
generate_template_pdf.py
역할: 실제 출력 PDF(pdf_exporter.py)와 동일한 섹션 구조를 가진 '빈 양식' 템플릿 PDF를 생성한다.
      모든 데이터 셀은 비워 두고 테이블 구조와 라벨만 표시하여 출력 양식을 미리 보여준다.

EY Korea 브랜드 스타일(색상·폰트·헤더/푸터 바)을 generate_manual.py와 동일하게 적용한다.

실행:  python3 generate_template_pdf.py
출력:  release/templates/pdf/출력_템플릿.pdf
"""
from __future__ import annotations

import os
import warnings
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle,
)

# ── 출력 경로 ─────────────────────────────────────────────────────────────────
_HERE = Path(__file__).resolve().parent
_OUT_PDF = _HERE / "templates" / "pdf" / "출력_템플릿.pdf"

# ── EY Korea 브랜드 색상 (generate_manual.py와 동일) ──────────────────────────
EY_YELLOW = colors.HexColor("#FFE600")
EY_DARK   = colors.HexColor("#2E2E38")
EY_GRAY   = colors.HexColor("#747480")
_WHITE    = colors.white
_LIGHTBG  = colors.HexColor("#F4F4F6")

# ── 페이지 (A4 세로) — 실제 출력 PDF 비율을 따르기 위해 15mm 여백 ─────────────
_PAGE   = A4
_MARGIN = 15 * mm
_WIDTH  = _PAGE[0] - 2 * _MARGIN

# ── 성격별 6종 (data_loader.NATURE_COLS) ──────────────────────────────────────
NATURE_COLS = ["계약비", "유지비", "손해조사비", "투자관리비", "간접비", "공통비"]

# ── 섹션별 컬럼 너비 (pdf_exporter.py 비율 재사용) ────────────────────────────
_COLS_ACCT  = [r / 7.0 * _WIDTH for r in [1.0, 1.0, 1.5, 1.0, 2.5]]
_COLS_2     = [0.22 * _WIDTH, 0.78 * _WIDTH]
_COLS_7     = [r / 7.2 * _WIDTH for r in [1.2, 1.0, 1.2, 0.8, 1.0, 1.2, 0.8]]
_COLS_8     = [r / 8.3 * _WIDTH for r in [1.3, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0]]
_COLS_DI    = [0.20 * _WIDTH, 0.80 * _WIDTH]
_COLS_BASIS = [0.18 * _WIDTH, 0.64 * _WIDTH, 0.18 * _WIDTH]


# ════════════════════════════════════════════════════════════════════════════
#  폰트 등록 (generate_manual.py와 동일 패턴)
# ════════════════════════════════════════════════════════════════════════════

def _resolve_font_paths() -> tuple[Path | None, Path | None]:
    env = os.environ.get("FONT_PATH", "")
    if env and Path(env).exists():
        p = Path(env)
        return p, p
    home = Path.home()
    reg_cands = [
        home / ".fonts" / "NanumGothic-Regular.ttf",
        home / ".local/share/fonts" / "NanumGothic-Regular.ttf",
        Path("/usr/share/fonts/truetype/nanum/NanumGothic.ttf"),
        Path("/usr/share/fonts/nanum/NanumGothic.ttf"),
    ]
    bold_cands = [
        home / ".fonts" / "NanumGothic-Bold.ttf",
        home / ".local/share/fonts" / "NanumGothic-Bold.ttf",
        Path("/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"),
        Path("/usr/share/fonts/nanum/NanumGothicBold.ttf"),
    ]
    reg = next((c for c in reg_cands if c.exists()), None)
    bold = next((c for c in bold_cands if c.exists()), None)
    return reg, bold


def setup_fonts() -> tuple[str, str]:
    reg, bold = _resolve_font_paths()
    if reg is None:
        warnings.warn("NanumGothic 미발견 — Helvetica 폴백(한글 깨짐).")
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
        warnings.warn(f"폰트 등록 실패({exc}) — Helvetica 폴백.")
        return "Helvetica", "Helvetica-Bold"


# ════════════════════════════════════════════════════════════════════════════
#  스타일
# ════════════════════════════════════════════════════════════════════════════

def build_styles(fn: str, fb: str) -> dict:
    return {
        "title":  ParagraphStyle("title", fontName=fb, fontSize=13, leading=17, textColor=_WHITE),
        "sec":    ParagraphStyle("sec", fontName=fb, fontSize=10, leading=13, textColor=EY_DARK),
        "small":  ParagraphStyle("small", fontName=fn, fontSize=7.5, leading=10, textColor=EY_GRAY),
        "note":   ParagraphStyle("note", fontName=fn, fontSize=8.5, leading=12, textColor=EY_GRAY),
        # 표 헤더(흰 글씨, 가운데)
        "hdr_w":  ParagraphStyle("hdr_w", fontName=fb, fontSize=8.5, leading=11,
                                 alignment=TA_CENTER, textColor=_WHITE),
        # 라벨(진한 글씨)
        "lab":    ParagraphStyle("lab", fontName=fb, fontSize=8.5, leading=11, textColor=EY_DARK),
        "body":   ParagraphStyle("body", fontName=fn, fontSize=8.5, leading=11, textColor=EY_DARK),
        "right":  ParagraphStyle("right", fontName=fn, fontSize=8.5, leading=11,
                                 alignment=TA_RIGHT, textColor=EY_DARK),
        "center": ParagraphStyle("center", fontName=fn, fontSize=8.5, leading=11,
                                 alignment=TA_CENTER, textColor=EY_DARK),
    }


def _P(text, style) -> Paragraph:
    return Paragraph(str(text), style)


def _base_style() -> list:
    return [
        ("GRID",         (0, 0), (-1, -1), 0.5, EY_GRAY),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING",   (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
    ]


# ════════════════════════════════════════════════════════════════════════════
#  페이지 데코레이션 (상/하단 바 — generate_manual.py와 동일)
# ════════════════════════════════════════════════════════════════════════════

_FOOTER_LEFT = "지급사업비 분석 자동화 — 출력 양식 템플릿"
_FOOTER_RIGHT = "© EY Korea"


def on_page(canvas, doc):
    canvas.saveState()
    w, h = _PAGE
    canvas.setFillColor(EY_YELLOW)
    canvas.rect(0, h - 6 * mm, w, 6 * mm, stroke=0, fill=1)
    canvas.setFillColor(EY_DARK)
    canvas.rect(0, h - 7 * mm, w, 1 * mm, stroke=0, fill=1)
    canvas.setStrokeColor(EY_GRAY)
    canvas.setLineWidth(0.5)
    canvas.line(_MARGIN, 12 * mm, w - _MARGIN, 12 * mm)
    canvas.setFont("NG" if "NG" in pdfmetrics._fonts else "Helvetica", 7.5)
    canvas.setFillColor(EY_GRAY)
    canvas.drawString(_MARGIN, 8 * mm, _FOOTER_LEFT)
    canvas.drawRightString(w - _MARGIN, 8 * mm, f"{_FOOTER_RIGHT}  |  {doc.page}")
    canvas.restoreState()


# ════════════════════════════════════════════════════════════════════════════
#  섹션 빌더 (모든 데이터 셀은 빈칸)
# ════════════════════════════════════════════════════════════════════════════

def _tbl_header_bar(s: dict) -> Table:
    t = Table([[""]], colWidths=[_WIDTH], rowHeights=[8])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), EY_DARK)]))
    return t


def _tbl_header(s: dict) -> Table:
    data = [[_P("[ 사업비 전표 분석 ]", s["title"]), _P("EY Korea", s["title"])]]
    t = Table(data, colWidths=[_WIDTH * 0.75, _WIDTH * 0.25])
    t.setStyle(TableStyle([
        ("BACKGROUND",  (0, 0), (-1, -1), EY_DARK),
        ("ALIGN",       (1, 0), (1, 0), "RIGHT"),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",(0, 0), (-1, -1), 8),
        ("TOPPADDING",  (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 7),
    ]))
    return t


def _tbl_account(s: dict) -> Table:
    data = [
        [_P("Account", s["hdr_w"]), _P("Account Code", s["hdr_w"]),
         _P("Account Name", s["hdr_w"]), _P("사업비 코드", s["hdr_w"]),
         _P("채널구분", s["hdr_w"])],
        ["", "", "", "", ""],
    ]
    cmds = _base_style() + [
        ("BACKGROUND", (0, 0), (-1, 0), EY_GRAY),
        ("BACKGROUND", (0, 1), (-1, 1), _WHITE),
        ("SPAN",       (0, 0), (0, 1)),
        ("BACKGROUND", (0, 0), (0, 1), EY_GRAY),
    ]
    return Table(data, colWidths=_COLS_ACCT, rowHeights=[None, 16], style=TableStyle(cmds))


def _tbl_target(s: dict) -> list:
    data = [
        [_P("대상정의", s["lab"]), ""],
        [_P("범위",     s["lab"]), ""],
        [_P("지급대상", s["lab"]), ""],
        [_P("산출기준", s["lab"]), ""],
    ]
    cmds = _base_style() + [
        ("BACKGROUND", (0, 0), (0, 3), _LIGHTBG),
        ("BACKGROUND", (1, 0), (1, 3), _WHITE),
    ]
    return [_P("1. 대상정의", s["sec"]), Spacer(1, 2),
            Table(data, colWidths=_COLS_2, rowHeights=[16] * 4, style=TableStyle(cmds))]


def _tbl_dept(s: dict) -> list:
    # 헤더 2행 + 빈 데이터 2행 + 총계 행
    data = [
        [_P("귀속현황", s["hdr_w"]),
         _P("주관부서 귀속", s["hdr_w"]), None, None,
         _P("실제 사용부서 귀속", s["hdr_w"]), None, None],
        [None,
         _P("부점명", s["hdr_w"]), _P("대상금액(단위: 원)", s["hdr_w"]), _P("구성비(%)", s["hdr_w"]),
         _P("부점명", s["hdr_w"]), _P("대상금액(단위: 원)", s["hdr_w"]), _P("구성비(%)", s["hdr_w"])],
    ]
    for _ in range(2):
        data.append(["", "", "", "", "", "", ""])
    tot_idx = len(data)
    data.append([_P("총계", s["hdr_w"]), "", "", "", "", "", ""])

    cmds = _base_style() + [
        ("SPAN", (0, 0), (0, 1)),
        ("SPAN", (1, 0), (3, 0)),
        ("SPAN", (4, 0), (6, 0)),
        ("BACKGROUND", (0, 0), (6, 1), EY_GRAY),
        ("BACKGROUND", (0, 2), (6, tot_idx - 1), _WHITE),
        ("BACKGROUND", (0, tot_idx), (6, tot_idx), _LIGHTBG),
        ("ALIGN", (0, 0), (0, 1), "CENTER"),
    ]
    title = _P("2. 부점귀속 현황", s["sec"])
    desc = _P("주관부서 귀속은 전표 발의 및 예산 통제 조직 기준이며, 실제 사용부서 귀속은 "
              "업무 수행 및 비용 효익이 실질적으로 귀속되는 조직 기준.", s["small"])
    return [title, Spacer(1, 2), desc, Spacer(1, 2),
            Table(data, colWidths=_COLS_7,
                  rowHeights=[None, None, 15, 15, 15], style=TableStyle(cmds))]


def _tbl_nature_31(s: dict) -> list:
    data = [
        [_P("구분", s["hdr_w"]), _P("분류금액", s["hdr_w"])],
        [_P("직접비", s["lab"]), ""],
        [_P("공통비", s["lab"]), ""],
    ]
    cmds = _base_style() + [
        ("BACKGROUND", (0, 0), (1, 0), EY_GRAY),
        ("BACKGROUND", (0, 1), (0, 2), _LIGHTBG),
        ("BACKGROUND", (1, 1), (1, 2), _WHITE),
    ]
    return [_P("3. 사업비 분류", s["sec"]), Spacer(1, 2),
            _P("3-1. 직접비 공통비 분류 결과", s["sec"]), Spacer(1, 2),
            Table(data, colWidths=_COLS_DI, rowHeights=[None, 16, 16], style=TableStyle(cmds))]


def _tbl_nature_32(s: dict) -> list:
    header_labels = NATURE_COLS + ["합계"]
    last_col = len(header_labels)
    data = [[_P("귀속현황", s["hdr_w"])]
            + [_P(c, s["hdr_w"]) for c in header_labels]]
    for _ in range(2):
        data.append([""] + ["" for _ in header_labels])
    tot_idx = len(data)
    data.append([_P("총계", s["hdr_w"])] + ["" for _ in header_labels])

    cmds = _base_style() + [
        ("BACKGROUND", (0, 0), (last_col, 0), EY_GRAY),
        ("ALIGN",      (0, 0), (0, 0), "CENTER"),
        ("BACKGROUND", (0, 1), (last_col, tot_idx - 1), _WHITE),
        ("BACKGROUND", (0, tot_idx), (last_col, tot_idx), _LIGHTBG),
    ]
    title = _P("3-2. 성격별 분류 결과", s["sec"])
    desc = _P("성격 분류는 직접비로 구분된 금액을 계정 성격에 따라 분류하며, 공통비로 구분된 "
              "금액은 귀속 부서 성격에 따라 공통부서에서 직접 부서로 순차적 배부됨.", s["small"])
    return [title, Spacer(1, 2), desc, Spacer(1, 2),
            Table(data, colWidths=_COLS_8, rowHeights=[None, 15, 15, 15], style=TableStyle(cmds))]


def _tbl_basis(s: dict) -> list:
    data = [
        [_P("직 • 공통비", s["lab"]), "", _P("총괄문서 참조", s["center"])],
        [_P("성격별 분류", s["lab"]), "", _P("총괄문서 참조", s["center"])],
    ]
    cmds = _base_style() + [
        ("BACKGROUND", (0, 0), (0, 1), _LIGHTBG),
        ("BACKGROUND", (1, 0), (1, 1), _WHITE),
        ("BACKGROUND", (2, 0), (2, 1), _WHITE),
        ("ALIGN", (0, 0), (0, 1), "CENTER"),
        ("ALIGN", (2, 0), (2, 1), "CENTER"),
    ]
    return [_P("4. 분류 근거", s["sec"]), Spacer(1, 2),
            Table(data, colWidths=_COLS_BASIS, rowHeights=[30, 30], style=TableStyle(cmds))]


def _tbl_footer(s: dict) -> Table:
    t = Table([[""]], colWidths=[_WIDTH], rowHeights=[8])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), EY_DARK)]))
    return t


# ════════════════════════════════════════════════════════════════════════════
#  메인
# ════════════════════════════════════════════════════════════════════════════

def build_story(s: dict) -> list:
    gap = Spacer(1, 5)
    story: list = []
    story.append(_P("※ 이 문서는 출력 양식 템플릿입니다. 실제 데이터는 프로그램 실행 후 생성됩니다.",
                    s["note"]))
    story.append(Spacer(1, 4))
    story.append(_tbl_header_bar(s))
    story.append(_tbl_header(s));        story.append(gap)
    story.append(_tbl_account(s));       story.append(gap)
    story.extend(_tbl_target(s));        story.append(gap)
    story.extend(_tbl_dept(s));          story.append(gap)
    story.extend(_tbl_nature_31(s));     story.append(Spacer(1, 3))
    story.extend(_tbl_nature_32(s));     story.append(gap)
    story.extend(_tbl_basis(s));         story.append(gap)
    story.append(_tbl_footer(s))
    return story


def main() -> None:
    _OUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    fn, fb = setup_fonts()
    s = build_styles(fn, fb)

    doc = BaseDocTemplate(
        str(_OUT_PDF), pagesize=_PAGE,
        leftMargin=_MARGIN, rightMargin=_MARGIN,
        topMargin=_MARGIN + 2 * mm, bottomMargin=_MARGIN,
        title="사업비 전표 분석 — 출력 양식 템플릿",
        author="EY Korea",
    )
    frame = Frame(_MARGIN, _MARGIN, _WIDTH, _PAGE[1] - 2 * _MARGIN - 2 * mm,
                  id="main", leftPadding=0, rightPadding=0,
                  topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="body", frames=[frame], onPage=on_page)])

    doc.build(build_story(s))
    print(f"[done] 템플릿 PDF 생성 완료 → {_OUT_PDF}  ({_OUT_PDF.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
