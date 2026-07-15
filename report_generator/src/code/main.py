"""
main.py
역할: 프로그램 진입점. tkinter 다이얼로그로 사용자 입력을 받고
      출력> 시트의 모든 출력전표 코드를 일괄 처리한다.

--noconsole 환경 (PyInstaller EXE): 터미널 창 없이 팝업 다이얼로그만 사용한다.
모든 오류는 tkinter.messagebox로 사용자에게 안내한 뒤 sys.exit(1)로 종료한다.
"""
from __future__ import annotations

import copy
import json
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, scrolledtext, ttk

import data_loader
import processor
import pdf_exporter
from data_loader import _SHEET_OUTPUT, COLUMN_MAP
from company_theme import CompanyTheme
from theme_manager import GRAY_DEFAULT, load_theme, save_theme
from config_manager import (
    load_config, save_config, validate_csv_columns, default_display_labels,
)

# ── 모듈 상수 ─────────────────────────────────────────────────────────────────
_PDF_FILENAME_TEMPLATE = "사업비_전표_분석_{}.pdf"

if getattr(sys, "frozen", False):
    _PATHS_DIR = Path(sys.executable).parent
else:
    _PATHS_DIR = Path(__file__).parents[3]

_LAST_PATHS_FILE = _PATHS_DIR / "last_paths.json"


def _load_last_paths() -> dict:
    try:
        with open(_LAST_PATHS_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_last_paths(paths: dict) -> None:
    try:
        with open(_LAST_PATHS_FILE, "w", encoding="utf-8") as f:
            json.dump(paths, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def _fmt_slot_label(path: str | None) -> str:
    """슬롯에 표시할 로고 파일명 반환 (최대 30자, 없으면 '미지정')."""
    if not path:
        return "미지정"
    name = Path(path).name
    return name if len(name) <= 30 else "..." + name[-27:]


# ── 필수 컬럼 삭제 가드 ────────────────────────────────────────────────────────
# JOIN 키·분류 기준으로 파이프라인이 구조적으로 의존하는 컬럼 — 어떤 동적 행
# 편집기에서도 이 컬럼을 가리키는 행은 삭제할 수 없다 (라벨/값 편집은 자유).
_REQUIRED_LOGICAL_KEYS: tuple[str, ...] = (
    "원가요소", "코스트센터", "계정번호", "cc_code", "직간접구분", "팀", "출력전표",
)


def _build_required_columns(config: dict) -> set[str]:
    """config['columns'] 매핑을 반영해 실제 필수 컬럼명 집합을 만든다.

    columns: 매핑은 GUI가 없어 사용자가 YAML을 직접 고칠 수 있으므로, 로직 키의
    리터럴 기본값이 아니라 현재 config 값을 기준으로 매번 계산한다 — 그래야
    예: cc_code 를 다른 실제 헤더명으로 재매핑해도 가드가 계속 유효하다.
    """
    cols = config.get("columns", {})
    return {str(cols.get(k, k)).strip() for k in _REQUIRED_LOGICAL_KEYS}


def _is_required_column(value: str, required_columns: set[str]) -> bool:
    return value.strip() in required_columns


def _make_rows_editor(
    parent: tk.Frame,
    dlg: tk.Toplevel,
    rows_state: list[dict],
    fields: list[tuple[str, str, int]],
    new_row,
    min_rows: int = 1,
    section_name: str = "항목",
    guard_field: str | None = None,
    required_columns: set[str] | None = None,
):
    """동적 행 편집기 팩토리 — 균일한 dict 행 리스트(계정 테이블/섹션1/2/3-1) 공용.

    3-2 성격 편집기(str 리스트 상태)와 섹션4 편집기(타입별 조건부 필드)는 상태 구조가
    달라 의도적으로 별도 구현을 유지한다 — 이 팩토리로 억지로 통합하지 말 것.

    flush()는 fields 에 선언된 키만 행 dict 에 되쓰므로, 선언되지 않은 키
    (예: account_table.columns 의 width)는 편집 중에도 보존된다.

    Args:
        parent:       행 프레임을 붙일 부모 (스크롤 영역 내부 프레임)
        dlg:          messagebox 부모 다이얼로그
        rows_state:   config에서 deep-copy 한 행 dict 리스트 (in-place 편집됨)
        fields:       [(행 dict 키, 캡션, Entry 너비), ...]
        new_row:      "+ 추가" 시 append 할 새 행 dict 를 반환하는 콜러블
        min_rows:     최소 행 수 (삭제 가드)
        section_name: 경고 메시지에 표시할 섹션명
        guard_field:  이 행이 CSV 컬럼을 가리키는 필드 키 (예: "csv_column").
                      None이면 필수 컬럼 삭제 가드를 적용하지 않는다 (예: 섹션3-1의
                      "keyword"는 컬럼명이 아닌 값 부분일치 문자열이라 대상이 아님).
        required_columns: 필수 컬럼명 집합 (_build_required_columns 결과).
                      guard_field가 주어지면 필수로 함께 전달해야 한다.

    Returns:
        (frame, flush, render, add) — 호출부가 frame 을 배치하고 render()를 1회 호출한다.
        flush()는 입력 중인 Entry 값을 rows_state 로 반영한다 (저장 전 필수 호출).
    """
    frame = tk.Frame(parent)
    vars_list: list[dict[str, tk.StringVar]] = []   # rows_state와 인덱스 정렬

    def flush() -> None:
        for row, var_map in zip(rows_state, vars_list):
            for key, var in var_map.items():
                row[key] = var.get()

    def _delete(idx: int) -> None:
        if guard_field is not None and idx < len(vars_list):
            live_val = vars_list[idx][guard_field].get()
            if _is_required_column(live_val, required_columns or set()):
                messagebox.showwarning(
                    "삭제 불가",
                    f"{section_name}의 필수 컬럼(\"{live_val.strip()}\")은 삭제할 수 없습니다.",
                    parent=dlg,
                )
                return
        if len(rows_state) <= min_rows:
            messagebox.showwarning(
                "최소 개수",
                f"{section_name}은(는) 최소 {min_rows}개 이상이어야 합니다.",
                parent=dlg,
            )
            return
        flush()
        del rows_state[idx]
        render()

    def render() -> None:
        for w in frame.winfo_children():
            w.destroy()
        vars_list.clear()
        for idx, row in enumerate(rows_state):
            rf = tk.Frame(frame)
            rf.grid(row=idx, column=0, sticky="w", padx=8, pady=2)
            var_map: dict[str, tk.StringVar] = {}
            guard_var: tk.StringVar | None = None
            for key, caption, width in fields:
                tk.Label(rf, text=caption).pack(side="left", padx=(4, 2))
                v = tk.StringVar(value=str(row.get(key, "")))
                var_map[key] = v
                tk.Entry(rf, textvariable=v, width=width).pack(side="left")
                if key == guard_field:
                    guard_var = v
            del_btn = tk.Button(
                rf, text="삭제", width=5,
                command=(lambda i=idx: _delete(i)),
            )
            del_btn.pack(side="left", padx=(4, 0))
            hint_lbl = tk.Label(rf, text="", fg="#888888", font=("맑은 고딕", 8))
            hint_lbl.pack(side="left", padx=(4, 0))

            if guard_var is not None:
                def _update_lock(*_args, v=guard_var, btn=del_btn, hint=hint_lbl) -> None:
                    locked = _is_required_column(v.get(), required_columns or set())
                    btn.configure(state=(tk.DISABLED if locked else tk.NORMAL))
                    hint.configure(text="(필수 컬럼 — 삭제 불가)" if locked else "")
                guard_var.trace_add("write", _update_lock)
                _update_lock()

            vars_list.append(var_map)

    def add() -> None:
        flush()
        rows_state.append(new_row())
        render()

    return frame, flush, render, add


def _show_theme_dialog(root: tk.Tk, theme: CompanyTheme, config: dict) -> CompanyTheme:
    """포인트 색상·로고·문구 설정 팝업 (탭 구성).

    [브랜드 설정] 탭: 포인트 색상 + 로고 최대 3개 슬롯.
    [문구 설정]   탭: PDF에 인쇄되는 모든 표시 문구(config["display_labels"]) 편집.

    Returns:
        [저장 후 계속] 클릭: 저장된 새 CompanyTheme (display_labels도 column_config.yaml에 저장)
        [취소] 또는 X 버튼: 원본 theme 반환 (저장 안 함)
        [기본값 복원] 클릭: GRAY_DEFAULT 저장 후 반환 (문구는 변경하지 않음)
    """
    result = [theme]
    current_hex = [theme.primary_hex]

    # 3슬롯 초기화: theme.logo_paths 에서 최대 3개, 부족분은 None 으로 패딩
    slots: list[str | None] = (list(theme.logo_paths[:3]) + [None, None, None])[:3]

    # display_labels 누락 방어 (정상 흐름에서는 load_config가 채워 줌)
    config.setdefault("display_labels", default_display_labels())

    # 필수 컬럼 삭제 가드용 — config["columns"] 매핑 기준으로 1회 계산 (다이얼로그가
    # 열려 있는 동안 columns 매핑은 바뀌지 않으므로 재계산 불필요).
    required_columns = _build_required_columns(config)

    dlg = tk.Toplevel(root)
    dlg.title("포인트 색상·로고·문구 설정")
    dlg.resizable(False, False)
    dlg.lift()
    dlg.grid_rowconfigure(0, weight=1)
    dlg.grid_columnconfigure(0, weight=1)

    notebook = ttk.Notebook(dlg)
    notebook.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 4))

    # ════════════════════════════════════════════════════════════════════
    #  탭 1: 브랜드 설정 (포인트 색상 + 로고)
    # ════════════════════════════════════════════════════════════════════
    brand = ttk.Frame(notebook)
    notebook.add(brand, text="브랜드 설정")

    # ── 색상 미리보기 레이블 ──────────────────────────────────────────────
    preview_frame = tk.Frame(brand, bd=1, relief="sunken")
    preview_frame.grid(row=0, column=0, columnspan=7, padx=12, pady=(12, 4), sticky="ew")

    swatch = tk.Label(preview_frame, bg=current_hex[0], width=8, height=2)
    swatch.pack(side="left", padx=4, pady=4)
    hex_label = tk.Label(preview_frame, text=current_hex[0], font=("Consolas", 10))
    hex_label.pack(side="left", padx=4)

    def _pick_color():
        chosen = colorchooser.askcolor(
            color=current_hex[0], parent=dlg, title="포인트 색상 선택"
        )
        if chosen[1]:
            current_hex[0] = chosen[1].upper()
            swatch.config(bg=current_hex[0])
            hex_label.config(text=current_hex[0])

    tk.Label(brand, text="섹션 헤더·상하단 바 색상:").grid(row=1, column=0, padx=12, pady=4, sticky="w")
    tk.Button(brand, text="색상 선택", command=_pick_color).grid(row=1, column=1, padx=4, pady=4)

    # ── 구분선 ────────────────────────────────────────────────────────────
    tk.Frame(brand, height=1, bg="#cccccc").grid(
        row=2, column=0, columnspan=7, sticky="ew", padx=12, pady=(4, 0)
    )

    # ── 로고 슬롯 (3행) ───────────────────────────────────────────────────
    tk.Label(brand, text="회사 로고 이미지 (최대 3개):").grid(
        row=3, column=0, columnspan=7, padx=12, pady=(8, 4), sticky="w"
    )

    slot_vars: list[tk.StringVar] = []
    scale_vars: list[tk.DoubleVar] = []

    def update_preview(*_args):  # no-op; PIL 블록에서 실제 함수로 교체됨
        pass

    for i in range(3):
        tk.Label(brand, text=f"로고 {i + 1}:", width=7, anchor="e").grid(
            row=4 + i, column=0, padx=(12, 4), pady=3, sticky="e"
        )

        var = tk.StringVar(value=_fmt_slot_label(slots[i]))
        slot_vars.append(var)
        tk.Label(brand, textvariable=var, font=("Consolas", 8), fg="#444444",
                 anchor="w", width=32, relief="sunken", padx=3).grid(
            row=4 + i, column=1, padx=2, pady=3, sticky="ew"
        )

        # default-argument 기법으로 클로저에서 i 고정
        def _make_pick(idx: int):
            def _pick():
                path = filedialog.askopenfilename(
                    parent=dlg,
                    title=f"로고 {idx + 1} 이미지 선택",
                    filetypes=[("이미지 파일", "*.png *.jpg *.jpeg"), ("모든 파일", "*.*")],
                )
                if path:
                    slots[idx] = path
                    slot_vars[idx].set(_fmt_slot_label(path))
                    update_preview()
            return _pick

        def _make_clear(idx: int):
            def _clear():
                slots[idx] = None
                slot_vars[idx].set("미지정")
                update_preview()
            return _clear

        tk.Button(brand, text="변경", width=6, command=_make_pick(i)).grid(
            row=4 + i, column=2, padx=2, pady=3
        )
        tk.Button(brand, text="삭제", width=6, command=_make_clear(i)).grid(
            row=4 + i, column=3, padx=(2, 8), pady=3
        )

        # ── 슬롯별 높이 슬라이더 (column 4~6) ───────────────────────────
        tk.Label(brand, text="높이:", anchor="e").grid(
            row=4 + i, column=4, padx=(8, 2), pady=3, sticky="e"
        )
        sv = tk.DoubleVar(value=theme.logo_heights[i])
        scale_vars.append(sv)
        val_lbl = tk.Label(brand, text=str(int(theme.logo_heights[i])), width=4, anchor="e")
        val_lbl.grid(row=4 + i, column=6, padx=(2, 12), pady=3, sticky="e")

        def _make_scale_cb(lbl):
            def _cb(val):
                lbl.config(text=str(int(float(val))))
            return _cb

        ttk.Scale(
            brand, from_=16, to=96, orient="horizontal",
            variable=sv, command=_make_scale_cb(val_lbl),
        ).grid(row=4 + i, column=5, padx=2, pady=3, sticky="ew")

    # ── 구분선 ────────────────────────────────────────────────────────────
    tk.Frame(brand, height=1, bg="#cccccc").grid(
        row=7, column=0, columnspan=7, sticky="ew", padx=12, pady=(8, 0)
    )

    # ── 로고 미리보기 캔버스 ─────────────────────────────────────────────
    _PREV_W, _PREV_H = 560, 100
    try:
        from PIL import Image, ImageTk  # optional; silently skipped if absent

        preview_canvas = tk.Canvas(
            brand, width=_PREV_W, height=_PREV_H,
            bg="#f5f5f5", highlightthickness=1,
            highlightbackground="#cccccc",
        )
        preview_canvas.grid(row=8, column=0, columnspan=7, padx=12, pady=(4, 12))

        def update_preview(*_args):  # 위 no-op 재바인딩
            preview_canvas.delete("all")
            preview_canvas._photos = {}
            x_offset = 4
            for i, path in enumerate(slots):
                if not path or not os.path.isfile(path):
                    continue
                logo_h_pt = int(scale_vars[i].get())
                logo_h_px = int(logo_h_pt * 96 / 72)
                try:
                    img = Image.open(path)
                    ratio = logo_h_px / img.height
                    new_w = int(img.width * ratio)
                    new_h = logo_h_px
                    if new_h > _PREV_H - 8:
                        ratio = (_PREV_H - 8) / img.height
                        new_w = int(img.width * ratio)
                        new_h = _PREV_H - 8
                    if new_w > _PREV_W - x_offset - 4:
                        ratio = (_PREV_W - x_offset - 4) / img.width
                        new_w = _PREV_W - x_offset - 4
                        new_h = int(img.height * ratio)
                    img = img.resize((max(1, new_w), max(1, new_h)), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    preview_canvas._photos[i] = photo
                    preview_canvas.create_image(x_offset, _PREV_H // 2, anchor="w", image=photo)
                    x_offset += new_w + 8
                except Exception:
                    pass

            if x_offset == 4:  # 표시된 로고 없음
                preview_canvas.create_text(
                    _PREV_W // 2, _PREV_H // 2,
                    text="로고 미리보기",
                    fill="#aaaaaa", justify="center", font=("맑은 고딕", 9),
                )

        for sv in scale_vars:
            sv.trace_add("write", update_preview)
        update_preview()
    except ImportError:
        pass

    # ════════════════════════════════════════════════════════════════════
    #  탭 2: 문구 설정 (PDF 표시 라벨)
    # ════════════════════════════════════════════════════════════════════
    labels_tab = ttk.Frame(notebook)
    notebook.add(labels_tab, text="문구 설정")

    _dl = config["display_labels"]

    # (그룹 제목, [(경로, 캡션), ...]) — 섹션 간 동일 텍스트는 캡션에 섹션을 명시한다.
    # 개수 가변 항목(계정 테이블 columns / 섹션1 rows / 섹션2 groups / 섹션3-1 rows /
    # 성격 컬럼 / 섹션4 rows)은 정적 목록이 아닌 동적 행 편집기로 렌더링한다.
    _label_groups: list[tuple[str, list[tuple[tuple, str]]]] = [
        ("상단", [
            (("header",), "제목 배너"),
        ]),
        ("계정 테이블", [
            (("account_table", "계정"), "계정 (병합 셀)"),
        ]),
        ("1. 대상정의", [
            (("section1", "제목"), "섹션 제목"),
        ]),
        ("2. 부점귀속 현황", [
            (("section2", "제목"), "섹션 제목"),
            (("section2", "귀속현황"), "귀속현황 (섹션2)"),
            (("section2", "부점명"), "부점명"),
            (("section2", "대상금액"), "대상금액 헤더"),
            (("section2", "구성비"), "구성비 헤더"),
            (("section2", "총계"), "부점귀속 총계 (섹션2)"),
            (("section2", "설명"), "설명 문구"),
        ]),
        ("3-1. 직접비 공통비 분류 결과", [
            (("section3_1", "제목"), "섹션 제목"),
            (("section3_1", "구분"), "구분"),
            (("section3_1", "분류금액"), "분류금액"),
        ]),
        # 성격 컬럼(nature_cols)은 개수가 가변이므로 아래 정적 목록이 아닌
        # 동적 편집 블록(_render_nature_rows)으로 렌더링한다.
        ("3-2. 성격별 분류 결과", [
            (("section3_2", "제목"), "섹션 제목"),
            (("section3_2", "귀속현황"), "귀속현황 (섹션3-2)"),
            (("section3_2", "합계"), "합계 헤더"),
            (("section3_2", "설명"), "설명 문구"),
        ]),
        # 4. 분류 근거는 정적 목록이 아닌 타입별 동적 행 편집기(_render_basis_rows)로
        # 렌더링한다. 섹션4 "제목"은 config 값으로 고정되며 이 탭에서 편집하지 않는다.
    ]

    def _dl_get(path: tuple):
        cur = _dl
        try:
            for k in path:
                cur = cur[k]
            return cur
        except (KeyError, IndexError, TypeError):
            return ""

    # ── 스크롤 가능한 라벨 편집 영역 ──────────────────────────────────────
    _lbl_canvas = tk.Canvas(labels_tab, width=600, height=380, highlightthickness=0)
    _lbl_vsb = tk.Scrollbar(labels_tab, orient="vertical", command=_lbl_canvas.yview)
    _lbl_inner = tk.Frame(_lbl_canvas)
    _lbl_inner.bind(
        "<Configure>",
        lambda e: _lbl_canvas.configure(scrollregion=_lbl_canvas.bbox("all")),
    )
    _lbl_canvas.create_window((0, 0), window=_lbl_inner, anchor="nw")
    _lbl_canvas.configure(yscrollcommand=_lbl_vsb.set)
    _lbl_canvas.pack(side="left", fill="both", expand=True)
    _lbl_vsb.pack(side="right", fill="y")

    # 마우스 휠 스크롤 — 탭에 들어왔을 때만 bind_all, 나갈 때 해제 (다이얼로그 간 누수 방지)
    def _on_wheel(event):
        _lbl_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    _lbl_canvas.bind("<Enter>", lambda e: _lbl_canvas.bind_all("<MouseWheel>", _on_wheel))
    _lbl_canvas.bind("<Leave>", lambda e: _lbl_canvas.unbind_all("<MouseWheel>"))

    # ── 성격 컬럼(3-2) 동적 편집 상태 — nature_cols(CSV명)=display명 통합 리스트 ──
    nature_vars: list[tk.StringVar] = [
        tk.StringVar(value=str(x)) for x in config.get("nature_cols", [])
    ]
    # 필수 컬럼 삭제 가드 스냅샷 — 다이얼로그 오픈 시점의 객체 참조를 보존한다
    # (저장 시 defense-in-depth 검증에서 사용; 편집은 in-place라 객체 동일성 유지됨).
    _nature_locked_snapshot = [v for v in nature_vars if _is_required_column(v.get(), required_columns)]
    nature_frame = tk.Frame(_lbl_inner)   # section3_2 블록 내부에 grid 됨
    add_var = tk.StringVar()

    def _render_nature_rows():
        for w in nature_frame.winfo_children():
            w.destroy()
        for idx, var in enumerate(nature_vars):
            tk.Label(nature_frame, text=f"성격 {idx + 1}:", width=8, anchor="e").grid(
                row=idx, column=0, padx=(8, 4), pady=2, sticky="e"
            )
            tk.Entry(nature_frame, textvariable=var, width=40).grid(
                row=idx, column=1, padx=(0, 4), pady=2, sticky="ew"
            )
            del_btn = tk.Button(
                nature_frame, text="삭제", width=5,
                command=(lambda i=idx: _delete_nature(i)),
            )
            del_btn.grid(row=idx, column=2, padx=(0, 4), pady=2)
            hint_lbl = tk.Label(nature_frame, text="", fg="#888888", font=("맑은 고딕", 8))
            hint_lbl.grid(row=idx, column=3, padx=(0, 8), pady=2, sticky="w")

            def _update_lock(*_args, v=var, btn=del_btn, hint=hint_lbl) -> None:
                locked = _is_required_column(v.get(), required_columns)
                btn.configure(state=(tk.DISABLED if locked else tk.NORMAL))
                hint.configure(text="(필수 컬럼 — 삭제 불가)" if locked else "")
            var.trace_add("write", _update_lock)
            _update_lock()

    def _delete_nature(idx: int):
        live_val = nature_vars[idx].get()
        if _is_required_column(live_val, required_columns):
            messagebox.showwarning(
                "삭제 불가",
                f"필수 컬럼(\"{live_val.strip()}\")은 삭제할 수 없습니다.",
                parent=dlg,
            )
            return
        if len(nature_vars) <= 1:
            messagebox.showwarning(
                "최소 1개", "성격 컬럼은 최소 1개 이상이어야 합니다.", parent=dlg
            )
            return
        del nature_vars[idx]
        _render_nature_rows()

    def _add_nature():
        name = add_var.get().strip()
        if not name:
            messagebox.showwarning("입력 필요", "추가할 성격 컬럼 이름을 입력하세요.", parent=dlg)
            return
        if name in [v.get().strip() for v in nature_vars]:
            messagebox.showwarning("중복", f"'{name}' 은(는) 이미 존재합니다.", parent=dlg)
            return
        nature_vars.append(tk.StringVar(value=name))
        add_var.set("")
        _render_nature_rows()

    # ── 섹션4 분류 근거(4) 동적 편집 상태 — 타입별 행 리스트 ──────────────────
    basis_rows_state: list[dict] = copy.deepcopy(
        config["display_labels"]["section4"].get("rows", [])
    )
    # 필수 컬럼 삭제 가드 스냅샷 — custom 타입 행의 csv_column만 대상 (직공통비/성격별분류는
    # 사용자 편집 가능한 csv_column 필드가 없음).
    _basis_locked_snapshot = [
        r for r in basis_rows_state
        if r.get("type") == "custom" and _is_required_column(str(r.get("csv_column", "")), required_columns)
    ]
    basis_vars: list[dict[str, tk.StringVar]] = []   # basis_rows_state와 인덱스 정렬
    basis_frame = tk.Frame(_lbl_inner)               # _label_groups 루프 후 grid 됨

    def _flush_basis_vars() -> None:
        """현재 Entry StringVar 값들을 basis_rows_state로 반영한다 (재렌더 전 호출).

        프레임을 부수고 다시 그리기 전에 호출하여 입력 중이던 값(라벨/CSV컬럼/참조 CSV 컬럼)이
        유실되지 않도록 한다.
        """
        for row, var_map in zip(basis_rows_state, basis_vars):
            for key, var in var_map.items():
                row[key] = var.get()

    def _render_basis_rows():
        for w in basis_frame.winfo_children():
            w.destroy()
        basis_vars.clear()
        for idx, row in enumerate(basis_rows_state):
            row_type = row.get("type", "custom")
            rf = tk.Frame(basis_frame)
            rf.grid(row=idx, column=0, sticky="w", padx=8, pady=2)
            var_map: dict[str, tk.StringVar] = {}

            tk.Label(rf, text=f"[{row_type}]", width=9, anchor="w",
                     fg="#2E2E38").pack(side="left")

            tk.Label(rf, text="라벨").pack(side="left", padx=(4, 2))
            v_label = tk.StringVar(value=str(row.get("label", "")))
            var_map["label"] = v_label
            tk.Entry(rf, textvariable=v_label, width=14).pack(side="left")

            v_csv: tk.StringVar | None = None
            if row_type == "custom":
                tk.Label(rf, text="근거 텍스트 컬럼").pack(side="left", padx=(4, 2))
                v_csv = tk.StringVar(value=str(row.get("csv_column", "")))
                var_map["csv_column"] = v_csv
                tk.Entry(rf, textvariable=v_csv, width=12).pack(side="left")

            tk.Label(rf, text="참조문서 컬럼").pack(side="left", padx=(4, 2))
            v_ref = tk.StringVar(value=str(row.get("참조_csv_column", "")))
            var_map["참조_csv_column"] = v_ref
            tk.Entry(rf, textvariable=v_ref, width=12).pack(side="left")

            del_btn = tk.Button(
                rf, text="삭제", width=5,
                command=(lambda i=idx: _delete_basis_row(i)),
            )
            del_btn.pack(side="left", padx=(4, 0))
            hint_lbl = tk.Label(rf, text="", fg="#888888", font=("맑은 고딕", 8))
            hint_lbl.pack(side="left", padx=(4, 0))

            if v_csv is not None:
                def _update_lock(*_args, v=v_csv, btn=del_btn, hint=hint_lbl) -> None:
                    locked = _is_required_column(v.get(), required_columns)
                    btn.configure(state=(tk.DISABLED if locked else tk.NORMAL))
                    hint.configure(text="(필수 컬럼 — 삭제 불가)" if locked else "")
                v_csv.trace_add("write", _update_lock)
                _update_lock()

            basis_vars.append(var_map)

    def _delete_basis_row(idx: int):
        live_csv = basis_vars[idx].get("csv_column") if idx < len(basis_vars) else None
        if live_csv is not None and _is_required_column(live_csv.get(), required_columns):
            messagebox.showwarning(
                "삭제 불가",
                f"필수 컬럼(\"{live_csv.get().strip()}\")은 삭제할 수 없습니다.",
                parent=dlg,
            )
            return
        _flush_basis_vars()
        del basis_rows_state[idx]
        _render_basis_rows()

    def _add_basis_row(row_type: str):
        _flush_basis_vars()
        if row_type == "직공통비":
            basis_rows_state.append(
                {"type": "직공통비", "label": "직 • 공통비", "참조_csv_column": ""}
            )
        elif row_type == "성격별분류":
            basis_rows_state.append(
                {"type": "성격별분류", "label": "성격별 분류", "참조_csv_column": ""}
            )
        else:
            basis_rows_state.append(
                {"type": "custom", "label": "", "csv_column": "", "참조_csv_column": ""}
            )
        _render_basis_rows()

    # ── 계정 테이블/섹션1/섹션2/섹션3-1 동적 행 편집기 (_make_rows_editor 공용) ──
    acct_cols_state: list[dict] = copy.deepcopy(_dl["account_table"].get("columns", []))
    sec1_rows_state: list[dict] = copy.deepcopy(_dl["section1"].get("rows", []))
    sec2_groups_state: list[dict] = copy.deepcopy(_dl["section2"].get("groups", []))
    sec31_rows_state: list[dict] = copy.deepcopy(_dl["section3_1"].get("rows", []))

    # 필수 컬럼 삭제 가드 스냅샷 — 다이얼로그 오픈 시점의 객체 참조를 보존한다 (저장 시
    # defense-in-depth 검증용). 섹션3-1은 keyword가 컬럼명이 아니라 가드 대상이 아님.
    _acct_locked_snapshot = [
        r for r in acct_cols_state if _is_required_column(str(r.get("csv_column", "")), required_columns)
    ]
    _sec1_locked_snapshot = [
        r for r in sec1_rows_state if _is_required_column(str(r.get("csv_column", "")), required_columns)
    ]
    _sec2_locked_snapshot = [
        r for r in sec2_groups_state if _is_required_column(str(r.get("csv_column", "")), required_columns)
    ]

    acct_frame, acct_flush, acct_render, acct_add = _make_rows_editor(
        _lbl_inner, dlg, acct_cols_state,
        [("label", "라벨", 14), ("csv_column", "계정정보 컬럼", 16)],
        lambda: {"label": "", "csv_column": ""},
        section_name="계정 테이블 컬럼",
        guard_field="csv_column", required_columns=required_columns,
    )
    sec1_frame, sec1_flush, sec1_render, sec1_add = _make_rows_editor(
        _lbl_inner, dlg, sec1_rows_state,
        [("label", "라벨", 14), ("csv_column", "계정정보 컬럼", 16)],
        lambda: {"label": "", "csv_column": ""},
        section_name="대상정의 행",
        guard_field="csv_column", required_columns=required_columns,
    )
    sec2_frame, sec2_flush, sec2_render, sec2_add = _make_rows_editor(
        _lbl_inner, dlg, sec2_groups_state,
        [("label", "라벨", 16), ("csv_column", "집계 기준 컬럼", 14)],
        lambda: {"label": "", "csv_column": ""},
        section_name="귀속 그룹",
        guard_field="csv_column", required_columns=required_columns,
    )
    sec31_frame, sec31_flush, sec31_render, sec31_add = _make_rows_editor(
        _lbl_inner, dlg, sec31_rows_state,
        [("label", "라벨", 14), ("keyword", "매칭 키워드", 14)],
        lambda: {"label": "", "keyword": ""},
        section_name="분류 행",
    )

    label_vars: dict[tuple, tk.StringVar] = {}

    def _mount_rows_editor(r: int, hint: str, frame: tk.Frame, render, add, add_label: str) -> int:
        """그룹 정적 필드 아래에 동적 행 편집기 블록(힌트/행/추가 버튼)을 배치한다."""
        tk.Label(
            _lbl_inner, text=hint,
            font=("맑은 고딕", 8), anchor="w", fg="#c0392b",
            wraplength=560, justify="left",
        ).grid(row=r, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 2))
        r += 1
        frame.grid(row=r, column=0, columnspan=2, sticky="w")
        r += 1
        render()
        btn_row = tk.Frame(_lbl_inner)
        btn_row.grid(row=r, column=0, columnspan=2, sticky="w", padx=8, pady=(2, 6))
        r += 1
        tk.Button(btn_row, text=add_label, command=add).pack(side="left")
        return r

    # 그룹 제목 → 그 그룹의 정적 필드 아래에 붙일 동적 편집기 마운트 콜백
    _editor_mounts = {
        "계정 테이블": lambda r: _mount_rows_editor(
            r, "※ 계정정보.csv에 없는 컬럼은 빈 칸으로 출력됩니다.",
            acct_frame, acct_render, acct_add, "+ 컬럼 추가"),
        "1. 대상정의": lambda r: _mount_rows_editor(
            r, "※ 계정정보.csv에 없는 컬럼은 빈 칸으로 출력됩니다.",
            sec1_frame, sec1_render, sec1_add, "+ 행 추가"),
        "2. 부점귀속 현황": lambda r: _mount_rows_editor(
            r, "※ 집계 기준 컬럼명 (주관부서=사업비정보 라인별, 팀=코스트센터→부서정보). "
               "각 그룹은 해당 컬럼 기준 전체 롤업으로 집계됩니다.",
            sec2_frame, sec2_render, sec2_add, "+ 그룹 추가"),
        "3-1. 직접비 공통비 분류 결과": lambda r: _mount_rows_editor(
            r, "※ 계정정보.csv 직간접구분 값 부분일치. 위에서부터 먼저 일치한 분류로 "
               "집계됩니다. (섹션4 '직공통비' 근거 문구와는 별개입니다.)",
            sec31_frame, sec31_render, sec31_add, "+ 분류 추가"),
    }

    _r = 0
    for group_title, fields in _label_groups:
        tk.Label(
            _lbl_inner, text=group_title, font=("맑은 고딕", 10, "bold"),
            anchor="w", fg="#2E2E38",
        ).grid(row=_r, column=0, columnspan=2, sticky="w", padx=8, pady=(10, 2))
        _r += 1
        for path, caption in fields:
            tk.Label(_lbl_inner, text=caption, width=20, anchor="e").grid(
                row=_r, column=0, padx=(8, 4), pady=2, sticky="e"
            )
            v = tk.StringVar(value=str(_dl_get(path)))
            label_vars[path] = v
            tk.Entry(_lbl_inner, textvariable=v, width=56).grid(
                row=_r, column=1, padx=(0, 12), pady=2, sticky="ew"
            )
            _r += 1

        # 그룹 정적 필드 직후에 동적 행 편집기 삽입 (계정/섹션1/2/3-1)
        _mount = _editor_mounts.get(group_title)
        if _mount is not None:
            _r = _mount(_r)

        # 3-2 그룹 직후에 성격 컬럼 동적 편집 블록을 삽입한다.
        # (성격 편집기는 str 리스트 상태 + 입력창 딸린 추가 버튼이라 별도 구현 유지)
        if group_title == "3-2. 성격별 분류 결과":
            tk.Label(
                _lbl_inner, text="성격 컬럼 (CSV 컬럼명 = 표시명, 3-2 표)",
                font=("맑은 고딕", 9, "bold"), anchor="w", fg="#2E2E38",
            ).grid(row=_r, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 0))
            _r += 1
            tk.Label(
                _lbl_inner,
                text="※ CSV 컬럼명과 표시명에 동일하게 적용됩니다. "
                     "사업비정보.csv에 없는 이름은 0으로 출력됩니다.",
                font=("맑은 고딕", 8), anchor="w", fg="#c0392b",
                wraplength=560, justify="left",
            ).grid(row=_r, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 2))
            _r += 1
            nature_frame.grid(row=_r, column=0, columnspan=2, sticky="w")
            _r += 1
            _render_nature_rows()
            add_row = tk.Frame(_lbl_inner)
            add_row.grid(row=_r, column=0, columnspan=2, sticky="w", padx=8, pady=(2, 6))
            _r += 1
            tk.Entry(add_row, textvariable=add_var, width=40).pack(side="left", padx=(0, 4))
            tk.Button(add_row, text="+ 추가", width=6, command=_add_nature).pack(side="left")

    # ── 4. 분류 근거 동적 행 편집기 (정적 _label_groups 대신) ────────────────
    tk.Label(
        _lbl_inner, text="4. 분류 근거", font=("맑은 고딕", 10, "bold"),
        anchor="w", fg="#2E2E38",
    ).grid(row=_r, column=0, columnspan=2, sticky="w", padx=8, pady=(10, 2))
    _r += 1
    tk.Label(
        _lbl_inner,
        text="※ 직공통비/성격별분류 행은 기존 분류 로직을, custom 행은 출력.csv의 지정 컬럼 "
             "첫 값을 사용합니다. 행을 삭제하면 PDF에서도 제외됩니다.",
        font=("맑은 고딕", 8), anchor="w", fg="#c0392b",
        wraplength=560, justify="left",
    ).grid(row=_r, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 2))
    _r += 1
    basis_frame.grid(row=_r, column=0, columnspan=2, sticky="w")
    _r += 1
    _render_basis_rows()
    basis_add_row = tk.Frame(_lbl_inner)
    basis_add_row.grid(row=_r, column=0, columnspan=2, sticky="w", padx=8, pady=(2, 8))
    _r += 1
    tk.Button(basis_add_row, text="+ 직공통비 추가",
              command=lambda: _add_basis_row("직공통비")).pack(side="left", padx=(0, 4))
    tk.Button(basis_add_row, text="+ 성격별분류 추가",
              command=lambda: _add_basis_row("성격별분류")).pack(side="left", padx=(0, 4))
    tk.Button(basis_add_row, text="+ 커스텀 행 추가",
              command=lambda: _add_basis_row("custom")).pack(side="left")

    def _dl_set(path: tuple, value: str) -> None:
        cur = config["display_labels"]
        for k in path[:-1]:
            cur = cur[k]
        cur[path[-1]] = value

    # ════════════════════════════════════════════════════════════════════
    #  버튼 행 (노트북 하단)
    # ════════════════════════════════════════════════════════════════════
    btn_frame = tk.Frame(dlg)
    btn_frame.grid(row=1, column=0, pady=(4, 12), padx=12)

    def _save_and_close():
        # 0) 성격 컬럼 검증 — 실패 시 아무것도 저장하지 않고 대화상자를 유지한다.
        new_nature = [v.get().strip() for v in nature_vars]
        if not new_nature:
            messagebox.showwarning(
                "최소 1개", "성격 컬럼은 최소 1개 이상이어야 합니다.", parent=dlg
            )
            return
        if any(not n for n in new_nature):
            messagebox.showwarning(
                "입력 오류", "비어 있는 성격 컬럼 이름이 있습니다.", parent=dlg
            )
            return
        if len(set(new_nature)) != len(new_nature):
            messagebox.showwarning(
                "중복 오류", "성격 컬럼 이름이 중복됩니다.", parent=dlg
            )
            return

        # 0-1) 섹션4 분류 근거 행 검증 — 입력 중인 값을 먼저 반영한 뒤 검사한다.
        _flush_basis_vars()
        for row in basis_rows_state:
            if not str(row.get("label", "")).strip():
                messagebox.showwarning(
                    "입력 오류", "분류 근거 행의 라벨은 비울 수 없습니다.", parent=dlg
                )
                return
            if row.get("type") == "custom" and not str(row.get("csv_column", "")).strip():
                messagebox.showwarning(
                    "입력 오류",
                    "커스텀 분류 근거 행에는 CSV 컬럼명을 입력해야 합니다.",
                    parent=dlg,
                )
                return

        # 0-2) 동적 행 편집기(계정 테이블/섹션1/2/3-1) 검증 — 실패 시 저장 중단.
        # (섹션명, flush, 상태 리스트, 라벨 외 필수 키, 그 키의 표시명)
        _dyn_sections = [
            ("계정 테이블 컬럼",   acct_flush,  acct_cols_state,   "csv_column", "계정정보 컬럼명"),
            ("대상정의(1) 행",     sec1_flush,  sec1_rows_state,   "csv_column", "계정정보 컬럼명"),
            ("부점귀속(2) 그룹",   sec2_flush,  sec2_groups_state, "csv_column", "출력.csv 컬럼명"),
            ("분류(3-1) 행",       sec31_flush, sec31_rows_state,  "keyword",    "매칭 키워드"),
        ]
        for sec_name, flush, state, key_field, key_caption in _dyn_sections:
            flush()
            if not state:
                messagebox.showwarning(
                    "최소 1개", f"{sec_name}은(는) 최소 1개 이상이어야 합니다.", parent=dlg
                )
                return
            for row in state:
                if not str(row.get("label", "")).strip():
                    messagebox.showwarning(
                        "입력 오류", f"{sec_name}의 라벨은 비울 수 없습니다.", parent=dlg
                    )
                    return
                if not str(row.get(key_field, "")).strip():
                    messagebox.showwarning(
                        "입력 오류",
                        f"{sec_name}의 {key_caption}은(는) 비울 수 없습니다.",
                        parent=dlg,
                    )
                    return
            keys = [str(row.get(key_field, "")).strip() for row in state]
            if len(set(keys)) != len(keys):
                messagebox.showwarning(
                    "중복 오류",
                    f"{sec_name}에 중복된 {key_caption}이(가) 있습니다.",
                    parent=dlg,
                )
                return
        # 검증 통과 후 문자열 값 공백 정리 (CSV 컬럼 조회가 공백에 민감)
        for _, _, state, _, _ in _dyn_sections:
            for row in state:
                for k, v in row.items():
                    if isinstance(v, str):
                        row[k] = v.strip()

        # 0-3) 필수 컬럼 삭제 방어 (defense in depth) — 삭제 가드가 정상 동작했다면
        # 이 검사는 항상 통과한다. 다이얼로그 오픈 시점에 필수 컬럼과 일치했던 행(객체)이
        # 현재 리스트에서 사라졌는지 객체 동일성으로 확인한다 (값 편집만으로는 사라지지
        # 않음 — flush/render는 같은 객체를 in-place로 변형하므로).
        _locked_snapshots = [
            ("계정 테이블 컬럼", _acct_locked_snapshot, acct_cols_state),
            ("대상정의(1) 행", _sec1_locked_snapshot, sec1_rows_state),
            ("부점귀속(2) 그룹", _sec2_locked_snapshot, sec2_groups_state),
            ("성격 컬럼(3-2)", _nature_locked_snapshot, nature_vars),
            ("분류 근거(4) 커스텀 행", _basis_locked_snapshot, basis_rows_state),
        ]
        for sec_name, snapshot, current_list in _locked_snapshots:
            missing = [row for row in snapshot if row not in current_list]
            if missing:
                messagebox.showwarning(
                    "삭제 불가",
                    f"{sec_name}의 필수 컬럼 항목이 삭제되었습니다.\n"
                    "편집을 취소하거나 다시 추가한 뒤 저장하세요.",
                    parent=dlg,
                )
                return

        valid_paths = [p for p in slots if p]
        new_theme = CompanyTheme(
            primary_hex=current_hex[0],
            logos=theme.logos,
            logo_paths=valid_paths,
            logo_max_height=max(int(v.get()) for v in scale_vars),
            logo_heights=[int(v.get()) for v in scale_vars],
        )
        save_theme(new_theme)

        # 편집된 표시 문구를 config에 반영 후 column_config.yaml에 저장
        for path, var in label_vars.items():
            _dl_set(path, var.get())
        # 성격 컬럼: nature_cols(CSV명) = display명 통합 리스트로 양쪽에 동일 반영
        config["nature_cols"] = new_nature
        config["display_labels"]["section3_2"]["nature_cols"] = list(new_nature)
        # 섹션4 분류 근거 행 설정 반영
        config["display_labels"]["section4"]["rows"] = basis_rows_state
        # 계정 테이블/섹션1/2/3-1 동적 리스트 반영
        config["display_labels"]["account_table"]["columns"] = acct_cols_state
        config["display_labels"]["section1"]["rows"] = sec1_rows_state
        config["display_labels"]["section2"]["groups"] = sec2_groups_state
        config["display_labels"]["section3_1"]["rows"] = sec31_rows_state
        if not save_config(config):
            messagebox.showwarning(
                "문구 저장 실패",
                "표시 문구를 column_config.yaml에 저장하지 못했습니다.\n"
                "(쓰기 권한을 확인하세요. 색상·로고 설정은 정상 저장되었습니다.)",
                parent=dlg,
            )

        result[0] = new_theme
        dlg.destroy()

    def _reset_default():
        save_theme(GRAY_DEFAULT)
        result[0] = GRAY_DEFAULT
        dlg.destroy()

    tk.Button(btn_frame, text="저장 후 계속", width=12, command=_save_and_close).pack(side="left", padx=4)
    tk.Button(btn_frame, text="취소",         width=8,  command=dlg.destroy).pack(side="left", padx=4)
    tk.Button(btn_frame, text="기본값 복원",  width=10, command=_reset_default).pack(side="left", padx=4)

    dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
    dlg.grab_set()
    dlg.wait_window()
    return result[0]


# 입력 파일 선택 다이얼로그 행 정의: (논리 키, 표시 라벨)
_FILE_SELECT_ROWS: list[tuple[str, str]] = [
    ("transaction", "사업비정보"),
    ("account",     "계정정보"),
    ("ccm",         "부서정보"),
    ("output",      "출력"),
]


def _show_file_select_dialog(root: tk.Tk) -> dict | None:
    """4개 입력 CSV 파일을 각각 선택받는 모달 다이얼로그.

    Returns:
        {"transaction":경로, "account":경로, "ccm":경로, "output":경로} — [확인] 시
        None — [취소] 또는 X 버튼 시
    """
    result: list[dict | None] = [None]

    dlg = tk.Toplevel(root)
    dlg.title("입력 파일 선택")
    dlg.resizable(False, False)

    tk.Label(
        dlg,
        text="아래 4개 CSV 파일을 각각 선택해 주세요.",
        font=("맑은 고딕", 10, "bold"),
    ).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 6))

    last_paths = _load_last_paths()
    path_vars: dict[str, tk.StringVar] = {}

    for i, (key, label) in enumerate(_FILE_SELECT_ROWS):
        tk.Label(dlg, text=f"{label}:", width=10, anchor="e").grid(
            row=1 + i, column=0, padx=(12, 4), pady=4, sticky="e"
        )
        saved = last_paths.get(key, "")
        initial = saved if saved and Path(saved).is_file() else ""
        var = tk.StringVar(value=initial)
        path_vars[key] = var
        tk.Entry(
            dlg, textvariable=var, width=58, state="readonly",
        ).grid(row=1 + i, column=1, padx=4, pady=4, sticky="ew")

        def _make_pick(k: str, lbl: str):
            def _pick():
                path = filedialog.askopenfilename(
                    parent=dlg,
                    title=f"{lbl} CSV 선택",
                    filetypes=[("CSV 파일", "*.csv"), ("모든 파일", "*.*")],
                )
                if path:
                    path_vars[k].set(path)
            return _pick

        tk.Button(dlg, text="찾아보기", width=10, command=_make_pick(key, label)).grid(
            row=1 + i, column=2, padx=(4, 12), pady=4
        )

    # ── 버튼 행 ──────────────────────────────────────────────────────────
    btn_frame = tk.Frame(dlg)
    btn_frame.grid(row=1 + len(_FILE_SELECT_ROWS), column=0, columnspan=3, pady=(8, 12))

    def _on_confirm():
        selected = {k: v.get().strip() for k, v in path_vars.items()}
        if not all(selected.values()):
            messagebox.showwarning("입력 필요", "모든 파일을 선택해주세요.", parent=dlg)
            return
        _save_last_paths(selected)
        result[0] = selected
        dlg.destroy()

    def _on_cancel():
        result[0] = None
        dlg.destroy()

    tk.Button(btn_frame, text="확인", width=12, command=_on_confirm).pack(side="left", padx=6)
    tk.Button(btn_frame, text="취소", width=8, command=_on_cancel).pack(side="left", padx=6)

    dlg.protocol("WM_DELETE_WINDOW", _on_cancel)

    # PyInstaller 환경 포커스 보장 + 중앙 배치
    dlg.update_idletasks()
    _w, _h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
    _sw, _sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    dlg.geometry(f"{_w}x{_h}+{(_sw - _w) // 2}+{(_sh - _h) // 2}")
    dlg.grab_set()
    dlg.lift()
    dlg.focus_force()

    dlg.wait_window()
    return result[0]


def show_unregistered_dialog(parent: tk.Tk, teams: list[str]) -> bool:
    """미등록 부서 감지 경고 대화상자를 표시한다.

    출력.csv에 등록되지 않은 부서 목록을 보여주고 사용자의 선택을 반환한다.
    부수효과(root.destroy / sys.exit)는 호출부에서 처리하도록 분리한다.

    Returns:
        True:  '무시하고 진행' 선택 (집계 누락을 감수하고 계속)
        False: '종료 후 수정' 또는 창 닫기 (프로그램을 종료해야 함)
    """
    _dept_lines = "\n".join(f"  {chr(8226)} {t}" for t in teams)
    _proceed = False

    dlg = tk.Toplevel(parent)
    dlg.title("미등록 부서 감지")
    dlg.resizable(False, False)

    def _refocus(event=None):
        dlg.focus_force()
    dlg.bind("<FocusOut>", _refocus)
    dlg.focus_force()

    # 상단 안내 문구
    _msg = (
        "실제 발생 내역에 존재하지만 '출력.csv'에 등록되지 않은 부서가 발견되었습니다.\n"
        "이대로 진행하면 해당 부서의 실적은 집계에서 누락됩니다."
    )
    tk.Label(
        dlg,
        text=_msg,
        wraplength=440,
        justify="left",
        fg="#c0392b",
        font=("Malgun Gothic", 10, "bold"),
        padx=16, pady=14,
    ).pack(fill="x")

    # 스크롤 가능한 부서 목록
    frame_list = tk.Frame(dlg, padx=16, pady=4)
    frame_list.pack(fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame_list, orient="vertical")
    txt = tk.Text(
        frame_list,
        height=10,
        width=50,
        yscrollcommand=scrollbar.set,
        font=("Malgun Gothic", 10),
        relief="sunken",
        bd=1,
    )
    scrollbar.config(command=txt.yview)
    scrollbar.pack(side="right", fill="y")
    txt.pack(side="left", fill="both", expand=True)
    txt.insert("1.0", _dept_lines)
    txt.config(state="disabled")

    # 하단 버튼
    frame_btn = tk.Frame(dlg, padx=16, pady=12)
    frame_btn.pack(fill="x")

    def _on_exit() -> None:
        dlg.destroy()

    def _on_proceed() -> None:
        nonlocal _proceed
        _proceed = True
        dlg.destroy()

    dlg.protocol("WM_DELETE_WINDOW", _on_exit)

    tk.Button(
        frame_btn, text="종료 후 수정",
        width=18, command=_on_exit,
    ).pack(side="left", padx=(0, 8))
    tk.Button(
        frame_btn, text="무시하고 진행",
        width=18, command=_on_proceed,
    ).pack(side="left")

    dlg.update_idletasks()
    _w = dlg.winfo_reqwidth()
    _h = dlg.winfo_reqheight()
    _sw = dlg.winfo_screenwidth()
    _sh = dlg.winfo_screenheight()
    dlg.geometry(f"{_w}x{_h}+{(_sw - _w) // 2}+{(_sh - _h) // 2}")

    parent.wait_window(dlg)
    return _proceed


def main() -> None:
    """
    프로그램 전체 흐름을 제어하는 메인 함수.

    실행 순서:
        1. tkinter root 초기화 (hide) — --noconsole 환경에서 대화상자만 표시
        2. 입력 CSV 파일 4개 선택 (_show_file_select_dialog) + 파일 검증
        3. 결과 저장 폴더 선택 (askdirectory)
        4. Phase 0 클렌징 포함 데이터 로드 (data_loader.load_all_csvs)
        5. 출력전표 코드 목록 추출 (출력.csv)
        6. 코드별 파이프라인 루프 (processor → pdf_exporter)
        7. 완료 요약 메시지
    """
    # ── tkinter 초기화 (대화상자 전용 — 메인 창 숨김) ─────────────────────
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # ── Step 0-A: config 로드 ─────────────────────────────────────────────
    config = load_config()

    # ── Step 0: 테마 로드 + 선택적 색상 설정 ──────────────────────────────
    theme = load_theme()   # config 없거나 손상 시 GRAY_DEFAULT 자동 적용
    if messagebox.askyesno(
        "포인트 색상 설정",
        "PDF 리포트의 포인트 색상 및 로고를 변경하시겠습니까?\n\n"
        f"현재 색상: {theme.primary_hex}\n"
        "('아니오' 선택 시 현재 설정으로 바로 진행합니다.)",
    ):
        theme = _show_theme_dialog(root, theme, config)

    # ── Step 1: 입력 CSV 파일 4개 선택 ────────────────────────────────────
    file_paths = _show_file_select_dialog(root)
    if file_paths is None:
        messagebox.showinfo("취소", "파일 선택이 취소되었습니다. 프로그램을 종료합니다.")
        sys.exit(0)

    # ── Step 1-B: 선택한 파일의 존재 + 필수 컬럼 검증 ─────────────────────
    problems = validate_csv_columns(file_paths, config)
    if problems:
        detail = "\n".join(problems)
        messagebox.showerror(
            "입력 파일 오류",
            "선택한 파일을 읽지 못했거나 필수 컬럼이 없습니다.\n\n"
            f"{detail}",
        )
        sys.exit(1)

    # -- Step 1-C: 미등록 부서 사전 검증 (차단 또는 무시) --------------------
    _unregistered = data_loader.check_unregistered_teams(file_paths, config)
    if _unregistered:
        if not show_unregistered_dialog(root, _unregistered):
            root.destroy()
            sys.exit(0)

    # ── Step 3: 데이터 로드 (Phase 0 클렌징 포함) ─────────────────────────
    try:
        sheets = data_loader.load_all_csvs(file_paths, config=config)
    except Exception as exc:
        _names = "·".join(
            Path(file_paths[k]).name for k in ("transaction", "account", "ccm", "output")
        )
        messagebox.showerror(
            "데이터 로드 오류",
            f"입력 CSV를 읽는 중 오류가 발생했습니다.\n\n"
            f"사유: {exc}\n\n"
            f"선택한 파일({_names})과 컬럼 구성을 확인하세요.",
        )
        sys.exit(1)

    # ── 설정 vs 실제 CSV 검사 — 비차단 경고 (전 코드 공통, 1회 계산) ──────────
    _dl_cfg = config.get("display_labels", {})

    def _cfg_entries(sec_key: str, list_key: str) -> list[dict]:
        sec = _dl_cfg.get(sec_key, {})
        entries = sec.get(list_key, []) if isinstance(sec, dict) else []
        return [e for e in entries if isinstance(e, dict)]

    # 설정된 성격 컬럼 중 사업비정보.csv에 실제로 없는 컬럼
    unknown_nature = [
        c for c in data_loader.NATURE_COLS
        if c not in sheets[data_loader._SHEET_TRANSACTION].columns
    ]

    # 계정 테이블/섹션1 의 csv_column 중 계정정보.csv에 없는 컬럼
    # (JOIN 전 원본 시트 기준이므로 계정번호도 그대로 존재 — 별칭 불필요)
    _acct_available = set(sheets[data_loader._SHEET_ACCOUNT].columns)
    unknown_account = [
        c for c in dict.fromkeys(
            e.get("csv_column", "")
            for e in _cfg_entries("account_table", "columns") + _cfg_entries("section1", "rows")
        )
        if c and c not in _acct_available
    ]

    # 섹션2 귀속 그룹의 csv_column 중 집계 소스 어디에도 없는 컬럼
    # (주관부서=사업비정보 / 팀=부서정보 / 그 밖=출력.csv — enrich 후 df_enriched 컬럼 기준)
    _group_available = (
        set(sheets[data_loader._SHEET_TRANSACTION].columns)
        | set(sheets[data_loader._SHEET_CCM].columns)
        | set(sheets[_SHEET_OUTPUT].columns)
    )
    unknown_group = [
        c for c in dict.fromkeys(
            e.get("csv_column", "") for e in _cfg_entries("section2", "groups")
        )
        if c and c not in _group_available
    ]

    # 섹션3-1 키워드 중 계정정보.csv 직간접구분 값에 전혀 일치하지 않는 키워드
    _gubun_col = COLUMN_MAP["직간접구분"]
    _df_account = sheets[data_loader._SHEET_ACCOUNT]
    unmatched_keywords: list[str] = []
    if _gubun_col in _df_account.columns:
        _gubun_series = _df_account[_gubun_col].astype(str)
        for e in _cfg_entries("section3_1", "rows"):
            kw = str(e.get("keyword", ""))
            if kw and not _gubun_series.str.contains(kw, na=False, regex=False).any():
                unmatched_keywords.append(kw)

    # ── 3-1 직간접구분 매핑 누락 검출 (조용한 오분류 방지) ──────────────────────
    # 두 원인 모두 enrich 후 "(미매칭)"으로 채워져 직접비/공통비 어디에도 집계되지
    # 않으면서 총계에는 포함된다("분류 합 ≠ 총계"). 실제 값은 강제하지 않고 목록만 알린다.
    _acct_col = COLUMN_MAP["계정번호"]
    _tx_col   = COLUMN_MAP["원가요소"]
    _df_tx    = sheets[data_loader._SHEET_TRANSACTION]

    # 케이스 1: 계정정보에는 있으나 직간접구분 값이 빈 계정번호
    unmapped_gubun_accounts: list[str] = []
    if _gubun_col in _df_account.columns and _acct_col in _df_account.columns:
        _g = _df_account[_gubun_col]
        _blank = _g.isna() | _g.astype(str).str.strip().isin(
            ["", "nan", "None", data_loader.FALLBACK_TEXT]
        )
        unmapped_gubun_accounts = (
            _df_account.loc[_blank, _acct_col].astype(str).str.strip()
            .loc[lambda s: s != ""].drop_duplicates().tolist()
        )

    # 케이스 2: 사업비정보 원가요소가 계정정보에 아예 없음 (계정 키 차집합)
    missing_account_codes: list[str] = []
    if _tx_col in _df_tx.columns and _acct_col in _df_account.columns:
        _acct_keys = set(_df_account[_acct_col].astype(str).str.strip())
        missing_account_codes = (
            _df_tx[_tx_col].astype(str).str.strip()
            .loc[lambda s: (s != "") & ~s.isin(_acct_keys)]
            .drop_duplicates().tolist()
        )

    # (경고 섹션 제목, 부연 설명, 항목 리스트) — _show_summary_dialog 가 ⚠️ 블록으로 렌더
    config_warnings: list[tuple[str, str, list[str]]] = []
    if unknown_nature:
        config_warnings.append((
            "미등록 성격 컬럼",
            "사업비정보.csv에 없는 컬럼 — 해당 컬럼은 빈 칸으로 출력됩니다",
            unknown_nature,
        ))
    if unknown_account:
        config_warnings.append((
            "미등록 계정정보 컬럼 (계정 테이블/섹션1)",
            "계정정보.csv에 없는 컬럼 — 해당 항목은 빈 칸으로 출력됩니다",
            unknown_account,
        ))
    if unknown_group:
        config_warnings.append((
            "미등록 부점귀속 그룹 컬럼 (섹션2)",
            "사업비정보/부서정보/출력.csv 어디에도 없는 컬럼 — 해당 그룹은 빈 표로 출력됩니다",
            unknown_group,
        ))
    if unmatched_keywords:
        config_warnings.append((
            "일치 계정 없는 분류 키워드 (섹션3-1)",
            "직간접구분 값에 일치하는 계정이 없어 0원으로 출력됩니다",
            unmatched_keywords,
        ))
    if missing_account_codes:
        config_warnings.append((
            "계정정보에 없는 원가요소 (섹션3-1)",
            "계정정보.csv에 매칭 계정번호가 없어 직간접구분을 못 얻습니다 — "
            "해당 금액이 직접비/공통비 어디에도 집계되지 않습니다 (총계에는 포함)",
            missing_account_codes,
        ))
    if unmapped_gubun_accounts:
        config_warnings.append((
            "직간접구분 미입력 계정번호 (섹션3-1)",
            "계정정보.csv에 있으나 직간접구분 값이 비어 있어 해당 계정 금액이 "
            "직접비/공통비 어디에도 집계되지 않습니다 (총계에는 포함)",
            unmapped_gubun_accounts,
        ))

    # ── Step 5: 출력전표 코드 목록 수집 (출력.csv의 출력전표 컬럼에서) ──────────
    col_output_code = COLUMN_MAP["출력전표"]
    df_output = sheets[_SHEET_OUTPUT]
    codes = (
        df_output[col_output_code].dropna().astype(str).str.strip().unique().tolist()
        if col_output_code in df_output.columns else []
    )
    codes = [c for c in codes if c]
    if not codes:
        messagebox.showerror(
            "데이터 오류",
            f"'출력.csv'의 '{col_output_code}' 컬럼에 처리할 출력전표 코드가 없습니다.",
        )
        sys.exit(1)

    # ── Step 6: 코드별 파이프라인 루프 ────────────────────────────────────
    successes: list[str] = []
    errors: list[tuple[str, str]] = []
    _pipeline_results: dict[str, dict] = {}

    # ── Stage 1: 파이프라인 전체 실행 및 결과 수집 ────────────────────────
    for code in codes:
        code_str = str(code).strip()

        try:
            results = processor.run_pipeline(sheets, code_str, labels=config["display_labels"])
        except ValueError as exc:
            errors.append((code_str, str(exc)))
            continue
        except Exception as exc:
            errors.append((code_str, f"파이프라인 오류: {exc}"))
            continue

        _pipeline_results[code_str] = results

    # ── Step 2: 결과 저장 폴더 선택 ───────────────────────────────────────
    out_dir = filedialog.askdirectory(
        title="결과 파일 저장 폴더 선택",
    )
    if not out_dir:
        messagebox.showinfo("취소", "저장 폴더 선택이 취소되었습니다. 프로그램을 종료합니다.")
        sys.exit(0)

    # ── Stage 3: PDF 생성 ─────────────────────────────────────────────────
    for code_str, results in _pipeline_results.items():
        pdf_path = os.path.join(out_dir, _PDF_FILENAME_TEMPLATE.format(code_str))
        ok = pdf_exporter.export(
            results, pdf_path, code_str, theme=theme,
            logo_heights=theme.logo_heights, labels=config["display_labels"],
        )
        if not ok:
            errors.append((code_str, "PDF 생성 실패"))
            continue

        successes.append(code_str)

    # ── Step 7: 완료 요약 ──────────────────────────────────────────────────
    _show_summary_dialog(root, successes, errors, _unregistered, config_warnings)



def _show_summary_dialog(
    root: tk.Tk,
    successes: list[str],
    errors: list[tuple[str, str]],
    unregistered: list[str],
    config_warnings: list[tuple[str, str, list[str]]],
) -> None:
    """PDF 일괄 생성 결과를 성공/경고/실패 섹션으로 구분하여 표시한다.

    Args:
        config_warnings: (경고 제목, 부연 설명, 항목 리스트) 튜플 목록 —
                         설정 vs CSV 불일치 등 비차단 경고를 ⚠️ 블록으로 렌더한다.
    """
    # 경고 '범주' 수 (미등록 팀 + 설정 경고 블록 수)
    n_warn = (1 if unregistered else 0) + len(config_warnings)

    win = tk.Toplevel(root)
    win.title("PDF 일괄 생성 결과")
    win.resizable(True, True)

    header = f"처리 완료: 성공 {len(successes)}건 / 경고 {n_warn}건 / 실패 {len(errors)}건"
    tk.Label(win, text=header, font=("맑은 고딕", 11, "bold"), pady=8).pack(fill="x", padx=10)

    txt = scrolledtext.ScrolledText(win, width=72, height=22, wrap="word", font=("Consolas", 9))
    txt.pack(padx=10, pady=(0, 6), fill="both", expand=True)

    txt.tag_configure("success",  foreground="#2d7a2d")
    txt.tag_configure("warning",  foreground="#c05621")
    txt.tag_configure("error",    foreground="#c0392b")
    txt.tag_configure("sec_bold", font=("맑은 고딕", 9, "bold"))

    def _write(text: str, *tags: str) -> None:
        txt.insert("end", text, tags)

    if successes:
        _write(f"✅ 성공 ({len(successes)}건)\n", "sec_bold", "success")
        for code in successes:
            _write(f"  • {code}\n", "success")
        _write("\n")

    if unregistered:
        _write(f"⚠️ 경고 PDF 목록 ({n_warn}건)\n", "sec_bold", "warning")
        _write(f"\n  [미등록 팀 — 집계 누락 가능성]\n", "warning")
        for team in unregistered:
            _write(f"  • {team}\n", "warning")
        _write("\n")

    for warn_title, warn_note, warn_items in config_warnings:
        _write(f"⚠️ {warn_title} ({len(warn_items)}건)\n", "sec_bold", "warning")
        _write(f"\n  [{warn_note}]\n", "warning")
        for item in warn_items:
            _write(f"  • {item}\n", "warning")
        _write("\n")

    if errors:
        _write(f"❌ 실패 ({len(errors)}건)\n", "sec_bold", "error")
        for code, reason in errors:
            _write(f"  • {code}: {reason}\n", "error")

    txt.configure(state="disabled")

    tk.Button(win, text="확인", width=10, command=win.destroy).pack(pady=(0, 10))

    win.wait_window()


if __name__ == "__main__":
    main()
