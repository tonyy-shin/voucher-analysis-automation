"""
main.py
역할: 프로그램 진입점. tkinter 다이얼로그로 사용자 입력을 받고
      출력> 시트의 모든 출력전표 코드를 일괄 처리한다.

--noconsole 환경 (PyInstaller EXE): 터미널 창 없이 팝업 다이얼로그만 사용한다.
모든 오류는 tkinter.messagebox로 사용자에게 안내한 뒤 sys.exit(1)로 종료한다.
"""
from __future__ import annotations

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
    _label_groups: list[tuple[str, list[tuple[tuple, str]]]] = [
        ("상단", [
            (("header",), "제목 배너"),
        ]),
        ("계정 테이블", [
            (("account_table", "계정"), "계정"),
            (("account_table", "계정코드"), "계정코드"),
            (("account_table", "계정명"), "계정명"),
            (("account_table", "사업비코드"), "사업비 코드"),
            (("account_table", "채널구분"), "채널구분"),
        ]),
        ("1. 대상정의", [
            (("section1", "제목"), "섹션 제목"),
            (("section1", "대상정의"), "대상정의"),
            (("section1", "범위"), "범위"),
            (("section1", "지급대상"), "지급대상"),
            (("section1", "산출기준"), "산출기준"),
        ]),
        ("2. 부점귀속 현황", [
            (("section2", "제목"), "섹션 제목"),
            (("section2", "귀속현황"), "귀속현황 (섹션2)"),
            (("section2", "주관부서귀속"), "주관부서 귀속"),
            (("section2", "실제사용부서귀속"), "실제 사용부서 귀속"),
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
            (("section3_1", "직접비"), "직접비"),
            (("section3_1", "공통비"), "공통비"),
        ]),
        # 성격 컬럼(nature_cols)은 개수가 가변이므로 아래 정적 목록이 아닌
        # 동적 편집 블록(_render_nature_rows)으로 렌더링한다.
        ("3-2. 성격별 분류 결과", [
            (("section3_2", "제목"), "섹션 제목"),
            (("section3_2", "귀속현황"), "귀속현황 (섹션3-2)"),
            (("section3_2", "합계"), "합계 헤더"),
            (("section3_2", "설명"), "설명 문구"),
        ]),
        ("4. 분류 근거", [
            (("section4", "제목"), "섹션 제목"),
            (("section4", "직공통비"), "좌측 라벨 1 (직 • 공통비)"),
            (("section4", "성격별분류"), "좌측 라벨 2 (성격별 분류)"),
            (("section4", "참조"), "우측 참조 문구"),
        ]),
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
            tk.Button(
                nature_frame, text="삭제", width=5,
                command=(lambda i=idx: _delete_nature(i)),
            ).grid(row=idx, column=2, padx=(0, 8), pady=2)

    def _delete_nature(idx: int):
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

    label_vars: dict[tuple, tk.StringVar] = {}
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

        # 3-2 그룹 직후에 성격 컬럼 동적 편집 블록을 삽입한다.
        if group_title == "3-2. 성격별 분류 결과":
            tk.Label(
                _lbl_inner, text="성격 컬럼 (CSV 컬럼명 = 표시명, 3-2 표)",
                font=("맑은 고딕", 9, "bold"), anchor="w", fg="#2E2E38",
            ).grid(row=_r, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 0))
            _r += 1
            tk.Label(
                _lbl_inner,
                text="※ CSV 컬럼명과 표시명에 동일하게 적용됩니다. "
                     "사업비정보.csv에 없는 이름은 빈 칸으로 출력됩니다.",
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

    # 설정된 성격 컬럼 중 사업비정보.csv에 실제로 없는 컬럼 (전 코드 공통 — 1회 계산)
    unknown_nature = [
        c for c in data_loader.NATURE_COLS
        if c not in sheets[data_loader._SHEET_TRANSACTION].columns
    ]

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
            results = processor.run_pipeline(sheets, code_str)
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
    _show_summary_dialog(root, successes, errors, _unregistered, unknown_nature)



def _show_summary_dialog(
    root: tk.Tk,
    successes: list[str],
    errors: list[tuple[str, str]],
    unregistered: list[str],
    unknown_nature: list[str],
) -> None:
    """PDF 일괄 생성 결과를 성공/경고/실패 섹션으로 구분하여 표시한다."""
    # 경고 '범주' 수 (미등록 팀 / 미등록 성격 컬럼)
    n_warn = (1 if unregistered else 0) + (1 if unknown_nature else 0)

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

    if unknown_nature:
        _write(f"⚠️ 미등록 성격 컬럼 ({len(unknown_nature)}건)\n", "sec_bold", "warning")
        _write("\n  [CSV에 없는 컬럼 — 해당 컬럼은 빈 칸으로 출력됩니다]\n", "warning")
        for col in unknown_nature:
            _write(f"  • {col}\n", "warning")
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
