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
    load_config, validate_csv_columns,
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


def _show_theme_dialog(root: tk.Tk, theme: CompanyTheme) -> CompanyTheme:
    """포인트 색상·로고 설정 팝업 (로고 최대 3개 슬롯).

    Returns:
        [저장 후 계속] 클릭: 저장된 새 CompanyTheme
        [취소] 또는 X 버튼: 원본 theme 반환 (저장 안 함)
        [기본값 복원] 클릭: GRAY_DEFAULT 저장 후 반환
    """
    result = [theme]
    current_hex = [theme.primary_hex]

    # 3슬롯 초기화: theme.logo_paths 에서 최대 3개, 부족분은 None 으로 패딩
    slots: list[str | None] = (list(theme.logo_paths[:3]) + [None, None, None])[:3]

    dlg = tk.Toplevel(root)
    dlg.title("포인트 색상·로고 설정")
    dlg.resizable(False, False)
    dlg.lift()

    # ── 색상 미리보기 레이블 ──────────────────────────────────────────────
    preview_frame = tk.Frame(dlg, bd=1, relief="sunken")
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

    tk.Label(dlg, text="섹션 헤더·상하단 바 색상:").grid(row=1, column=0, padx=12, pady=4, sticky="w")
    tk.Button(dlg, text="색상 선택", command=_pick_color).grid(row=1, column=1, padx=4, pady=4)

    # ── 구분선 ────────────────────────────────────────────────────────────
    tk.Frame(dlg, height=1, bg="#cccccc").grid(
        row=2, column=0, columnspan=7, sticky="ew", padx=12, pady=(4, 0)
    )

    # ── 로고 슬롯 (3행) ───────────────────────────────────────────────────
    tk.Label(dlg, text="회사 로고 이미지 (최대 3개):").grid(
        row=3, column=0, columnspan=7, padx=12, pady=(8, 4), sticky="w"
    )

    slot_vars: list[tk.StringVar] = []
    scale_vars: list[tk.DoubleVar] = []

    def update_preview(*_args):  # no-op; PIL 블록에서 실제 함수로 교체됨
        pass

    for i in range(3):
        tk.Label(dlg, text=f"로고 {i + 1}:", width=7, anchor="e").grid(
            row=4 + i, column=0, padx=(12, 4), pady=3, sticky="e"
        )

        var = tk.StringVar(value=_fmt_slot_label(slots[i]))
        slot_vars.append(var)
        tk.Label(dlg, textvariable=var, font=("Consolas", 8), fg="#444444",
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

        tk.Button(dlg, text="변경", width=6, command=_make_pick(i)).grid(
            row=4 + i, column=2, padx=2, pady=3
        )
        tk.Button(dlg, text="삭제", width=6, command=_make_clear(i)).grid(
            row=4 + i, column=3, padx=(2, 8), pady=3
        )

        # ── 슬롯별 높이 슬라이더 (column 4~6) ───────────────────────────
        tk.Label(dlg, text="높이:", anchor="e").grid(
            row=4 + i, column=4, padx=(8, 2), pady=3, sticky="e"
        )
        sv = tk.DoubleVar(value=theme.logo_heights[i])
        scale_vars.append(sv)
        val_lbl = tk.Label(dlg, text=str(int(theme.logo_heights[i])), width=4, anchor="e")
        val_lbl.grid(row=4 + i, column=6, padx=(2, 12), pady=3, sticky="e")

        def _make_scale_cb(lbl):
            def _cb(val):
                lbl.config(text=str(int(float(val))))
            return _cb

        ttk.Scale(
            dlg, from_=16, to=96, orient="horizontal",
            variable=sv, command=_make_scale_cb(val_lbl),
        ).grid(row=4 + i, column=5, padx=2, pady=3, sticky="ew")

    # ── 구분선 ────────────────────────────────────────────────────────────
    tk.Frame(dlg, height=1, bg="#cccccc").grid(
        row=7, column=0, columnspan=7, sticky="ew", padx=12, pady=(8, 0)
    )

    # ── 로고 미리보기 캔버스 ─────────────────────────────────────────────
    _PREV_W, _PREV_H = 560, 100
    try:
        from PIL import Image, ImageTk  # optional; silently skipped if absent

        preview_canvas = tk.Canvas(
            dlg, width=_PREV_W, height=_PREV_H,
            bg="#f5f5f5", highlightthickness=1,
            highlightbackground="#cccccc",
        )
        preview_canvas.grid(row=8, column=0, columnspan=7, padx=12, pady=(4, 0))

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

    # ── 버튼 행 ──────────────────────────────────────────────────────────
    btn_frame = tk.Frame(dlg)
    btn_frame.grid(row=9, column=0, columnspan=7, pady=(8, 12), padx=12)

    def _save_and_close():
        valid_paths = [p for p in slots if p]
        new_theme = CompanyTheme(
            primary_hex=current_hex[0],
            logos=theme.logos,
            logo_paths=valid_paths,
            logo_max_height=max(int(v.get()) for v in scale_vars),
            logo_heights=[int(v.get()) for v in scale_vars],
        )
        save_theme(new_theme)
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
        font=("맑은 고딕", 10, "bold"), padx=12, pady=(12, 6),
    ).grid(row=0, column=0, columnspan=3, sticky="w")

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
        "PDF 리포트의 포인트 색상을 변경하시겠습니까?\n\n"
        f"현재 색상: {theme.primary_hex}\n"
        "('아니오' 선택 시 현재 색상으로 바로 진행합니다.)",
    ):
        theme = _show_theme_dialog(root, theme)

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
        _dept_lines = "\n".join(f"  {chr(8226)} {t}" for t in _unregistered)

        _proceed = False

        def _build_unregistered_dialog(parent: tk.Tk) -> None:
            nonlocal _proceed
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
                root.destroy()
                sys.exit(0)

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

        _build_unregistered_dialog(root)
        if not _proceed:
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
        ok = pdf_exporter.export(results, pdf_path, code_str, theme=theme, logo_heights=theme.logo_heights)
        if not ok:
            errors.append((code_str, "PDF 생성 실패"))
            continue

        successes.append(code_str)

    # ── Step 7: 완료 요약 ──────────────────────────────────────────────────
    _show_summary_dialog(root, successes, errors, _unregistered)



def _show_summary_dialog(
    root: tk.Tk,
    successes: list[str],
    errors: list[tuple[str, str]],
    unregistered: list[str],
) -> None:
    """PDF 일괄 생성 결과를 성공/경고/실패 섹션으로 구분하여 표시한다."""
    n_warn = 1 if unregistered else 0

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

    if errors:
        _write(f"❌ 실패 ({len(errors)}건)\n", "sec_bold", "error")
        for code, reason in errors:
            _write(f"  • {code}: {reason}\n", "error")

    txt.configure(state="disabled")

    tk.Button(win, text="확인", width=10, command=win.destroy).pack(pady=(0, 10))

    win.wait_window()


if __name__ == "__main__":
    main()
