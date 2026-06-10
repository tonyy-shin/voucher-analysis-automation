"""
main.py
역할: 프로그램 진입점. tkinter 다이얼로그로 사용자 입력을 받고
      출력> 시트의 모든 출력전표 코드를 일괄 처리한다.

--noconsole 환경 (PyInstaller EXE): 터미널 창 없이 팝업 다이얼로그만 사용한다.
모든 오류는 tkinter.messagebox로 사용자에게 안내한 뒤 sys.exit(1)로 종료한다.
"""
from __future__ import annotations

import copy
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
    ColWarning, Mismatch, SHEET_LABELS,
    load_config, save_config, validate_against_excel, validate_optional_cols,
)

# ── 모듈 상수 ─────────────────────────────────────────────────────────────────
_PDF_FILENAME_TEMPLATE = "사업비_전표_분석_{}.pdf"


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
    from theme_manager import _derive_light  # 내부 헬퍼 재사용

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
    preview_frame.grid(row=0, column=0, columnspan=4, padx=12, pady=(12, 4), sticky="ew")

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
        row=2, column=0, columnspan=4, sticky="ew", padx=12, pady=(4, 0)
    )

    # ── 로고 슬롯 (3행) ───────────────────────────────────────────────────
    tk.Label(dlg, text="회사 로고 이미지 (최대 3개):").grid(
        row=3, column=0, columnspan=4, padx=12, pady=(8, 4), sticky="w"
    )

    slot_vars: list[tk.StringVar] = []

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
            return _pick

        def _make_clear(idx: int):
            def _clear():
                slots[idx] = None
                slot_vars[idx].set("미지정")
            return _clear

        tk.Button(dlg, text="변경", width=6, command=_make_pick(i)).grid(
            row=4 + i, column=2, padx=2, pady=3
        )
        tk.Button(dlg, text="삭제", width=6, command=_make_clear(i)).grid(
            row=4 + i, column=3, padx=(2, 12), pady=3
        )

    # ── 버튼 행 ──────────────────────────────────────────────────────────
    btn_frame = tk.Frame(dlg)
    btn_frame.grid(row=7, column=0, columnspan=4, pady=(8, 12), padx=12)

    def _save_and_close():
        valid_paths = [p for p in slots if p]
        new_theme = CompanyTheme(
            primary_hex=current_hex[0],
            primary_light_hex=_derive_light(current_hex[0]),
            logos=theme.logos,
            logo_paths=valid_paths,
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



def _show_validation_dialog(
    root: tk.Tk,
    mismatches: list,
    config: dict,
) -> dict | None:
    """
    Excel 파일과 설정 파일 간 불일치를 표시하고 사용자 수정을 요청하는 팝업.

    Returns:
        dict  — [이번만 적용] 또는 [설정 저장 후 계속] 선택 시 수정된 config
        None  — [취소] 또는 X 버튼 클릭 시
    """
    result: list[dict | None] = [None]

    dlg = tk.Toplevel(root)
    dlg.title("컬럼·시트명 검증")
    dlg.geometry("820x520")
    dlg.resizable(True, True)

    # ── 경고 헤더 ──────────────────────────────────────────────────────────
    hdr_frame = tk.Frame(dlg, bg="#FFF3CD", pady=8)
    hdr_frame.pack(fill="x")
    tk.Label(
        hdr_frame,
        text="⚠️  설정 파일과 Excel 파일의 이름이 일치하지 않습니다",
        bg="#FFF3CD", font=("맑은 고딕", 10, "bold"), fg="#856404",
    ).pack(side="left", padx=12)
    tk.Label(
        hdr_frame,
        text="  아래 목록에서 올바른 항목을 선택해 주세요.",
        bg="#FFF3CD", font=("맑은 고딕", 9), fg="#856404",
    ).pack(side="left")

    # ── 스크롤 영역 ────────────────────────────────────────────────────────
    scroll_frame = tk.Frame(dlg)
    scroll_frame.pack(fill="both", expand=True, padx=8, pady=4)

    canvas = tk.Canvas(scroll_frame, highlightthickness=0)
    vsb = tk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    inner = tk.Frame(canvas)
    inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_inner_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _on_canvas_configure(event):
        canvas.itemconfig(inner_id, width=event.width)

    inner.bind("<Configure>", _on_inner_configure)
    canvas.bind("<Configure>", _on_canvas_configure)
    def _on_canvas_mousewheel(e):
        if e.widget.winfo_class() in ('Listbox', 'TCombobox'):
            return  # 드롭다운 스크롤은 위젯이 직접 처리하도록 전파 차단
        canvas.yview_scroll(int(-1 * e.delta / 120), "units")
    canvas.bind_all("<MouseWheel>", _on_canvas_mousewheel)

    # ── 테이블 헤더 ────────────────────────────────────────────────────────
    hdr_style = {"bg": "#4A5568", "fg": "white", "font": ("맑은 고딕", 9, "bold"), "padx": 6, "pady": 4}
    for col_i, (txt, w) in enumerate([
        ("용도 / 논리명", 22), ("설정 파일 현재 값", 26), ("Excel에서 선택", 30), ("상태", 10)
    ]):
        tk.Label(inner, text=txt, width=w, anchor="w", **hdr_style).grid(
            row=0, column=col_i, sticky="ew", padx=1, pady=1
        )

    # ── 행 생성 ───────────────────────────────────────────────────────────
    combo_vars: list[tuple] = []  # (Mismatch, StringVar, Label)
    row_idx = 1

    sheet_mm = [m for m in mismatches if m.is_sheet_mismatch]
    col_mm   = [m for m in mismatches if not m.is_sheet_mismatch]

    def _section_header(text: str, r: int) -> int:
        tk.Label(inner, text=text, anchor="w", font=("맑은 고딕", 9, "bold"),
                 bg="#EDF2F7", fg="#2D3748", pady=3, padx=8
                 ).grid(row=r, column=0, columnspan=4, sticky="ew", pady=(6, 2))
        return r + 1

    def _add_row(m, r: int) -> int:
        tk.Label(inner, text=m.display_label, width=22, anchor="w",
                 font=("맑은 고딕", 9), padx=6, pady=3
                 ).grid(row=r, column=0, sticky="ew", padx=1)
        tk.Label(inner, text=m.expected, width=26, anchor="w",
                 font=("맑은 고딕", 9), fg="#C53030", padx=6
                 ).grid(row=r, column=1, sticky="ew", padx=1)

        var = tk.StringVar(value=m.fuzzy_guess or "")
        cb = ttk.Combobox(inner, textvariable=var, values=m.actual_cols, width=28, state="normal")
        cb.grid(row=r, column=2, sticky="ew", padx=4, pady=2)

        has_guess = bool(m.fuzzy_guess)
        status_lbl = tk.Label(
            inner,
            text="✅ 추천됨" if has_guess else "⚠️ 미선택",
            width=10, anchor="w", font=("맑은 고딕", 9),
            fg="#276749" if has_guess else "#C05621",
        )
        status_lbl.grid(row=r, column=3, sticky="ew", padx=4)
        combo_vars.append((m, var, status_lbl))
        return r + 1

    if sheet_mm:
        row_idx = _section_header("📋 시트명 불일치", row_idx)
        for m in sheet_mm:
            row_idx = _add_row(m, row_idx)

    if col_mm:
        row_idx = _section_header("📊 컬럼명 불일치", row_idx)
        current_key = None
        for m in col_mm:
            if m.sheet_key != current_key:
                current_key = m.sheet_key
                tk.Label(
                    inner,
                    text=f"  ▸ {SHEET_LABELS.get(m.sheet_key, m.sheet_key)} 시트",
                    anchor="w", font=("맑은 고딕", 9, "italic"),
                    bg="#F7FAFC", fg="#4A5568", pady=2, padx=12,
                ).grid(row=row_idx, column=0, columnspan=4, sticky="ew", pady=(4, 1))
                row_idx += 1
            row_idx = _add_row(m, row_idx)

    # ── 버튼 영역 ──────────────────────────────────────────────────────────
    btn_frame = tk.Frame(dlg)
    btn_frame.pack(fill="x", padx=8, pady=(4, 8))

    btn_auto   = tk.Button(btn_frame, text="자동 매칭 시도",   width=14)
    btn_apply  = tk.Button(btn_frame, text="이번만 적용",      width=14, state="disabled")
    btn_save   = tk.Button(btn_frame, text="설정 저장 후 계속", width=16, state="disabled")
    btn_cancel = tk.Button(btn_frame, text="취소",             width=8)

    btn_auto.pack(side="left", padx=4)
    btn_apply.pack(side="left", padx=4)
    btn_save.pack(side="left", padx=4)
    btn_cancel.pack(side="right", padx=4)

    def _check_buttons(*_):
        all_filled = all(v.get().strip() for _, v, _ in combo_vars)
        state = "normal" if all_filled else "disabled"
        btn_apply.config(state=state)
        btn_save.config(state=state)

    def _build_config() -> dict:
        updated = copy.deepcopy(config)
        for m, var, _ in combo_vars:
            chosen = var.get().strip()
            if not chosen:
                continue
            if m.is_sheet_mismatch:
                updated["sheets"][m.sheet_key] = chosen
            else:
                updated["columns"][m.logical_name] = chosen
        return updated

    def _auto_match():
        for m, var, lbl in combo_vars:
            if not var.get().strip() and m.fuzzy_guess:
                var.set(m.fuzzy_guess)
                lbl.config(text="✅ 자동선택", fg="#276749")
        _check_buttons()

    def _apply_only():
        result[0] = _build_config()
        dlg.destroy()

    def _save_and_continue():
        updated = _build_config()
        if not save_config(updated):
            messagebox.showwarning(
                "설정 저장 실패",
                "설정 파일 저장에 실패했습니다 (쓰기 권한 부족 또는 디스크 오류).\n\n"
                "이번 실행에는 수정된 설정이 적용되지만 다음 실행부터는 반영되지 않습니다.",
            )
        result[0] = updated
        dlg.destroy()

    def _cancel():
        result[0] = None
        dlg.destroy()

    # 콤보박스 변경 시 버튼 상태·레이블 갱신
    for m, var, lbl in combo_vars:
        def _make_trace(v, l):
            def _trace(*_):
                filled = bool(v.get().strip())
                l.config(text="✅ 선택됨" if filled else "⚠️ 미선택",
                         fg="#276749" if filled else "#C05621")
                _check_buttons()
            return _trace
        var.trace_add("write", _make_trace(var, lbl))

    btn_auto.config(command=_auto_match)
    btn_apply.config(command=_apply_only)
    btn_save.config(command=_save_and_continue)
    btn_cancel.config(command=_cancel)
    dlg.protocol("WM_DELETE_WINDOW", _cancel)

    _check_buttons()

    # PyInstaller 환경 포커스 보장
    dlg.update_idletasks()
    dlg.grab_set()
    dlg.lift()
    dlg.focus_force()

    dlg.wait_window()
    canvas.unbind_all("<MouseWheel>")
    return result[0]

def main() -> None:
    """
    프로그램 전체 흐름을 제어하는 메인 함수.

    실행 순서:
        1. tkinter root 초기화 (hide) — --noconsole 환경에서 대화상자만 표시
        2. 입력 데이터 파일 선택 (filedialog)
        3. 결과 저장 폴더 선택 (askdirectory)
        4. Phase 0 클렌징 포함 데이터 로드 (data_loader)
        5. 출력전표 코드 목록 추출 (출력> 시트)
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

    # ── Step 1: 입력 데이터 파일 선택 ─────────────────────────────────────
    input_path = filedialog.askopenfilename(
        title="입력 데이터 파일 선택",
        filetypes=[("모든 파일", "*.*"), ("Excel 파일", "*.xlsx *.xlsm *.xls")],
    )
    if not input_path:
        messagebox.showinfo("취소", "파일 선택이 취소되었습니다. 프로그램을 종료합니다.")
        sys.exit(0)

    # ── Step 1-B: 설정 사전 검증 ─────────────────────────────────────────
    mismatches = validate_against_excel(input_path, config)
    if mismatches:
        config = _show_validation_dialog(root, mismatches, config)
        if config is None:
            sys.exit(0)

    # ── Step 1-C: 선택적 컬럼 경고 (비차단) ──────────────────────────────
    col_warnings = validate_optional_cols(input_path, config)
    if col_warnings:
        detail_lines = []
        for w in col_warnings:
            missing_str = "\n    ".join(w.missing_cols)
            detail_lines.append(f"• {w.display_label}:\n    {missing_str}")
        detail = "\n\n".join(detail_lines)
        proceed = messagebox.askyesno(
            "선택적 컬럼 누락 경고",
            f"일부 컬럼을 찾을 수 없습니다:\n\n{detail}\n\n"
            "해당 컬럼이 완전히 누락된 경우 해당 출력전표 처리가 실패할 수 있습니다.\n"
            "계속 진행하시겠습니까?",
        )
        if not proceed:
            sys.exit(0)

    # -- Step 1-D: 미등록 부서 사전 검증 (차단 또는 무시) --------------------
    _unregistered = data_loader.check_unregistered_teams(input_path, config)
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
                "실제 발생 내역에 존재하지만 '출력>' 시트에 등록되지 않은 부서가 발견되었습니다.\n"
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
        sheets = data_loader.load_all_sheets(input_path, config=config)
    except Exception as exc:
        messagebox.showerror(
            "데이터 로드 오류",
            f"입력 파일을 읽는 중 오류가 발생했습니다.\n\n"
            f"사유: {exc}\n\n"
            f"시트명({config['sheets']['transaction']}·{config['sheets']['account']}·{config['sheets']['ccm']}·{config['sheets']['output']})과 "
            "파일 형식을 확인하세요.",
        )
        sys.exit(1)

    # ── Step 5: 출력전표 코드 목록 수집 (출력> 시트의 출력전표 컬럼에서) ──────────
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
            f"'{_SHEET_OUTPUT}' 시트의 '{col_output_code}' 컬럼에 처리할 출력전표 코드가 없습니다.",
        )
        sys.exit(1)

    # ── Step 6: 코드별 파이프라인 루프 ────────────────────────────────────
    successes: list[str] = []
    errors: list[tuple[str, str]] = []
    _pipeline_results: dict[str, dict] = {}
    _unclassified: dict[str, dict] = {}
    _dept_warnings: dict[str, dict] = {}

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

        _w = results.pop("_미분류경고", None)
        if _w:
            _unclassified[code_str] = {
                "warning": _w,
                "계정명": results["account_info"].get("계정명", ""),
            }

        d_주관 = results.get("dept_주관")
        d_사용 = results.get("dept_사용")
        _wflags = []
        if d_주관 is not None and d_주관.empty:
            _wflags.append("주관부서 귀속 없음")
        if d_사용 is not None and d_사용.empty:
            _wflags.append("사용부서 귀속 없음")
        if _wflags:
            _dept_warnings[code_str] = {
                "계정명": results["account_info"].get("계정명", ""),
                "flags": _wflags,
            }

        _pipeline_results[code_str] = results

    # ── Stage 2: 미분류건 사전 경고 (있을 때만) ───────────────────────────
    if _unclassified:
        _proceed = False

        def _build_unclassified_dialog(parent: tk.Tk) -> None:
            nonlocal _proceed
            dlg = tk.Toplevel(parent)
            dlg.title("미분류건 감지")
            dlg.resizable(False, False)
            def _refocus(event=None):
                dlg.focus_force()
            dlg.bind("<FocusOut>", _refocus)
            dlg.focus_force()

            _msg = (
                "직접비·공통비로 분류되지 않은 행이 발견되었습니다.\n"
                "이대로 진행하면 해당 금액이 직간접 분류 합계에서 누락됩니다.\n"
                "PDF 생성 전에 원본 데이터의 직간접 구분 값을 수정해 주세요."
            )
            tk.Label(
                dlg,
                text=_msg,
                wraplength=460,
                justify="left",
                fg="#c0392b",
                font=("Malgun Gothic", 10, "bold"),
                padx=16, pady=14,
            ).pack(fill="x")

            frame_list = tk.Frame(dlg, padx=16, pady=4)
            frame_list.pack(fill="both", expand=True)

            scrollbar = tk.Scrollbar(frame_list, orient="vertical")
            txt = tk.Text(
                frame_list,
                height=12,
                width=60,
                yscrollcommand=scrollbar.set,
                font=("Malgun Gothic", 10),
                relief="sunken",
                bd=1,
            )
            scrollbar.config(command=txt.yview)
            scrollbar.pack(side="right", fill="y")
            txt.pack(side="left", fill="both", expand=True)

            lines = []
            for _code, info in _unclassified.items():
                w = info["warning"]
                lines.append(
                    f"• [{_code}] {info['계정명']}   미분류 금액: {w['미분류금액']:,.0f}원"
                )
                for row in w.get("미분류행", []):
                    분류값 = row.get("분류값", "")
                    분류값_str = str(분류값).strip() if 분류값 else "(빈칸)"
                    lines.append(
                        f"    - 원가요소: {row.get('원가요소', '')} / "
                        f"분류값: \"{분류값_str}\" / "
                        f"금액: {row.get('금액', 0):,.0f}원"
                    )
                lines.append("")
            txt.insert("1.0", "\n".join(lines))
            txt.config(state="disabled")

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
            _w_px = dlg.winfo_reqwidth()
            _h_px = dlg.winfo_reqheight()
            _sw = dlg.winfo_screenwidth()
            _sh = dlg.winfo_screenheight()
            dlg.geometry(f"{_w_px}x{_h_px}+{(_sw - _w_px) // 2}+{(_sh - _h_px) // 2}")

            parent.wait_window(dlg)

        _build_unclassified_dialog(root)
        if not _proceed:
            sys.exit(0)

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
        ok = pdf_exporter.export(results, pdf_path, code_str, theme=theme)
        if not ok:
            errors.append((code_str, "PDF 생성 실패"))
            continue

        successes.append(code_str)

    # ── Step 7: 완료 요약 ──────────────────────────────────────────────────
    _show_summary_dialog(root, successes, errors, _unclassified, _dept_warnings)


def _show_scrollable_error(root: tk.Tk, title: str, message: str) -> None:
    """
    긴 메시지를 스크롤 가능한 창으로 표시한다.

    Args:
        root:    tkinter 루트 창 (부모)
        title:   창 제목
        message: 표시할 전체 메시지
    """
    win = tk.Toplevel(root)
    win.title(title)
    win.resizable(True, True)

    lbl = tk.Label(win, text=title, font=("", 11, "bold"), fg="red", pady=6)
    lbl.pack(fill="x", padx=10)

    txt = scrolledtext.ScrolledText(win, width=70, height=20, wrap="word", font=("Consolas", 9))
    txt.pack(padx=10, pady=(0, 6), fill="both", expand=True)
    txt.insert("1.0", message)
    txt.configure(state="disabled")

    btn = tk.Button(win, text="확인", width=10, command=win.destroy)
    btn.pack(pady=(0, 10))

    win.grab_set()
    win.wait_window()


def _show_summary_dialog(
    root: tk.Tk,
    successes: list[str],
    errors: list[tuple[str, str]],
    unclassified: dict[str, dict],
    dept_warnings: dict[str, dict],
) -> None:
    """PDF 일괄 생성 결과를 성공/경고/실패 섹션으로 구분하여 표시한다."""
    warn_codes = set(unclassified) | set(dept_warnings)
    n_warn = len(warn_codes)

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

    if warn_codes:
        _write(f"⚠️ 경고 PDF 목록 ({n_warn}건)\n", "sec_bold", "warning")
        for code in sorted(warn_codes):
            계정명 = (
                unclassified.get(code, {}).get("계정명")
                or dept_warnings.get(code, {}).get("계정명")
                or ""
            )
            parts: list[str] = []
            if code in unclassified:
                w = unclassified[code]["warning"]
                parts.append(f"직간접 미분류 (미분류금액: {w['미분류금액']:,.0f}원)")
            if code in dept_warnings:
                parts.extend(dept_warnings[code]["flags"])
            _write(f"  • [{code}] {계정명} — {' / '.join(parts)}\n", "warning")
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
