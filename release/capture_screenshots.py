"""
capture_screenshots.py
역할: 사용설명서(generate_manual.py)에 삽입할 앱 화면 스크린샷 6장을 자동 캡처한다.

main.py의 실제 다이얼로그 함수를 그대로 호출하여 화면을 띄우고, 표시되는 동안
PIL.ImageGrab으로 창 영역만 잘라 PNG로 저장한다. 모든 다이얼로그는 wait_window()로
블로킹되므로, 캡처는 root.after()로 다이얼로그 이벤트 루프 안에서 실행한 뒤 destroy()로
루프를 해제한다.

전제: 실제 Windows 데스크톱 세션(화면 잠금 해제, 창이 가려지지 않음). pyautogui 등 추가
      의존성 없이 이미 설치된 Pillow의 ImageGrab만 사용한다.

실행:  python release/capture_screenshots.py
출력:  release/screenshots/01_brand_tab.png ... 06_summary.png
"""
from __future__ import annotations

import ctypes
import json
import sys
import time
from pathlib import Path

import tkinter as tk
from tkinter import ttk

from PIL import ImageGrab

# ── 경로 ──────────────────────────────────────────────────────────────────────
_HERE = Path(__file__).resolve().parent                       # release/
_PROJ = _HERE.parent                                          # python_data_cleaning/
_CODE = _PROJ / "report_generator" / "src" / "code"
_OUT_DIR = _HERE / "screenshots"
_TEMPLATES = _HERE / "templates" / "input_data"

# main.py와 형제 모듈(data_loader 등)을 import 하기 위해 코드 디렉터리를 경로에 추가
sys.path.insert(0, str(_CODE))

import main as app                                            # noqa: E402
from theme_manager import load_theme                          # noqa: E402
from config_manager import load_config                        # noqa: E402

_OPEN_DELAY_MS = 800   # 다이얼로그가 그려질 시간을 준 뒤 캡처 드라이버 실행

# 파일 선택 다이얼로그 스크린샷(04)에 실제 사용자명이 노출되지 않도록 사용하는 익명 더미 경로
_DUMMY_PATHS = {
    "transaction": "C:/data/사업비정보.csv",
    "account":     "C:/data/계정정보.csv",
    "ccm":         "C:/data/부서정보.csv",
    "output":      "C:/data/출력.csv",
}


# ── 공통 유틸 ─────────────────────────────────────────────────────────────────

def _set_dpi_aware() -> None:
    """High-DPI 환경에서 winfo_* 논리좌표와 ImageGrab 물리픽셀의 불일치를 방지한다."""
    for fn in (
        lambda: ctypes.windll.shcore.SetProcessDpiAwareness(1),
        lambda: ctypes.windll.user32.SetProcessDPIAware(),
    ):
        try:
            fn()
            return
        except Exception:
            continue


def _latest_toplevel(root: tk.Tk) -> tk.Toplevel | None:
    tops = [w for w in root.winfo_children() if isinstance(w, tk.Toplevel)]
    return tops[-1] if tops else None


def _find_descendant(widget, cls):
    for child in widget.winfo_children():
        if isinstance(child, cls):
            return child
        found = _find_descendant(child, cls)
        if found is not None:
            return found
    return None


def _find_all(widget, cls) -> list:
    """widget 하위 트리에서 cls 인스턴스를 깊이우선 순서로 모두 수집한다."""
    out = []
    for child in widget.winfo_children():
        if isinstance(child, cls):
            out.append(child)
        out.extend(_find_all(child, cls))
    return out


def _find_label_startswith(widget, prefix: str):
    """text가 prefix로 시작하는 첫 tk.Label을 재귀 탐색하여 반환한다(없으면 None)."""
    for child in widget.winfo_children():
        if isinstance(child, tk.Label):
            try:
                if str(child.cget("text")).startswith(prefix):
                    return child
            except tk.TclError:
                pass
        found = _find_label_startswith(child, prefix)
        if found is not None:
            return found
    return None


def _scroll_to_widget(canvas: tk.Canvas, inner: tk.Frame, target) -> None:
    """스크롤 캔버스에서 target 위젯이 보이도록 yview를 이동한다(target을 상단 근처에)."""
    canvas.update_idletasks()
    total = inner.winfo_height()
    if total <= 0:
        return
    rel_y = target.winfo_rooty() - inner.winfo_rooty()
    frac = max(0.0, min(1.0, rel_y / total))
    canvas.yview_moveto(frac)


def _settle(win, ms: int = 250) -> None:
    """이벤트 루프를 잠깐 더 돌려 위젯 렌더링을 안정화한다(ImageGrab은 실제 화면을 캡처)."""
    win.update_idletasks()
    win.update()
    end = time.time() + ms / 1000
    while time.time() < end:
        win.update()
        time.sleep(0.01)


def _grab(win, out_path: Path) -> None:
    win.attributes("-topmost", True)
    win.lift()
    win.focus_force()
    _settle(win)
    x, y = win.winfo_rootx(), win.winfo_rooty()
    w, h = win.winfo_width(), win.winfo_height()
    ImageGrab.grab(bbox=(x, y, x + w, y + h)).save(out_path)
    print(f"[shot] {out_path.name}  ({w}x{h} @ {x},{y})")


# ── last_paths.json 임시 시드 (파일 선택 다이얼로그에 경로가 보이도록) ──────────

def _seed_last_paths() -> tuple[Path, str | None]:
    path = app._LAST_PATHS_FILE
    backup = path.read_text(encoding="utf-8") if path.exists() else None
    seed = {
        "transaction": str(_TEMPLATES / "사업비정보.csv"),
        "account":     str(_TEMPLATES / "계정정보.csv"),
        "ccm":         str(_TEMPLATES / "부서정보.csv"),
        "output":      str(_TEMPLATES / "출력.csv"),
    }
    try:
        path.write_text(json.dumps(seed, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass
    return path, backup


def _restore_last_paths(state: tuple[Path, str | None]) -> None:
    path, content = state
    try:
        if content is None:
            if path.exists():
                path.unlink()
        else:
            path.write_text(content, encoding="utf-8")
    except OSError:
        pass


# ── 캡처 시나리오 ─────────────────────────────────────────────────────────────

def _capture_theme(root: tk.Tk, theme, config) -> None:
    """포인트 색상·로고·문구 설정 다이얼로그 → 브랜드/문구/성격컬럼 3장."""
    def driver() -> None:
        dlg = _latest_toplevel(root)
        if dlg is None:
            print("[warn] theme dialog not found"); return
        nb = _find_descendant(dlg, ttk.Notebook)
        tabs = nb.tabs()

        # 1) 브랜드 설정 탭
        nb.select(tabs[0]); _settle(dlg)
        _grab(dlg, _OUT_DIR / "01_brand_tab.png")

        # 2) 문구 설정 탭 (상단)
        nb.select(tabs[1]); _settle(dlg)
        canvas = _find_descendant(nb.nametowidget(tabs[1]), tk.Canvas)
        inner = _find_descendant(canvas, tk.Frame) if canvas is not None else None
        if canvas is not None:
            canvas.update_idletasks()
            canvas.yview_moveto(0.0)
        _settle(dlg)
        _grab(dlg, _OUT_DIR / "02_labels_tab.png")

        # 3) 성격 컬럼 추가/삭제 영역 (해당 블록으로 스크롤)
        if canvas is not None:
            lbl = _find_label_startswith(inner, "성격 컬럼") if inner is not None else None
            if inner is not None and lbl is not None:
                _scroll_to_widget(canvas, inner, lbl)
            else:
                canvas.yview_moveto(1.0)
        _settle(dlg)
        _grab(dlg, _OUT_DIR / "03_nature_cols.png")

        # 4) 분류 근거(섹션 4) 동적 행 편집기 (탭 최하단으로 스크롤)
        if canvas is not None:
            canvas.yview_moveto(1.0)
        _settle(dlg)
        _grab(dlg, _OUT_DIR / "07_basis_rows.png")

        dlg.destroy()   # wait_window 해제 (저장 안 함 = 부수효과 없음)

    root.after(_OPEN_DELAY_MS, driver)
    app._show_theme_dialog(root, theme, config)


def _capture_file_select(root: tk.Tk) -> None:
    # 익명화: 스크린샷에 실제 사용자명이 담긴 경로가 보이지 않도록 더미 경로를 표시한다.
    #   1) last_paths.json 을 더미 경로로 임시 덮어쓴 뒤(요청사항), finally에서 원복한다.
    #   2) 다만 다이얼로그는 존재하는 파일 경로만 표시하므로(main._show_file_select_dialog의
    #      Path(saved).is_file() 가드), 캡처 직전 읽기전용 Entry 4칸 텍스트를 더미 경로로 직접
    #      세팅하여 안정적으로 노출시킨다.
    path = app._LAST_PATHS_FILE
    state = (path, path.read_text(encoding="utf-8") if path.exists() else None)

    def driver() -> None:
        dlg = _latest_toplevel(root)
        if dlg is None:
            print("[warn] file-select dialog not found"); return
        entries = _find_all(dlg, tk.Entry)
        keys = [k for k, _ in app._FILE_SELECT_ROWS]
        for entry, key in zip(entries, keys):
            try:
                entry.config(state="normal")
                entry.delete(0, "end")
                entry.insert(0, _DUMMY_PATHS[key])
                entry.config(state="readonly")
            except tk.TclError:
                pass
        _grab(dlg, _OUT_DIR / "04_file_select.png")
        dlg.destroy()

    try:
        try:
            path.write_text(
                json.dumps(_DUMMY_PATHS, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except OSError:
            pass
        root.after(_OPEN_DELAY_MS, driver)
        app._show_file_select_dialog(root)
    finally:
        _restore_last_paths(state)


def _capture_unregistered(root: tk.Tk) -> None:
    teams = ["디지털영업부", "리테일심사팀", "기업금융2부", "투자전략실"]

    def driver() -> None:
        dlg = _latest_toplevel(root)
        if dlg is None:
            print("[warn] unregistered dialog not found"); return
        _grab(dlg, _OUT_DIR / "05_unregistered.png")
        dlg.destroy()

    root.after(_OPEN_DELAY_MS, driver)
    app.show_unregistered_dialog(root, teams)


def _capture_summary(root: tk.Tk) -> None:
    successes = ["5001", "5002", "5003", "6010"]
    errors = [("7002", "저장 폴더 쓰기 권한 없음")]
    unregistered = ["디지털영업부", "리테일심사팀"]
    unknown_nature = ["신규성격비"]   # 사업비정보.csv에 없는 성격 컬럼 → 경고 노출용

    def driver() -> None:
        win = _latest_toplevel(root)
        if win is None:
            print("[warn] summary dialog not found"); return
        _grab(win, _OUT_DIR / "06_summary.png")
        win.destroy()

    root.after(_OPEN_DELAY_MS, driver)
    app._show_summary_dialog(root, successes, errors, unregistered, unknown_nature)


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main() -> None:
    _set_dpi_aware()
    _OUT_DIR.mkdir(parents=True, exist_ok=True)

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    theme = load_theme()
    config = load_config()

    state = _seed_last_paths()
    try:
        _capture_theme(root, theme, config)
        _capture_file_select(root)
        _capture_unregistered(root)
        _capture_summary(root)
    finally:
        _restore_last_paths(state)
        try:
            root.destroy()
        except tk.TclError:
            pass

    print(f"[done] 스크린샷 캡처 완료 → {_OUT_DIR}")


if __name__ == "__main__":
    main()
