# Expense Statement Analysis Automation

A Windows tool that auto-generates **Business Expense Statement Analysis PDFs** in bulk from Excel source data (expense pivot tables), one PDF per output code.

**Version:** v2.2.0 (2026/06)

---

## Key Features

- **Batch processing by output code** — Iterates through codes registered in the `출력>` sheet and generates a separate PDF for each
- **Triple LEFT JOIN** — Actual Expense ↔ Account Info ↔ Cost Center Master ↔ Output sheet
- **Flexible column & sheet mapping** — Externalized via `column_config.yaml`; GUI correction dialog opens on any mismatch, with fuzzy-match auto-suggestions
- **PDF theme settings** — Point color and logos (up to 3 slots) selectable at runtime with a **live preview canvas**
- **Per-slot logo heights** — Each of the 3 logo slots has an independent height slider (16–96 pt range)
- **Unregistered team detection** — Pre-flight check warns if transaction-sheet teams are absent from the output sheet
- **Unclassified item warnings** — Detects and reports individual rows where amounts could not be classified, with cost center, classification value, and amount details
- **Dynamic header row detection** — Scans the first 20 rows to locate actual column headers, tolerating variable file layouts
- **Case-insensitive amount column mapping** — e.g., `"sum of da_p"` is automatically resolved to `"Sum of DA_P"`
- **Fuzzy matching** — Auto-suggests correct column names when version strings change (e.g., `대상정의 v3.0_0415`)
- **문구 설정 탭 (Display-label editing tab)** — A second tab in the theme dialog where every PDF display label (titles, headers, section labels) can be edited from a scrollable form
- **성격 컬럼 동적 추가/삭제 (Dynamic nature-column add/delete)** — Add or remove nature classification columns from the 문구 설정 탭; each entry sets both the CSV column name and the PDF display name at once
- **미등록 성격 컬럼 경고 (Unregistered nature-column warning)** — If a configured nature column is missing from `사업비정보.csv`, it is listed as a warning in the final summary dialog
- **Section 4 분류 근거 동적 행 (Dynamic classification-basis rows)** — Section 4 ("분류 근거") is a typed, dynamic row system editable in the 문구 설정 탭. Three row types — **직공통비**, **성격별분류**, and **custom** — can be freely added and deleted. For custom rows the **근거 텍스트 컬럼** field names a `출력.csv` column whose first value becomes the cell content, and for every row type the **참조문서 컬럼** field names a `출력.csv` column whose first value fills the reference cell (left blank if the column is missing or empty)

---

## Cautions / 주의사항

- **CSV 파일 열린 채로 실행 금지** — If any of the 4 input CSV files are open in Excel or another program when the app runs, a file-access error occurs. Close all CSV files before launching.
- **문구 설정 컬럼명 = CSV 컬럼명 필수** — Nature column names entered in the 문구 설정 탭 must exactly match the column headers in `사업비정보.csv` (characters and spacing). A mismatch produces a **blank column in the PDF with no error** — only a warning in the summary dialog.
- **첫 실행 시 몇 초 대기 권장** — On Windows, running the EXE immediately after download may fail due to antivirus/OS scanning. Wait a few seconds after download before running.

---

## Requirements

| Item | Minimum Version |
|------|----------------|
| Python | 3.10+ (3.11+ recommended) |
| OS | Windows (Malgun Gothic font required) |
| Font | `C:\Windows\Fonts\malgun.ttf` |
| tkinter | Bundled with Python (no separate install needed) |

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Running (Development)

```bash
cd report_generator/src/code
python main.py
```

Execution flow:
1. Select point color and configure logos (up to 3 slots) with live preview
2. Select the input Excel file
3. Pre-flight validation — unregistered team check, schema mismatch detection
4. If mismatches exist: GUI correction dialog with fuzzy-match suggestions and three save options — **Apply once**, **Save & continue**, or **Cancel**
5. Select the output folder
6. Auto-generate PDFs per code → completion summary with warnings

---

## EXE Build

```bat
build.bat
```

After the build completes:
- `dist/전표분석서_자동화.exe` — Standalone executable
- `전표분석서_자동화_배포용.zip` — Distribution package (EXE + 2 config files)

---

## File Structure

```
python_data_cleaning/
├── report_generator/src/code/
│   ├── main.py              # Entry point — tkinter UI, full execution flow
│   ├── config_manager.py    # column_config.yaml load/save/validation + fuzzy matching
│   ├── data_loader.py       # 4-sheet load + Phase 0 cleansing + unregistered team check
│   ├── processor.py         # Pipeline — filtering / triple JOIN / aggregation / warnings
│   ├── pdf_exporter.py      # ReportLab A4 landscape PDF with dynamic logo layouts
│   ├── company_theme.py     # Theme dataclass (color + 3-slot logo heights)
│   ├── theme_manager.py     # theme_config.json save/load with migration support
│   └── templates/
│       └── report_template.html  # Layout reference HTML
├── images/                  # Default logo images (ABL.png, 우리금융그룹.png)
├── manual/                  # User manual
├── build.bat                # Build launcher (double-click)
├── build.ps1                # Build automation script (PyInstaller)
├── column_config.yaml       # Column & sheet name config (edit when Excel structure changes)
├── theme_config.json        # Theme settings (auto-saved by the program)
└── requirements.txt         # Python dependencies
```

---

## Configuration Files

### `column_config.yaml`

Edit this file when sheet names or column names in the Excel source change.

```yaml
sheets:
  transaction: "실제발생사업비"   # Must match the sheet tab name exactly
  account: "계정정보"
  ccm: "Cost Center Master"
  output: "출력>"

columns:
  대상정의: "대상정의 v3.0_0415"  # Update when the version string changes
  cc_name: "Cost Center name"     # Case-sensitive
  # ... (remaining columns)

amount_cols:                       # 10 amount columns (case-insensitive matching)
  - "Sum of DA_P"
  - "Sum of DA_N"
  # ...

nature_cols:                       # 5 nature classification columns
  - "계약체결비"
  - "계약유지비"
  # ...
```

- If a mismatch is detected at runtime, the GUI correction dialog allows you to fix and save immediately.
- If the file is missing, it is auto-generated with annotated defaults on first run.

### `theme_config.json`

Stores the PDF point color, per-slot logo heights, and logo file paths. Auto-saved by the program — no manual editing required.

```json
{
  "primary_hex": "#2C5F8A",
  "logo_heights": [40, 36, 32],
  "logo_paths": ["images/ABL.png", "images/우리금융그룹.png", null]
}
```

---

## Data Pipeline

```
Excel Input
    │
    ▼
[data_loader]  4-sheet load + Phase 0 cleansing
    │  - Dynamic header row detection (scans first 20 rows)
    │  - Type coercion (JOIN keys unified as str)
    │  - Whitespace stripping on all string columns
    │  - Numeric NaN → 0
    │  - Case-insensitive amount column name normalization
    │  - Unregistered team pre-flight check
    │
    ▼
[processor]  Per-code pipeline
    │  Step 1. Filter by cost element keyword
    │  Step 2. Triple LEFT JOIN + target amount aggregation
    │  Step 3. Principal / user department attribution
    │  Step 4. Direct / indirect cost classification
    │  Step 5. Nature × department cross-tabulation
    │  Step 6. Account header metadata extraction
    │  Step 7. Unclassified item detection & warning generation
    │
    ▼
[pdf_exporter]  A4 landscape PDF (ReportLab)
    │  - Dynamic logo layout (1 / 2 / 3 slot column widths)
    │  - PNG transparency → white background conversion (Pillow)
    │  - Per-slot logo heights from theme config
    │  - Font fallback: Malgun Gothic → Helvetica (with warning)
    │  - FONT_PATH environment variable override supported
```
