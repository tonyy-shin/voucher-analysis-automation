# ================================================================
#  전표분석서 자동화 프로그램 - 빌드 자동화 스크립트
#  사용법: build.bat 를 더블클릭
# ================================================================

# 콘솔 인코딩 UTF-8 설정 (한글 깨짐 방지)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# 스크립트가 있는 폴더를 작업 디렉토리로 고정
# (-File 방식: $PSScriptRoot 자동 설정 / -Command 방식: bat의 cd /d 로 이미 이동됨)
if ($PSScriptRoot) { Set-Location $PSScriptRoot }

# ── 출력 헬퍼 함수 ──────────────────────────────────────────────
function Write-Step  { param([string]$msg) Write-Host "`n[진행 중] $msg" -ForegroundColor Yellow }
function Write-Done  { param([string]$msg) Write-Host "[완료]    $msg`n" -ForegroundColor Green }
function Write-Fail  { param([string]$msg) Write-Host "`n[오류]    $msg" -ForegroundColor Red }
function Write-Info  { param([string]$msg) Write-Host "          $msg" -ForegroundColor Gray }

# ================================================================
#  시작 화면
# ================================================================
Clear-Host
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "     전표분석서 자동화 프로그램  |  빌드 자동화 스크립트       " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host " 이 스크립트는 아래 4단계를 순서대로 자동 실행합니다." -ForegroundColor White
Write-Host "   STEP 1. 이전 빌드 파일 정리" -ForegroundColor White
Write-Host "   STEP 2. PyInstaller 설치 확인" -ForegroundColor White
Write-Host "   STEP 3. 실행 파일(.exe) 생성 (1~3분 소요)" -ForegroundColor White
Write-Host "   STEP 4. 배포용 ZIP 파일 자동 생성" -ForegroundColor White
Write-Host ""


# ================================================================
#  STEP 1. 이전 빌드 파일 정리
# ================================================================
Write-Step "STEP 1/4 - 이전 빌드 파일 정리 중..."

# __pycache__ 폴더 전체 삭제
$pycaches = Get-ChildItem -Path . -Filter "__pycache__" -Recurse -Directory -ErrorAction SilentlyContinue
if ($pycaches.Count -gt 0) {
    foreach ($d in $pycaches) { Remove-Item $d.FullName -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Info "__pycache__ 폴더 $($pycaches.Count)개 삭제 완료"
} else {
    Write-Info "__pycache__ 폴더 없음 (스킵)"
}

# .pyc 파일 삭제
$pycFiles = Get-ChildItem -Path . -Filter "*.pyc" -Recurse -ErrorAction SilentlyContinue
if ($pycFiles.Count -gt 0) {
    $pycFiles | Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Info ".pyc 파일 $($pycFiles.Count)개 삭제 완료"
} else {
    Write-Info ".pyc 파일 없음 (스킵)"
}

# build/ 폴더 삭제
if (Test-Path "build") {
    Remove-Item "build" -Recurse -Force
    Write-Info "build/ 폴더 삭제 완료"
} else {
    Write-Info "build/ 폴더 없음 (스킵)"
}

# dist/ 폴더 삭제
if (Test-Path "dist") {
    Remove-Item "dist" -Recurse -Force
    Write-Info "dist/ 폴더 삭제 완료"
} else {
    Write-Info "dist/ 폴더 없음 (스킵)"
}

# .spec 파일 삭제
$specFiles = Get-ChildItem -Path . -Filter "*.spec" -ErrorAction SilentlyContinue
if ($specFiles.Count -gt 0) {
    $specFiles | Remove-Item -Force
    Write-Info ".spec 파일 $($specFiles.Count)개 삭제 완료"
} else {
    Write-Info ".spec 파일 없음 (스킵)"
}

# 기존 배포용 ZIP 삭제
if (Test-Path "전표분석서_자동화_배포용.zip") {
    Remove-Item "전표분석서_자동화_배포용.zip" -Force
    Write-Info "기존 배포용 ZIP 파일 삭제 완료"
} else {
    Write-Info "기존 배포용 ZIP 없음 (스킵)"
}

Write-Done "STEP 1 완료 - 프로젝트 폴더가 깔끔하게 정리되었습니다."


# ================================================================
#  STEP 2. PyInstaller 설치 확인 및 자동 설치
# ================================================================
Write-Step "STEP 2/4 - PyInstaller 및 필수 패키지 확인 중..."

# Python 실행 가능 여부 확인
$pythonVersion = python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Fail "Python을 찾을 수 없습니다. Python 설치 여부를 확인해 주세요."
    Write-Host ""
    Read-Host "아무 키나 누르면 창이 닫힙니다"
    exit 1
}
Write-Info "Python 버전: $pythonVersion"

# Pillow 설치 확인 (pdf_exporter.py 필수 의존성)
$pillowCheck = python -c "import PIL" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "          Pillow가 설치되어 있지 않습니다. 자동 설치 중..." -ForegroundColor Yellow
    python -m pip install Pillow --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Fail "Pillow 설치 실패. 인터넷 연결을 확인해 주세요."
        Read-Host "아무 키나 누르면 창이 닫힙니다"
        exit 1
    }
    Write-Info "Pillow 설치 완료"
} else {
    Write-Info "Pillow 설치 확인 완료"
}

# PyInstaller 설치 확인
$pyiVersion = python -m PyInstaller --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "          PyInstaller가 설치되어 있지 않습니다. 자동 설치 중..." -ForegroundColor Yellow
    python -m pip install pyinstaller --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Fail "PyInstaller 설치 실패. 인터넷 연결을 확인해 주세요."
        Read-Host "아무 키나 누르면 창이 닫힙니다"
        exit 1
    }
    Write-Info "PyInstaller 설치 완료"
} else {
    Write-Info "PyInstaller 버전: $pyiVersion"
}

# 필수 데이터 파일 존재 여부 확인
Write-Host ""
Write-Host "          [필수 파일 존재 여부 확인]" -ForegroundColor Gray
$requiredFiles = @(
    "column_config.yaml",
    "theme_config.json",
    "report_generator\src\code\main.py",
    "report_generator\src\code\templates\report_template.html"
)
$allExists = $true
foreach ($f in $requiredFiles) {
    if (Test-Path $f) {
        Write-Info "  [OK] $f"
    } else {
        Write-Host "          [누락] $f" -ForegroundColor Red
        $allExists = $false
    }
}

if (-not $allExists) {
    Write-Fail "필수 파일이 누락되었습니다. 위 목록을 확인해 주세요."
    Read-Host "아무 키나 누르면 창이 닫힙니다"
    exit 1
}

Write-Done "STEP 2 완료 - 모든 패키지와 파일 준비 완료."


# ================================================================
#  STEP 3. PyInstaller 빌드 실행
# ================================================================
Write-Step "STEP 3/4 - 실행 파일(.exe) 생성 중..."
Write-Host "          ※ 이 단계는 1~3분 정도 소요됩니다." -ForegroundColor Gray
Write-Host "          ※ 화면이 멈춘 것처럼 보여도 정상입니다. 기다려 주세요." -ForegroundColor Gray
Write-Host ""

python -m PyInstaller `
    --onefile `
    --noconsole `
    --name "전표분석서_자동화" `
    --add-data "column_config.yaml;." `
    --add-data "theme_config.json;." `
    --add-data "report_generator/src/code/templates/report_template.html;templates" `
    --add-data "images;images" `
    --paths "report_generator/src/code" `
    --exclude-module matplotlib `
    --exclude-module scipy `
    --exclude-module IPython `
    --exclude-module notebook `
    --exclude-module jupyter `
    --exclude-module sklearn `
    --exclude-module win32com `
    --exclude-module jinja2 `
    --exclude-module cryptography `
    --exclude-module tkinter.test `
    "report_generator/src/code/main.py"

# 빌드 결과 확인 - 실패 시 여기서 즉시 종료
$exePath = "dist\전표분석서_자동화.exe"
if (-not (Test-Path $exePath)) {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host "   빌드 실패                                                   " -ForegroundColor Red
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host ""
    Write-Fail "dist\전표분석서_자동화.exe 파일이 생성되지 않았습니다."
    Write-Host ""
    Write-Host "  [주요 원인 및 해결 방법]" -ForegroundColor Yellow
    Write-Host "  1. 필수 패키지 미설치" -ForegroundColor Yellow
    Write-Host "     -> 터미널에서 실행: python -m pip install -r requirements.txt" -ForegroundColor White
    Write-Host "  2. 소스 파일 경로 오류" -ForegroundColor Yellow
    Write-Host "     -> report_generator\src\code\main.py 경로 확인" -ForegroundColor White
    Write-Host "  3. 바이러스 백신 프로그램이 빌드를 차단" -ForegroundColor Yellow
    Write-Host "     -> 백신 프로그램에서 이 폴더를 예외 처리 후 재시도" -ForegroundColor White
    Write-Host ""
    Read-Host "아무 키나 누르면 창이 닫힙니다"
    exit 1
}

$exeSizeMB = [math]::Round((Get-Item $exePath).Length / 1MB, 1)
Write-Done "STEP 3 완료 - 실행 파일 생성 성공 (${exeSizeMB} MB)"


# ================================================================
#  STEP 4. 배포용 ZIP 파일 생성
# ================================================================
Write-Step "STEP 4/4 - 배포용 ZIP 파일 생성 중..."
Write-Info ".exe + 설정 파일 2개를 하나의 ZIP으로 묶습니다..."

$zipName = "전표분석서_자동화_배포용.zip"
$zipPath = Join-Path (Get-Location) $zipName

# 압축 대상 파일 목록
$filesToZip = @(
    "dist\전표분석서_자동화.exe",
    "column_config.yaml",
    "theme_config.json"
)

try {
    Compress-Archive -Path $filesToZip -DestinationPath $zipPath -Force
    $zipSizeMB   = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
    $zipFullPath = (Get-Item $zipPath).FullName
    $exeFullPath = (Get-Item $exePath).FullName

    Write-Done "STEP 4 완료 - 배포용 ZIP 생성 성공 (${zipSizeMB} MB)"

    Write-Host "================================================================" -ForegroundColor Green
    Write-Host "   전체 완료!  빌드 + 배포 패키징 모두 성공                    " -ForegroundColor Green
    Write-Host "================================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  [생성된 파일]" -ForegroundColor Cyan
    Write-Host "  실행 파일 : $exeFullPath" -ForegroundColor White
    Write-Host "             크기: ${exeSizeMB} MB" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  배포 ZIP  : $zipFullPath" -ForegroundColor White
    Write-Host "             크기: ${zipSizeMB} MB" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  [ZIP 파일 안에 들어있는 내용]" -ForegroundColor Cyan
    Write-Host "    전표분석서_자동화.exe   <- 실행 파일 (사용자가 더블클릭)" -ForegroundColor White
    Write-Host "    column_config.yaml      <- 열 설정 파일" -ForegroundColor White
    Write-Host "    theme_config.json       <- 테마 설정 파일" -ForegroundColor White
    Write-Host ""
    Write-Host "  [배포 순서]" -ForegroundColor Cyan
    Write-Host "  1. 먼저 dist 폴더의 .exe 파일을 직접 실행해서 정상 동작을 확인하세요." -ForegroundColor White
    Write-Host "  2. 테스트 통과 후 '전표분석서_자동화_배포용.zip' 파일을 전달하세요." -ForegroundColor White
    Write-Host "  3. 받는 사람은 ZIP 압축 해제 후 .exe 파일을 더블클릭하면 됩니다." -ForegroundColor White
    Write-Host ""
} catch {
    Write-Fail "ZIP 압축 중 오류가 발생했습니다: $_"
    Write-Host ""
    Write-Host "  빌드 자체는 성공했습니다! 아래 3개 파일을 직접 ZIP으로 압축하세요:" -ForegroundColor Yellow
    Write-Host "    - dist\전표분석서_자동화.exe" -ForegroundColor White
    Write-Host "    - column_config.yaml" -ForegroundColor White
    Write-Host "    - theme_config.json" -ForegroundColor White
    Write-Host ""
}

Read-Host "아무 키나 누르면 창이 닫힙니다"
