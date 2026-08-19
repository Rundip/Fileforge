@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title FileForge EXE 만들기

echo ============================================================
echo   FileForge - EXE 만들기
echo   최초 1회 필요한 부품을 인터넷에서 자동 설치합니다
echo ============================================================
echo.

REM ---------- 1. Python 이 있는지 확인 ----------
python --version >nul 2>&1
if errorlevel 1 (
    echo [중단] 이 PC에 Python이 설치되어 있지 않습니다.
    echo.
    echo   해결 방법: python.org 에서 Python을 설치한 뒤 다시 실행하세요.
    echo   설치 화면에서 "Add Python to PATH" 에 반드시 체크하세요.
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo  [확인] 설치된 Python: %PYVER%
echo.

REM ---------- 2. 프로그램 원본이 옆에 있는지 확인 ----------
if not exist "파일명일괄변경_윈도우용.pyw" (
    echo [중단] 프로그램 원본 파일을 찾을 수 없습니다.
    echo.
    echo   이 폴더 안에 "파일명일괄변경_윈도우용.pyw" 파일이 함께 있어야 합니다.
    echo   저장소를 받은 폴더 그대로에서 이 파일을 실행해 주세요.
    echo.
    pause
    exit /b 1
)

REM ---------- 3. 부품 설치 ----------
echo  [1/2] 필요한 부품을 설치합니다...
echo.
if exist "wheels" (
    echo  ※ wheels 폴더가 있어 인터넷 없이 설치합니다.
    python -m pip install --no-index --find-links=wheels pyinstaller pypdf Pillow pywin32 openpyxl pymupdf pillow-heif
) else (
    python -m pip install --upgrade pyinstaller pypdf Pillow pywin32 openpyxl pymupdf pillow-heif
)
if errorlevel 1 (
    echo.
    echo [중단] 부품 설치에 실패했습니다.
    echo.
    echo   이 상태로 진행하면 PDF·이미지 기능이 동작하지 않는
    echo   반쪽짜리 EXE가 만들어지므로 여기서 멈춥니다.
    echo.
    echo   가장 흔한 원인
    echo    1^) 이 PC가 인터넷에 연결되어 있지 않음
    echo    2^) 회사 보안 정책으로 외부 다운로드가 막혀 있음
    echo.
    echo   해결 방법: 담당자에게 "오프라인 빌드 꾸러미"를 요청하세요.
    echo   그 안의 wheels 폴더를 이 폴더에 복사한 뒤 다시 실행하면
    echo   인터넷 없이 설치됩니다.
    echo.
    echo   위에 나온 영문 메시지를 그대로 담당자에게 알려주세요.
    echo.
    pause
    exit /b 1
)
echo.
echo  [확인] 부품 설치 완료.
echo.

REM ---------- 4. 실제로 불러올 수 있는지 최종 점검 ----------
REM  설치는 됐는데 못 불러오는 경우가 있어, 반쪽짜리 EXE 를 막기 위해 확인한다
echo  [점검] 부품이 정상 동작하는지 확인합니다...
python -c "import pypdf, PIL, fitz, pillow_heif, openpyxl" 2>nul
if errorlevel 1 (
    echo.
    echo [중단] 부품은 설치됐지만 불러오지 못했습니다.
    echo        이대로 만들면 PDF·이미지 기능이 동작하지 않는 EXE가 됩니다.
    echo        위에 나온 영문 메시지를 담당자에게 알려주세요.
    echo.
    pause
    exit /b 1
)
REM  한글(HWP) 변환용 pywin32 는 없어도 빌드 가능 - LibreOffice 로 대체되므로 경고만
python -c "import win32com.client" 2>nul
if errorlevel 1 (
    echo  [참고] 한글 문서 변환용 부품을 불러오지 못했습니다.
    echo         빌드는 그대로 진행합니다. 한글 문서는 LibreOffice로 변환됩니다.
) else (
    echo  [확인] 부품 정상.
)
echo.

REM ---------- 5. EXE 빌드 ----------
echo  [2/2] EXE를 만듭니다. 3~10분 정도 걸립니다. 창을 닫지 마세요.
echo.
if exist "FileForge.ico" (set ICONOPT=--icon FileForge.ico) else (set ICONOPT=)
python -m PyInstaller --onefile --noconsole --name FileForge %ICONOPT% ^
  --hidden-import pypdf --collect-all pypdf ^
  --hidden-import fitz --hidden-import pymupdf --collect-all pymupdf ^
  --hidden-import pillow_heif --collect-all pillow_heif ^
  --hidden-import PIL --hidden-import PIL.Image --collect-all PIL ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import pythoncom --hidden-import pywintypes --collect-all win32com ^
  --hidden-import openpyxl --collect-all openpyxl ^
  파일명일괄변경_윈도우용.pyw

echo.
echo ============================================================
if exist "dist\FileForge.exe" (
    echo  [완료] EXE가 만들어졌습니다.
    echo.
    echo    위치: %CD%\dist\FileForge.exe
    echo.
    echo    이 파일 하나만 복사하면 Python이 없는 PC에서도 실행됩니다.
    echo    실행할 때 인터넷은 필요 없습니다.
    echo.
    echo    참고 1. 엑셀을 PDF로 바꾸려면 LibreOffice를 따로 설치하세요.
    echo    참고 2. 한글 문서는 PC에 한글이 깔려 있으면 자동으로 사용합니다.
    echo.
    choice /c YN /n /m "  지금 dist 폴더를 열어볼까요? Y=예 / N=아니오 : "
    if not errorlevel 2 explorer "%CD%\dist"
) else (
    echo  [실패] EXE가 만들어지지 않았습니다.
    echo         위에 나온 영문 메시지를 그대로 담당자에게 알려주세요.
)
echo ============================================================
echo.
pause
