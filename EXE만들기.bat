@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  파일명 일괄 변경 도구 EXE 빌드
echo  (최초 1회 필요한 라이브러리를 자동 설치합니다)
echo ============================================
python -m pip install --upgrade pyinstaller pypdf Pillow pywin32 openpyxl pymupdf pillow-heif

echo.
echo 빌드 중... (PDF/이미지변환(pymupdf) + HEIC(pillow-heif) + 한글변환(pywin32) + 엑셀(openpyxl) 포함)
python -m PyInstaller --onefile --noconsole --name FileForge --icon FileForge.ico ^
  --hidden-import pypdf --collect-all pypdf ^
  --hidden-import fitz --hidden-import pymupdf --collect-all pymupdf ^
  --hidden-import pillow_heif --collect-all pillow_heif ^
  --hidden-import PIL --hidden-import PIL.Image --collect-all PIL ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import pythoncom --hidden-import pywintypes --collect-all win32com ^
  --hidden-import openpyxl --collect-all openpyxl ^
  파일명일괄변경_윈도우용.pyw

echo.
if exist "dist\FileForge.exe" (
    echo [완료] dist\FileForge.exe 생성됨
    echo  ※ 엑셀→PDF 기능을 쓰려면 LibreOffice를 따로 설치하세요.
    echo  ※ 한글→PDF 기능은 PC에 한글(아래아한글)이 설치돼 있으면 자동 사용됩니다.
) else (
    echo [실패] 빌드 중 오류가 발생했습니다. 위 메시지를 확인하세요.
)
pause
