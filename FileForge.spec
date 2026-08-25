# -*- mode: python ; coding: utf-8 -*-
#
# onefile exe 는 실행할 때마다 번들 전체를 임시폴더에 푼다.
# 담긴 용량과 파일 개수가 그대로 시작 시간이 되므로, 실행에 쓰이지 않는 것은 담지 않는다.
# 기능이 하나라도 빠지면 안 되므로, 확실히 안 쓰는 것만 걸러낸다.
#
import os
from PyInstaller.utils.hooks import collect_all

# HEIC 읽기: pi-heif 가 있으면 그걸 쓴다. pillow-heif 와 읽은 결과는 같고,
# 저장용 인코더(libx265, 21MB)가 빠져 있어 훨씬 가볍다.
try:
    import pi_heif                     # noqa: F401
    HEIF_PKG = 'pi_heif'
except ImportError:
    HEIF_PKG = 'pillow_heif'

datas = []
binaries = []
hiddenimports = ['pypdf', 'fitz', 'pymupdf', HEIF_PKG, 'PIL', 'PIL.Image',
                 'win32com', 'win32com.client', 'pythoncom', 'pywintypes', 'openpyxl']
for pkg in ('pypdf', 'pymupdf', HEIF_PKG, 'PIL', 'win32com', 'openpyxl'):
    d, b, h = collect_all(pkg)
    datas += d; binaries += b; hiddenimports += h


# ---------------- 실행에 필요 없는 항목 걷어내기 ----------------
# collect_all 은 패키지 폴더를 통째로 담아서, 개발용 헤더·정적 라이브러리·
# 파이썬 원본 소스처럼 실행과 무관한 파일까지 따라 들어온다.

_DROP_SUFFIX = (
    '.pyi', '.h', '.hpp', '.lib', '.a', '.pdb', '.chm', '.exp',
    '.py',              # 모듈은 PYZ 에 바이트코드로 들어간다. 원본 소스는 중복 (415개 / 7.9MB)
)
_DROP_PARTS = (
    'mupdf-devel',      # PyMuPDF 의 C++ 개발용 헤더·lib (약 8MB)
    'pythonwin',        # pywin32 에 딸린 GUI 편집기 — 이 앱은 쓰지 않는다
    '__pycache__',      # PYZ 와 중복
)
# 확실히 안 쓰는 코덱만. HEIC 읽기에 필요한 libheif·libde265 는 절대 건드리지 않는다.
_DROP_BINARY = (
    '_avif',            # PIL 의 AVIF 코덱 (7.5MB) — 이 앱은 AVIF 를 입출력하지 않는다
)


def _prune(items, label, drop_names=()):
    kept, dropped, saved = [], 0, 0
    for item in items:
        dest = item[0].replace('\\', '/')
        low = dest.lower()
        parts = low.split('/')
        base = parts[-1]
        if low.endswith(_DROP_SUFFIX) or any(p in parts for p in _DROP_PARTS) \
                or 'mupdf-devel' in low or 'tests' in parts or 'test' in parts \
                or any(base.startswith(name) for name in drop_names):
            try:
                saved += os.path.getsize(item[1])
            except (OSError, TypeError):
                pass
            dropped += 1
            continue
        kept.append(item)
    print('[spec] %s: %d개 제외 (%.1f MB)' % (label, dropped, saved / 1048576.0))
    return kept


datas = _prune(datas, 'datas')
binaries = _prune(binaries, 'binaries')

excludes = [
    # 이 앱이 쓰지 않는데 딸려 들어오기 쉬운 것들
    'numpy', 'scipy', 'matplotlib', 'pandas',
    'pytest', 'setuptools', 'pip', 'wheel', 'pkg_resources',
    'doctest', 'pdb', 'unittest',
    'PIL.ImageQt', 'PIL.ImageTk',          # Qt·Tk 연동 이미지 표시 (미사용)
    'pythonwin', 'win32ui', 'win32uiole',  # pywin32 GUI 편집기
]
if HEIF_PKG == 'pi_heif':
    excludes.append('pillow_heif')          # 둘 다 깔려 있어도 가벼운 쪽만 담는다

a = Analysis(
    ['파일명일괄변경_윈도우용.pyw'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

# 위 목록 필터는 collect_all 이 모아온 것만 걸러낸다. 확장모듈(.pyd)과 의존 DLL 은
# Analysis 가 스스로 찾아 넣으므로, 다 끝난 뒤 한 번 더 걸러야 실제로 빠진다.
a.binaries = _prune(a.binaries, 'a.binaries', _DROP_BINARY)
a.datas = _prune(a.datas, 'a.datas')

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FileForge',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['FileForge.ico'],
)
