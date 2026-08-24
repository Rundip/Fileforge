"""
FileForge — 파일 자동화 도구 (Windows 에디션, Nike 디자인 시스템 적용)

- 실행하면 로컬 서버가 뜨고 기본 브라우저에서 UI가 열림
- 폴더 선택은 Windows 폴더 선택 다이얼로그 사용
- PDF 병합/분할 등은 최초 사용 시 pypdf·Pillow를 자동 설치합니다

실행 방법: Python 설치 후 이 파일을 더블클릭
"""

import http.client
import importlib
import json
import os
import queue
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HOST = "127.0.0.1"
BASE_PORT = 8765          # 여기서부터 비어 있는 포트를 찾는다

# ---------------- 파일 작업 로직 ----------------

UNDO_STACK = []  # [{"moves": [(현재경로, 이전경로)], "dirs": [생성폴더]}]
LOCK = threading.Lock()

# ---------------- UI 수명 주기 (탭 닫으면 서버·프로세스 완전 종료) ----------------
ACTIVE_TABS = {}          # tab_id -> 만료시각(epoch). 비면 열린 UI 없음 → 종료
TABS_LOCK = threading.Lock()
STOP = threading.Event()  # 서버 종료 신호
SERVER = None             # 실행 중인 HTTP 서버(전역 참조)
_ARMED = False            # 브라우저가 최소 1번 연결됐는지(시작 직후 오종료 방지)
_SHUTTING = False         # 종료 절차 중복 실행 방지
HEARTBEAT_TTL = 70        # ping이 없으면 이 초 뒤 탭 만료(백그라운드 스로틀 여유)
CLOSE_GRACE = 3           # 닫힘 신호 후 유예(새로고침 재연결 대기 → 오종료 방지)
ARM_TIMEOUT = 180         # 이 시간 안에 UI가 한 번도 안 열리면 스스로 종료(좀비 프로세스 방지)
                          # 느린 PC에서 브라우저가 뜨는 시간을 넉넉히 감안한 값
_START_TS = 0.0           # 서버 시작 시각


def notify_shell_change(paths):
    """파일을 옮기거나 만든 뒤 탐색기(Windows Explorer)가 즉시 새로고침되도록
    셸에 변경 사실을 알린다.

    파이썬의 os.rename/shutil.move 등은 셸에 알림을 보내지 않아서, 알림 없이는
    탐색기가 한동안 옛 목록을 그대로 보여줄 수 있다(수동 F5 필요). SHChangeNotify
    로 해당 폴더가 바뀌었음을 알려 자동 새로고침되게 한다. (Windows 전용)"""
    if os.name != "nt":
        return
    try:
        import ctypes
        SHCNE_UPDATEDIR = 0x00002000          # 폴더 내용이 바뀜
        SHCNF_PATHW = 0x0005                  # 경로(유니코드)로 지정
        SHCNF_FLUSH = 0x1000                  # 알림을 즉시 전달
        seen = set()
        for p in paths:
            if not p:
                continue
            d = os.path.abspath(p)
            key = d.lower()
            if key in seen or not os.path.isdir(d):
                continue
            seen.add(key)
            ctypes.windll.shell32.SHChangeNotify(
                SHCNE_UPDATEDIR, SHCNF_PATHW | SHCNF_FLUSH,
                ctypes.c_wchar_p(d), None)
    except Exception:
        pass  # 알림 실패는 기능에 영향 없음 (최악의 경우 수동 새로고침)


def safe_join(folder, rel):
    """folder 바깥을 벗어나는 경로 차단"""
    target = os.path.realpath(os.path.join(folder, rel))
    base = os.path.realpath(folder)
    if not target.startswith(base + os.sep):
        raise ValueError(f"허용되지 않는 경로: {rel}")
    return target


def list_files(folder, exts, include="files"):
    """폴더 안의 항목 목록.

    include — "files": 파일만 / "dirs": 폴더만 / "both": 파일+폴더.
    이름 변경 규칙은 폴더에도 그대로 쓸 수 있어서, 폴더도 목록에 담아 돌려준다.
    (폴더는 확장자가 없으므로 확장자 필터의 영향을 받지 않는다)"""
    want_files = include in ("files", "both")
    want_dirs = include in ("dirs", "both")
    # scandir 은 디렉터리를 한 번 훑으면서 종류·수정시각을 함께 돌려준다.
    # listdir + isdir + isfile + getmtime 은 항목마다 stat 을 3~4번 호출해서,
    # 파일이 많은 폴더나 네트워크 드라이브에서 눈에 띄게 느렸다.
    items, all_names = [], []
    with os.scandir(folder) as it:
        for entry in it:
            name = entry.name
            if name.startswith("."):
                continue
            all_names.append(name)
            try:
                is_dir = entry.is_dir()
            except OSError:
                continue
            if is_dir:
                if not want_dirs:
                    continue
            else:
                if not want_files or not entry.is_file():
                    continue
                ext = os.path.splitext(name)[1].lower().lstrip(".")
                if exts and ext not in exts:
                    continue
            try:
                mtime = entry.stat().st_mtime
            except OSError:
                mtime = 0
            items.append({"name": name, "mtime": mtime, "isDir": is_dir})
    # 폴더를 먼저, 그 다음 파일 — 각각 사람이 보기 좋은 자연 정렬(숫자 순)
    items.sort(key=lambda f: (not f["isDir"], _natkey(f["name"])))
    return items, all_names


def execute_jobs(folder, jobs, mode, folders=None):
    """jobs: [{"src": 파일명, "dst": 상대경로, "d": 폴더번호}] — 검증 후 실행, undo 배치 반환

    folders 를 주면 각 job 의 "d" 가 가리키는 폴더를 기준으로 삼는다.
    (키워드별 지정값으로 변경을 여러 폴더에 한꺼번에 적용할 때 사용)"""
    resolved = []
    for j in jobs:
        src = j["src"]
        if os.sep in src or src in (".", ".."):
            raise ValueError(f"잘못된 파일명: {src}")
        base = folder
        if folders:
            try:
                base = folders[int(j.get("d", 0))]
            except (TypeError, ValueError, IndexError):
                raise ValueError(f"폴더 지정이 잘못되었습니다: {src}")
        resolved.append((safe_join(base, src), safe_join(base, j["dst"])))

    batch, made_dirs, errors = [], [], []
    if mode == "organize":
        for src, dst in resolved:
            try:
                if os.path.exists(dst):
                    errors.append(f"{os.path.basename(src)}: 대상에 이미 존재함")
                    continue
                d = os.path.dirname(dst)
                if not os.path.isdir(d):
                    os.makedirs(d, exist_ok=True)
                    made_dirs.append(d)
                shutil.move(src, dst)
                batch.append((dst, src))
            except OSError as e:
                errors.append(f"{os.path.basename(src)}: {e}")
    else:
        # 2단계 이름 변경: 임시 이름을 거쳐 맞바꾸기 충돌도 안전하게 처리
        temps = []
        try:
            for src, dst in resolved:
                tmp = src + ".tmp_" + uuid.uuid4().hex[:8]
                os.rename(src, tmp)
                temps.append((tmp, src, dst))
        except OSError as e:
            for tmp, src, _ in temps:
                os.rename(tmp, src)
            raise RuntimeError(f"이름 변경 실패 — 모두 원복했습니다. ({e})")
        for tmp, src, dst in temps:
            try:
                if os.path.exists(dst):
                    os.rename(tmp, src)
                    errors.append(f"{os.path.basename(src)}: 대상에 이미 존재함")
                else:
                    os.rename(tmp, dst)
                    batch.append((dst, src))
            except OSError as e:
                os.rename(tmp, src)
                errors.append(f"{os.path.basename(src)}: {e}")

    if batch:
        UNDO_STACK.append({"moves": batch, "dirs": made_dirs})
    dirs = set(folders) if folders else {folder}
    for dst, src in batch:
        dirs.add(os.path.dirname(dst))
        dirs.add(os.path.dirname(src))
    dirs.update(made_dirs)
    notify_shell_change(dirs)
    return len(batch), errors


def undo_last():
    if not UNDO_STACK:
        return 0, []
    last = UNDO_STACK.pop()
    errors = []
    for cur, old in reversed(last["moves"]):
        try:
            os.makedirs(os.path.dirname(old), exist_ok=True)
            shutil.move(cur, old)
        except OSError as e:
            errors.append(f"{os.path.basename(cur)}: {e}")
    removed_dirs = 0
    for d in reversed(last["dirs"]):
        try:
            os.rmdir(d)  # 비어 있을 때만 삭제됨
            removed_dirs += 1
        except OSError:
            pass
    dirs = set()
    for cur, old in last["moves"]:
        dirs.add(os.path.dirname(cur))
        dirs.add(os.path.dirname(old))
    dirs.update(last["dirs"])
    notify_shell_change(dirs)
    return len(last["moves"]), errors, removed_dirs


def make_dirs(folder, names, moves=None):
    """폴더를 한꺼번에 만든다. (선택적으로 파일을 그 폴더로 이동)

    - 이미 있는 폴더는 건너뛴다(덮어쓰지 않음)
    - 일부가 실패해도 나머지는 계속 진행하고, 실패 사유를 모아서 돌려준다
    - 만든 폴더·옮긴 파일을 undo 배치 1개로 묶어 '실행 취소'로 한 번에 되돌릴 수 있다"""
    base = os.path.realpath(folder)
    created_dirs = []   # 실제로 새로 만들어진 모든 폴더(중간 경로 포함) — 실행 취소용
    made, skipped, errors = 0, [], []

    for rel in names:
        rel = (rel or "").strip().strip("/\\")
        if not rel:
            continue
        try:
            path = safe_join(folder, rel)
        except ValueError as e:
            errors.append(str(e))
            continue
        if os.path.isdir(path):
            skipped.append(rel)
            continue
        if os.path.exists(path):
            errors.append(f"{rel}: 같은 이름의 파일이 이미 있어 폴더를 만들 수 없습니다")
            continue
        # 새로 생기는 상위 경로까지 기록해 둬야 실행 취소가 깨끗하게 된다
        missing, p = [], path
        while os.path.realpath(p) != base and not os.path.isdir(p):
            missing.append(p)
            p = os.path.dirname(p)
        try:
            os.makedirs(path, exist_ok=True)
            created_dirs.extend(reversed(missing))   # 상위 → 하위 순서
            made += 1
        except OSError as e:
            errors.append(f"{rel}: {e}")

    move_batch = []
    for m in (moves or []):
        src_name = m.get("src", "")
        if os.sep in src_name or "/" in src_name or src_name in (".", ".."):
            errors.append(f"잘못된 파일명: {src_name}")
            continue
        try:
            src = safe_join(folder, src_name)
            dst = safe_join(folder, m.get("dst", ""))
        except ValueError as e:
            errors.append(str(e))
            continue
        try:
            if not os.path.exists(src):
                errors.append(f"{src_name}: 파일을 찾을 수 없습니다")
                continue
            if os.path.exists(dst):
                errors.append(f"{src_name}: 대상 폴더에 같은 이름이 이미 있습니다")
                continue
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.move(src, dst)
            move_batch.append((dst, src))
        except OSError as e:
            errors.append(f"{src_name}: {e}")

    if created_dirs or move_batch:
        UNDO_STACK.append({"moves": move_batch, "dirs": created_dirs})

    notify = {folder}
    notify.update(created_dirs)
    for dst, src in move_batch:
        notify.add(os.path.dirname(dst))
        notify.add(os.path.dirname(src))
    notify_shell_change(notify)
    return {"made": made, "skipped": len(skipped), "skippedNames": skipped[:20],
            "moved": len(move_batch), "errors": errors}


def open_in_explorer(path):
    """결과를 바로 확인할 수 있게 탐색기(Finder)에서 폴더를 연다."""
    if not path or not os.path.isdir(path):
        raise ValueError("폴더를 찾을 수 없습니다.")
    try:
        if os.name == "nt":
            os.startfile(path)                       # Windows 탐색기
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])         # macOS Finder
        else:
            subprocess.Popen(["xdg-open", path])     # Linux
    except OSError as e:
        raise RuntimeError(f"폴더를 열지 못했습니다: {e}")


# ---------------- PDF · 변환 도구 ----------------

def ensure_pkg(mod_name, pip_name=None):
    """필요한 라이브러리를 import; 없으면 pip로 자동 설치 후 재시도."""
    try:
        return importlib.import_module(mod_name)
    except ImportError:
        pass
    if getattr(sys, "frozen", False):  # 빌드된 앱: 라이브러리가 내장돼 있어야 함
        raise RuntimeError(f"필요한 구성요소({pip_name or mod_name})가 빠져 있습니다. 앱을 다시 설치해 주세요.")
    pip_name = pip_name or mod_name
    proc = subprocess.run([sys.executable, "-m", "pip", "install", "--user", pip_name],
                          capture_output=True, text=True, timeout=600)
    importlib.invalidate_caches()
    try:
        return importlib.import_module(mod_name)
    except ImportError:
        raise RuntimeError(f"{pip_name} 설치에 실패했습니다. 터미널에서 'pip3 install {pip_name}' 실행 후 다시 시도하세요.\n"
                           + (proc.stderr or "")[:300])


def _natkey(name):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def _unique_path(folder, name):
    """folder 안에서 겹치지 않는 파일 경로 (겹치면 _1, _2 …)"""
    base, ext = os.path.splitext(name)
    cand = os.path.join(folder, name)
    i = 1
    while os.path.exists(cand):
        cand = os.path.join(folder, f"{base}_{i}{ext}")
        i += 1
    return cand


def _select_paths(folder, files, exts, order="asc"):
    """선택된 파일명들을 검증·정렬해 [(이름, 절대경로)] 반환. files 비면 폴더 내 해당 확장자 전체."""
    if files:
        items = []
        for n in files:
            if os.sep in n or n in (".", ".."):
                raise ValueError(f"잘못된 파일명: {n}")
            p = safe_join(folder, n)
            if not os.path.isfile(p):
                raise ValueError(f"파일을 찾을 수 없습니다: {n}")
            items.append((n, p))
    else:
        listed, _ = list_files(folder, set(exts))
        items = [(f["name"], os.path.join(folder, f["name"])) for f in listed]
    items.sort(key=lambda it: _natkey(it[0]), reverse=(order == "desc"))
    return items


def find_soffice():
    """LibreOffice 실행 파일 경로 탐색 (없으면 None)"""
    cands = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in cands:
        if os.path.isfile(c):
            return c
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    return None


def find_gs():
    for name in ("gs", "gswin64c", "gswin32c"):
        p = shutil.which(name)
        if p:
            return p
    return None


def _parse_page_ranges(spec, total, fname):
    """'1-5, 7-8' → [0,1,2,3,4,6,7] (0-기반 인덱스). 빈 값이면 None(=전체 페이지)."""
    spec = (spec or "").strip()
    if not spec:
        return None
    pages = []
    for part in spec.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            a, _, b = part.partition("-")
            try:
                a, b = int(a), int(b)
            except ValueError:
                raise RuntimeError(f"{fname}: 페이지 범위 형식이 잘못되었습니다 → '{part}' (예: 1-5)")
            if a < 1 or b < 1 or a > b:
                raise RuntimeError(f"{fname}: 페이지 범위가 잘못되었습니다 → '{part}'")
            if b > total:
                raise RuntimeError(f"{fname}: {total}페이지뿐인데 {b}페이지를 요청했습니다.")
            pages.extend(range(a - 1, b))
        else:
            try:
                n = int(part)
            except ValueError:
                raise RuntimeError(f"{fname}: 페이지 번호가 잘못되었습니다 → '{part}'")
            if n < 1 or n > total:
                raise RuntimeError(f"{fname}: {total}페이지뿐인데 {n}페이지를 요청했습니다.")
            pages.append(n - 1)
    return pages or None


def pdf_merge(folder, files, order, out_name, ranges=None):
    pypdf = ensure_pkg("pypdf")
    items = _select_paths(folder, files, ["pdf"], order)
    if len(items) < 2:
        raise RuntimeError("병합하려면 PDF를 2개 이상 선택하세요.")
    ranges = ranges or {}
    writer = pypdf.PdfWriter()
    total_pages = 0
    for name, path in items:
        try:
            reader = pypdf.PdfReader(path)
            sel = _parse_page_ranges(ranges.get(name), len(reader.pages), name)
            if sel is None:
                for page in reader.pages:
                    writer.add_page(page)
                    total_pages += 1
            else:
                for idx in sel:
                    writer.add_page(reader.pages[idx])
                    total_pages += 1
        except RuntimeError:
            raise  # 페이지 범위 오류는 그대로 전달
        except Exception as e:
            raise RuntimeError(f"{name} 읽기 실패: {e}")
    out_name = (out_name or "병합").strip()
    if not out_name.lower().endswith(".pdf"):
        out_name += ".pdf"
    out_path = _unique_path(folder, out_name)
    with open(out_path, "wb") as fh:
        writer.write(fh)
    return {"output": os.path.basename(out_path), "count": len(items),
            "pages": total_pages, "order": order}


def pdf_split(folder, file, pages_per):
    pypdf = ensure_pkg("pypdf")
    if not file or os.sep in file or file in (".", ".."):
        raise ValueError("분할할 PDF를 1개 선택하세요.")
    path = safe_join(folder, file)
    if not os.path.isfile(path):
        raise ValueError("파일을 찾을 수 없습니다.")
    try:
        pages_per = int(pages_per)
    except (TypeError, ValueError):
        raise RuntimeError("분할 기준 페이지 수를 숫자로 입력하세요.")
    if pages_per < 1:
        raise RuntimeError("분할 기준 페이지 수는 1 이상이어야 합니다.")
    reader = pypdf.PdfReader(path)
    total = len(reader.pages)
    if total <= pages_per:
        raise RuntimeError(f"전체 {total}페이지뿐이라 {pages_per}페이지 기준으로 나눌 필요가 없습니다.")
    base = os.path.splitext(os.path.basename(path))[0]
    outdir = _unique_path(folder, base + "_분할")
    os.makedirs(outdir, exist_ok=True)
    parts = 0
    for idx, start in enumerate(range(0, total, pages_per), 1):
        writer = pypdf.PdfWriter()
        end = min(start + pages_per, total)
        for p in range(start, end):
            writer.add_page(reader.pages[p])
        part = os.path.join(outdir, f"{base}_{idx:03d}_p{start + 1}-{end}.pdf")
        with open(part, "wb") as fh:
            writer.write(fh)
        parts += 1
    return {"folder": os.path.basename(outdir), "parts": parts, "total_pages": total}


def pdf_delete_pages(folder, file, spec):
    """PDF에서 지정한 페이지/구간(예: '1-3, 5, 8-10')을 삭제한 새 PDF를 만든다. (원본 유지)"""
    pypdf = ensure_pkg("pypdf")
    if not file or os.sep in file or file in (".", ".."):
        raise ValueError("페이지를 삭제할 PDF를 1개 선택하세요.")
    path = safe_join(folder, file)
    if not os.path.isfile(path):
        raise ValueError("파일을 찾을 수 없습니다.")
    reader = pypdf.PdfReader(path)
    total = len(reader.pages)
    sel = _parse_page_ranges(spec, total, os.path.basename(path))
    if not sel:
        raise RuntimeError("삭제할 페이지를 입력하세요. (예: 1-3, 5, 8-10)")
    delset = set(sel)
    keep = [i for i in range(total) if i not in delset]
    if not keep:
        raise RuntimeError("모든 페이지를 삭제할 수는 없습니다. 최소 1페이지는 남겨야 합니다.")
    writer = pypdf.PdfWriter()
    for i in keep:
        writer.add_page(reader.pages[i])
    base, ext = os.path.splitext(os.path.basename(path))
    out_path = _unique_path(folder, f"{base}_페이지삭제{ext}")
    with open(out_path, "wb") as fh:
        writer.write(fh)
    return {"output": os.path.basename(out_path), "removed": len(delset),
            "remaining": len(keep), "total": total}


def images_to_pdf(folder, files, order, out_name):
    ensure_pkg("PIL", "Pillow")
    from PIL import Image
    items = _select_paths(folder, files,
                          ["jpg", "jpeg", "png", "bmp", "tif", "tiff", "gif", "webp"], order)
    if not items:
        raise RuntimeError("이미지 파일을 선택하세요.")
    imgs = []
    for name, path in items:
        try:
            im = Image.open(path)
            imgs.append(im.convert("RGB") if im.mode != "RGB" else im)
        except Exception as e:
            raise RuntimeError(f"{name} 열기 실패: {e}")
    out_name = (out_name or "이미지병합").strip()
    if not out_name.lower().endswith(".pdf"):
        out_name += ".pdf"
    out_path = _unique_path(folder, out_name)
    imgs[0].save(out_path, "PDF", save_all=True, append_images=imgs[1:])
    return {"output": os.path.basename(out_path), "count": len(imgs)}


def pdf_to_images(folder, files, fmt, dpi):
    """선택한 PDF의 각 페이지를 이미지(JPG/PNG)로 저장. 파일명은 원본 그대로 유지하고,
    여러 페이지면 이름_1, 이름_2… 로 저장한다. (원본 PDF는 그대로 둠)"""
    fitz = ensure_pkg("fitz", "pymupdf")
    ensure_pkg("PIL", "Pillow")
    from PIL import Image
    f = (fmt or "jpg").lower()
    ext = "png" if f == "png" else ("jpeg" if f == "jpeg" else "jpg")
    save_fmt = "PNG" if ext == "png" else "JPEG"
    try:
        dpi = int(dpi)
    except (TypeError, ValueError):
        dpi = 150
    dpi = max(50, min(600, dpi))
    items = _select_paths(folder, files, ["pdf"])
    if not items:
        raise RuntimeError("PDF 파일을 선택하세요.")
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)   # 72dpi 기준 확대 배율
    results = []
    total = 0
    for name, path in items:
        try:
            doc = fitz.open(path)
            n = doc.page_count
            base = os.path.splitext(os.path.basename(path))[0]
            width = len(str(n))          # 페이지 번호 자릿수 (예: 12페이지 → 2자리)
            count = 0
            for i in range(n):
                pix = doc.load_page(i).get_pixmap(matrix=mat, alpha=False)
                im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                out_name = f"{base}.{ext}" if n == 1 else f"{base}_{str(i + 1).zfill(width)}.{ext}"
                out_path = _unique_path(folder, out_name)
                if save_fmt == "JPEG":
                    im.save(out_path, "JPEG", quality=90)
                else:
                    im.save(out_path, "PNG")
                count += 1
            doc.close()
            total += count
            results.append({"name": name, "count": count})
        except Exception as e:
            results.append({"name": name, "error": str(e)})
    return {"results": results, "fmt": ext, "total": total}


def _save_image_as(im, out_path, save_fmt):
    """PIL 이미지를 JPEG/PNG로 저장. JPEG는 투명 배경을 흰색으로 평면화."""
    from PIL import Image
    if save_fmt == "JPEG":
        if im.mode in ("RGBA", "LA", "P"):
            im = im.convert("RGBA")
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bg.paste(im, mask=im.split()[-1])
            im = bg
        elif im.mode != "RGB":
            im = im.convert("RGB")
        im.save(out_path, "JPEG", quality=90)
    else:  # PNG
        if im.mode == "P":
            im = im.convert("RGBA")
        im.save(out_path, "PNG")


def _pdf_pages_to_images(path, folder, ext, save_fmt, dpi):
    """PDF 한 개의 모든 페이지를 이미지로 저장하고 저장된 파일명 목록 반환.
    이름은 원본 그대로, 여러 페이지면 이름_1, 이름_2… 로 저장."""
    fitz = ensure_pkg("fitz", "pymupdf")
    from PIL import Image
    doc = fitz.open(path)
    n = doc.page_count
    base = os.path.splitext(os.path.basename(path))[0]
    width = len(str(n))
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    outs = []
    try:
        for i in range(n):
            pix = doc.load_page(i).get_pixmap(matrix=mat, alpha=False)
            im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            out_name = f"{base}.{ext}" if n == 1 else f"{base}_{str(i + 1).zfill(width)}.{ext}"
            out_path = _unique_path(folder, out_name)
            _save_image_as(im, out_path, save_fmt)
            outs.append(os.path.basename(out_path))
    finally:
        doc.close()
    return outs


def images_convert(folder, files, fmt, dpi=150):
    """선택한 이미지·PDF를 JPG/JPEG/PNG 중 고른 형식으로 일괄 변환한다.
    파일명(확장자 제외)은 그대로 두고, PDF는 페이지별로 이름_1, 이름_2… 로 저장한다.
    (원본은 유지 · HEIC/HEIF도 지원 · 투명 배경은 흰색으로 채움)"""
    ensure_pkg("PIL", "Pillow")
    from PIL import Image
    f = (fmt or "jpg").lower()
    if f not in ("jpg", "jpeg", "png"):
        raise RuntimeError("형식은 JPG, JPEG, PNG 중에서 선택하세요.")
    save_fmt = "PNG" if f == "png" else "JPEG"
    try:
        dpi = int(dpi)
    except (TypeError, ValueError):
        dpi = 150
    dpi = max(50, min(600, dpi))
    items = _select_paths(folder, files,
                          ["jpg", "jpeg", "png", "bmp", "gif", "tif", "tiff", "webp",
                           "heic", "heif", "pdf"])
    if not items:
        raise RuntimeError("이미지 또는 PDF 파일을 선택하세요.")
    # HEIC/HEIF가 포함돼 있으면 pillow-heif로 열기 지원 등록
    if any(os.path.splitext(p)[1].lower() in (".heic", ".heif") for _, p in items):
        heif = ensure_pkg("pillow_heif", "pillow-heif")
        heif.register_heif_opener()
    results = []
    ok = 0
    for name, path in items:
        try:
            if path.lower().endswith(".pdf"):
                outs = _pdf_pages_to_images(path, folder, f, save_fmt, dpi)
                ok += 1
                results.append({"name": name, "output": outs[0] if outs else "",
                                "count": len(outs)})
                continue
            im = Image.open(path)
            base = os.path.splitext(os.path.basename(path))[0]
            out_path = _unique_path(folder, f"{base}.{f}")
            _save_image_as(im, out_path, save_fmt)
            ok += 1
            results.append({"name": name, "output": os.path.basename(out_path), "count": 1})
        except Exception as e:
            results.append({"name": name, "error": str(e)})
    return {"results": results, "fmt": f, "count": ok}


def compress_pdf(folder, files, level):
    items = _select_paths(folder, files, ["pdf"])
    if not items:
        raise RuntimeError("PDF 파일을 선택하세요.")
    gs = find_gs()
    results = []
    for name, path in items:
        base, ext = os.path.splitext(name)
        out_path = _unique_path(folder, f"{base}_압축{ext}")
        before = os.path.getsize(path)
        try:
            if gs:
                setting = {"high": "/screen", "medium": "/ebook", "low": "/printer"}.get(level, "/ebook")
                r = subprocess.run([gs, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                                    f"-dPDFSETTINGS={setting}", "-dNOPAUSE", "-dBATCH", "-dQUIET",
                                    f"-sOutputFile={out_path}", path],
                                   capture_output=True, text=True, timeout=600)
                if r.returncode != 0 or not os.path.exists(out_path):
                    results.append({"name": name, "error": "압축 실패"})
                    continue
            else:
                pypdf = ensure_pkg("pypdf")
                reader = pypdf.PdfReader(path)
                writer = pypdf.PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                # compress_content_streams()는 PdfWriter에 속한 페이지에만 호출 가능
                # (최신 pypdf에서 reader 페이지에 직접 호출하면 ValueError 발생)
                for page in writer.pages:
                    page.compress_content_streams()
                with open(out_path, "wb") as fh:
                    writer.write(fh)
            after = os.path.getsize(out_path)
            results.append({"name": name, "output": os.path.basename(out_path),
                            "before": before, "after": after,
                            "saved": round((1 - after / before) * 100, 1) if before else 0})
        except Exception as e:
            results.append({"name": name, "error": str(e)})
    return {"results": results, "engine": "ghostscript" if gs else "pypdf"}


def excel_to_pdf(folder, files):
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("엑셀→PDF 변환에는 LibreOffice(무료)가 필요합니다.\n"
                           "https://www.libreoffice.org/download 에서 설치하거나, "
                           "맥은 터미널에서 'brew install --cask libreoffice' 후 다시 시도하세요.")
    items = _select_paths(folder, files, ["xlsx", "xls", "xlsm", "ods", "csv"])
    if not items:
        raise RuntimeError("엑셀 파일을 선택하세요.")
    results = []
    for name, path in items:
        try:
            with tempfile.TemporaryDirectory() as td:
                r = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", td, path],
                                   capture_output=True, text=True, timeout=600)
                produced = os.path.join(td, os.path.splitext(os.path.basename(path))[0] + ".pdf")
                if os.path.exists(produced):
                    dst = _unique_path(folder, os.path.splitext(name)[0] + ".pdf")
                    shutil.move(produced, dst)
                    results.append({"name": name, "output": os.path.basename(dst)})
                else:
                    results.append({"name": name, "error": (r.stderr or "변환 실패")[:200]})
        except Exception as e:
            results.append({"name": name, "error": str(e)})
    return {"results": results}


def find_hwp_com():
    """한글(Hancom Office) COM 객체 생성 시도. 실패하면 None.
    윈도우 + 한글 설치 + pywin32 가 있어야 동작한다."""
    if os.name != "nt":
        return None
    try:
        import win32com.client as win32
    except ImportError:
        try:
            ensure_pkg("win32com.client", "pywin32")
            import win32com.client as win32
        except Exception:
            return None
    try:
        hwp = win32.Dispatch("HWPFrame.HwpObject")
    except Exception:
        return None  # 한글이 설치돼 있지 않음
    # 파일 열기 시 뜨는 자동화 보안 경고창 억제 (보안 모듈이 등록돼 있을 때 동작)
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        pass
    return hwp


def _hwp_save_pdf(hwp, out_path):
    """현재 열려 있는 한글 문서를 PDF로 저장. 성공 시 True."""
    # 방법 1: SaveAs 직접 호출 (대부분의 한글 버전에서 동작)
    try:
        hwp.SaveAs(out_path, "PDF")
        if os.path.exists(out_path):
            return True
    except Exception:
        pass
    # 방법 2: FileSaveAsPdf 액션 (방법 1이 안 먹는 버전 대비)
    try:
        pset = hwp.HParameterSet.HFileOpenSave
        hwp.HAction.GetDefault("FileSaveAsPdf", pset.HSet)
        pset.filename = out_path
        pset.Format = "PDF"
        hwp.HAction.Execute("FileSaveAsPdf", pset.HSet)
    except Exception:
        pass
    return os.path.exists(out_path)


def hwp_to_pdf(folder, files):
    """한글 문서(.hwp/.hwpx)를 PDF로 변환.
    1순위: 한글(Hancom Office) COM 자동화 — 서식을 그대로 재현
    2순위: LibreOffice — 한글이 없을 때 (복잡한 서식은 일부 차이가 날 수 있음)
    """
    items = _select_paths(folder, files, ["hwp", "hwpx"])
    if not items:
        raise RuntimeError("한글 파일(.hwp/.hwpx)을 선택하세요.")

    hwp = find_hwp_com()
    if hwp is not None:
        results = []
        try:
            for name, path in items:
                try:
                    out_path = _unique_path(folder, os.path.splitext(name)[0] + ".pdf")
                    hwp.Open(path, "", "forceopen:true")
                    if _hwp_save_pdf(hwp, out_path):
                        results.append({"name": name, "output": os.path.basename(out_path)})
                    else:
                        results.append({"name": name, "error": "PDF 저장 실패"})
                    try:
                        hwp.Clear(1)  # 저장하지 않고 문서 닫기
                    except Exception:
                        pass
                except Exception as e:
                    results.append({"name": name, "error": str(e)})
        finally:
            try:
                hwp.Quit()
            except Exception:
                pass
        return {"results": results, "engine": "hancom"}

    # ----- 한글이 없으면 LibreOffice 폴백 -----
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("한글→PDF 변환에는 한글(아래아한글) 또는 LibreOffice가 필요합니다.\n"
                           "· 한글이 설치돼 있으면 자동으로 사용합니다.\n"
                           "· 없으면 https://www.libreoffice.org/download 에서 LibreOffice(무료)를 설치한 뒤 다시 시도하세요.")
    results = []
    for name, path in items:
        try:
            with tempfile.TemporaryDirectory() as td:
                r = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", td, path],
                                   capture_output=True, text=True, timeout=600)
                produced = os.path.join(td, os.path.splitext(os.path.basename(path))[0] + ".pdf")
                if os.path.exists(produced):
                    dst = _unique_path(folder, os.path.splitext(name)[0] + ".pdf")
                    shutil.move(produced, dst)
                    results.append({"name": name, "output": os.path.basename(dst)})
                else:
                    results.append({"name": name,
                                    "error": (r.stderr or "변환 실패 (이 한글 형식은 LibreOffice가 지원하지 않을 수 있습니다)")[:200]})
        except Exception as e:
            results.append({"name": name, "error": str(e)})
    return {"results": results, "engine": "libreoffice"}


# ---------------- 엑셀 합치기 도구 ----------------

def _sanitize_sheet_title(title, used):
    """엑셀 시트 이름 규칙(31자·금지문자)에 맞추고, 중복되면 _2, _3… 으로 구분."""
    bad = set('\\/?*[]:')
    t = "".join("_" if c in bad else c for c in (title or "Sheet")).strip() or "Sheet"
    t = t[:31]
    base, i = t, 2
    while t.lower() in used:
        suffix = f"_{i}"
        t = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(t.lower())
    return t


def _copy_sheet_content(src_ws, dst_ws):
    """워크시트의 값·서식·열너비·병합셀을 다른 워크북의 시트로 복사."""
    from copy import copy
    for row in src_ws.iter_rows():
        for cell in row:
            nc = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                nc.font = copy(cell.font)
                nc.fill = copy(cell.fill)
                nc.border = copy(cell.border)
                nc.alignment = copy(cell.alignment)
                nc.protection = copy(cell.protection)
                nc.number_format = cell.number_format
    for col, dim in src_ws.column_dimensions.items():
        d = dst_ws.column_dimensions[col]
        if dim.width is not None:
            d.width = dim.width
        d.hidden = dim.hidden
    for idx, dim in src_ws.row_dimensions.items():
        d = dst_ws.row_dimensions[idx]
        if dim.height is not None:
            d.height = dim.height
        d.hidden = dim.hidden
    for mc in list(src_ws.merged_cells.ranges):
        dst_ws.merge_cells(str(mc))
    if src_ws.sheet_state:
        dst_ws.sheet_state = src_ws.sheet_state


def excel_merge_sheets(folder, files, order, out_name):
    """선택한 엑셀 파일들의 모든 시트를 하나의 새 워크북으로 모은다.
    order: asc/desc(시트 이름 기준) 또는 file(파일·시트 원래 순서)."""
    openpyxl = ensure_pkg("openpyxl")
    items = _select_paths(folder, files, ["xlsx", "xlsm"], "asc")
    if not items:
        raise RuntimeError("합칠 엑셀 파일(.xlsx/.xlsm)을 선택하세요.")

    collected, opened = [], []
    try:
        for name, path in items:
            try:
                wb = openpyxl.load_workbook(path)
            except Exception as e:
                raise RuntimeError(f"{name} 열기 실패: {e}")
            opened.append(wb)
            for ws in wb.worksheets:
                collected.append({"file": name, "sheet": ws.title, "ws": ws})
        if not collected:
            raise RuntimeError("시트를 찾을 수 없습니다.")
        if order == "asc":
            collected.sort(key=lambda c: _natkey(c["sheet"]))
        elif order == "desc":
            collected.sort(key=lambda c: _natkey(c["sheet"]), reverse=True)
        # order == "file": 정렬하지 않고 파일(asc)·시트 원래 순서를 유지

        out_wb = openpyxl.Workbook()
        out_wb.remove(out_wb.active)
        used = set()
        for c in collected:
            dst = out_wb.create_sheet(title=_sanitize_sheet_title(c["sheet"], used))
            _copy_sheet_content(c["ws"], dst)

        out_name = (out_name or "시트합치기").strip()
        if not out_name.lower().endswith((".xlsx", ".xlsm")):
            out_name += ".xlsx"
        out_path = _unique_path(folder, out_name)
        out_wb.save(out_path)
        return {"output": os.path.basename(out_path),
                "sheets": len(collected), "files": len(items)}
    finally:
        for wb in opened:
            try:
                wb.close()
            except Exception:
                pass


def excel_consolidate_rows(folder, files, out_name, skip_header=True):
    """첫 번째로 선택한 파일의 헤더를 기준으로, 선택한 모든 파일의 모든 시트의
    행을 위에서 아래로 이어 붙여 하나의 시트로 통합한다.
    files 순서 = 사용자가 선택한 순서(첫 항목이 헤더 기준)."""
    openpyxl = ensure_pkg("openpyxl")
    if not files:
        raise RuntimeError("통합할 엑셀 파일(.xlsx/.xlsm)을 선택하세요.")
    items = []
    for n in files:  # 선택 순서를 그대로 유지 (정렬하지 않음)
        if os.sep in n or n in (".", ".."):
            raise ValueError(f"잘못된 파일명: {n}")
        p = safe_join(folder, n)
        if not os.path.isfile(p):
            raise ValueError(f"파일을 찾을 수 없습니다: {n}")
        if os.path.splitext(n)[1].lower().lstrip(".") not in ("xlsx", "xlsm"):
            raise ValueError(f"지원하지 않는 형식입니다(.xlsx/.xlsm만): {n}")
        items.append((n, p))

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "통합"
    total_rows, sheets_used, header_written = 0, 0, False
    try:
        for fi, (name, path) in enumerate(items):
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                for si, ws in enumerate(wb.worksheets):
                    sheets_used += 1
                    first = True
                    for row in ws.iter_rows(values_only=True):
                        if first:
                            first = False
                            if skip_header:
                                # 첫 번째 파일·첫 번째 시트의 헤더만 한 번 기록
                                if fi == 0 and si == 0 and not header_written:
                                    out_ws.append(list(row))
                                    header_written = True
                                continue
                        if row is None or all(v is None for v in row):
                            continue
                        out_ws.append(list(row))
                        total_rows += 1
            finally:
                wb.close()

        out_name = (out_name or "행통합").strip()
        if not out_name.lower().endswith((".xlsx", ".xlsm")):
            out_name += ".xlsx"
        out_path = _unique_path(folder, out_name)
        out_wb.save(out_path)
        return {"output": os.path.basename(out_path), "rows": total_rows,
                "files": len(items), "sheets": sheets_used, "header": header_written}
    finally:
        try:
            out_wb.close()
        except Exception:
            pass


_TK_ROOT = None                    # 숨은 Tk 루트 — 매번 만들지 않고 재사용한다
_DIALOG_REQ = queue.Queue()        # 워커 스레드 → 메인 스레드 다이얼로그 요청 통로
_DIALOG_READY = threading.Event()  # 메인 스레드 디스패처가 돌고 있는지


def _show_folder_dialog():
    """tkinter 폴더 선택 다이얼로그를 띄우고 선택 경로(문자열)를 돌려준다.
    반드시 자신을 소유한 스레드(메인 스레드 권장)에서 호출할 것."""
    global _TK_ROOT
    import tkinter as tk
    from tkinter import filedialog
    if _TK_ROOT is None:
        _TK_ROOT = tk.Tk()
        _TK_ROOT.withdraw()
    _TK_ROOT.attributes('-topmost', True)
    return filedialog.askdirectory(title='대상 폴더를 선택하세요') or ''


# ---------- 여러 폴더를 한 번에 고르기 (Windows IFileOpenDialog) ----------
# tkinter 의 폴더 선택창은 한 번에 하나만 고를 수 있다. 윈도우 기본 파일 대화상자를
# '폴더 선택 + 다중 선택' 모드로 열면 탐색기처럼 드래그·Ctrl 클릭으로 여러 폴더를
# 한꺼번에 고를 수 있어, COM 인터페이스를 직접 호출한다.
_FOS_PICKFOLDERS      = 0x00000020
_FOS_FORCEFILESYSTEM  = 0x00000040
_FOS_ALLOWMULTISELECT = 0x00000200
_SIGDN_FILESYSPATH    = 0x80058000


def _com_call(ptr, index, restype, argtypes, *args):
    """COM 객체 vtable 의 index 번째 메서드를 호출한다."""
    import ctypes
    vtbl = ctypes.cast(ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
    fn = ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)(vtbl[index])
    return fn(ptr, *args)


def _show_multi_folder_dialog():
    """폴더를 여러 개 고르는 대화상자. 취소하면 빈 목록을 돌려준다.

    윈도우가 아니거나 대화상자를 못 여는 환경에서는 기존 단일 선택창으로 물러난다."""
    if os.name != "nt":
        one = _show_folder_dialog()
        return [one] if one else []

    import ctypes
    from ctypes import POINTER, byref, c_void_p, c_uint, c_ulong, c_ushort, c_ubyte, c_wchar_p

    class GUID(ctypes.Structure):
        _fields_ = [("Data1", c_ulong), ("Data2", c_ushort),
                    ("Data3", c_ushort), ("Data4", c_ubyte * 8)]

    ole32 = ctypes.OleDLL("ole32")
    HR = ctypes.HRESULT

    def guid(text):
        g = GUID()
        ole32.CLSIDFromString(text, byref(g))
        return g

    try:
        ole32.CoInitialize(None)
    except OSError:
        pass                      # 이미 초기화된 스레드면 그대로 쓴다

    dlg = c_void_p()
    ole32.CoCreateInstance(byref(guid("{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}")),  # FileOpenDialog
                           None, 1,
                           byref(guid("{D57C7288-D4AD-4768-BE02-9D969532D960}")),  # IFileOpenDialog
                           byref(dlg))
    try:
        opts = c_uint()
        _com_call(dlg, 10, HR, [POINTER(c_uint)], byref(opts))            # GetOptions
        _com_call(dlg, 9, HR, [c_uint],                                   # SetOptions
                  opts.value | _FOS_PICKFOLDERS | _FOS_FORCEFILESYSTEM | _FOS_ALLOWMULTISELECT)
        _com_call(dlg, 17, HR, [c_wchar_p],                               # SetTitle
                  "폴더를 여러 개 선택하세요 — 드래그하거나 Ctrl 을 누른 채 클릭")
        _com_call(dlg, 18, HR, [c_wchar_p], "이 폴더들 선택")              # SetOkButtonLabel
        try:
            _com_call(dlg, 3, HR, [c_void_p], None)                       # Show
        except OSError as e:
            if (e.winerror or 0) & 0xFFFF == 1223:                        # 사용자가 취소함
                return []
            raise
        arr = c_void_p()
        _com_call(dlg, 27, HR, [POINTER(c_void_p)], byref(arr))           # GetResults
        try:
            count = c_uint()
            _com_call(arr, 7, HR, [POINTER(c_uint)], byref(count))        # GetCount
            picked = []
            for i in range(count.value):
                item = c_void_p()
                _com_call(arr, 8, HR, [c_uint, POINTER(c_void_p)], i, byref(item))   # GetItemAt
                try:
                    name = c_wchar_p()
                    _com_call(item, 5, HR, [c_ulong, POINTER(c_wchar_p)],             # GetDisplayName
                              _SIGDN_FILESYSPATH, byref(name))
                    if name.value:
                        picked.append(name.value)
                        ole32.CoTaskMemFree(name)
                finally:
                    _com_call(item, 2, c_ulong, [])                        # Release
            return picked
        finally:
            _com_call(arr, 2, c_ulong, [])
    finally:
        _com_call(dlg, 2, c_ulong, [])


def _dialog_dispatcher():
    """메인 스레드에서 돌면서 폴더 선택 요청만 처리한다.

    Tk 는 자신을 만든 스레드에서만 다뤄야 해서 HTTP 워커 스레드가 직접 띄울 수
    없다. 예전에는 exe 자신을 다시 실행해 다이얼로그를 띄웠는데, onefile exe 는
    실행할 때마다 50MB를 임시폴더에 다시 푸느라 5초 넘게 걸렸다. 지금은 이미
    떠 있는 프로세스의 메인 스레드가 곧바로 띄운다."""
    _DIALOG_READY.set()
    while not STOP.is_set():
        try:
            kind, resp = _DIALOG_REQ.get(timeout=0.5)
        except queue.Empty:
            continue
        try:
            resp.put(_show_multi_folder_dialog() if kind == "multi" else _show_folder_dialog())
        except Exception:      # 다이얼로그 실패가 서버를 멈추게 하지는 않는다
            resp.put([] if kind == "multi" else "")


def _pick_folder_via_self():
    """frozen exe 전용: exe 자신을 특수 모드로 재실행하여 자식 프로세스의
    '메인 스레드'에서 다이얼로그를 띄우고, 결과를 임시 파일로 받는다.
    (HTTP 워커 스레드에서 직접 Tk를 돌리지 않으므로 안전)"""
    import tempfile
    fd, tmp = tempfile.mkstemp(suffix=".ffpick"); os.close(fd)
    try:
        env = dict(os.environ)
        env["FILEFORGE_PICK_OUT"] = tmp
        subprocess.run([sys.executable], env=env, timeout=300)
        with open(tmp, "r", encoding="utf-8") as f:
            picked = f.read()
    except (OSError, subprocess.TimeoutExpired):
        return None
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass
    picked = picked.strip().rstrip("/")
    return picked.replace("/", os.sep) if picked else None


def _ask_dialog(kind, on_timeout):
    """메인 스레드 디스패처에 대화상자를 요청하고 결과를 기다린다."""
    resp = queue.Queue(1)
    _DIALOG_REQ.put((kind, resp))
    try:
        return resp.get(timeout=600)
    except queue.Empty:
        return on_timeout


def pick_folders_native():
    """폴더 여러 개를 한 번에 고른다. 정리된 절대경로 목록을 돌려준다."""
    if not _DIALOG_READY.is_set():
        one = pick_folder_native()
        return [one] if one else []
    picked = _ask_dialog("multi", []) or []
    out, seen = [], set()
    for p in picked:
        p = (p or "").strip().rstrip("/").replace("/", os.sep)
        if not p or not os.path.isdir(p):
            continue
        key = os.path.normcase(os.path.abspath(p))
        if key in seen:
            continue
        seen.add(key)
        out.append(os.path.abspath(p))
    return out


def map_scan(folders, rules):
    """선택한 폴더들에서 매핑 규칙에 걸리는 파일을 찾아 바뀔 이름을 계산한다.

    화면(자바스크립트)의 '키워드별 지정값으로 변경'과 같은 규칙을 쓴다.
    확장자를 뺀 이름이 왼쪽 값과 (앞뒤 공백·대소문자 무시) 같으면 오른쪽 값으로 바꾼다.
    위쪽 줄이 우선한다."""
    pairs = []
    for r in rules or []:
        a = (r.get("a") or "").strip()
        b = (r.get("b") or "").strip()
        if a and b:
            pairs.append((a.lower(), b))

    out = []
    for idx, folder in enumerate(folders or []):
        info = {"folder": folder, "missing": False, "total": 0, "jobs": []}
        if not os.path.isdir(folder):
            info["missing"] = True
            out.append(info)
            continue
        if pairs:
            try:
                with os.scandir(folder) as it:
                    for entry in it:
                        if entry.name.startswith("."):
                            continue
                        try:
                            if not entry.is_file():
                                continue
                        except OSError:
                            continue
                        info["total"] += 1
                        stem, ext = os.path.splitext(entry.name)
                        key = stem.strip().lower()
                        for a, b in pairs:
                            if a == key:
                                dst = b + ext
                                if dst != entry.name:
                                    info["jobs"].append({"src": entry.name, "dst": dst, "d": idx})
                                break
            except OSError as e:
                info["missing"] = True
                info["error"] = str(e)
        out.append(info)
    return out


def pick_folder_native():
    """폴더 선택 다이얼로그 (HTTP 워커 스레드에서 호출된다).

    메인 스레드 디스패처가 돌고 있으면 같은 프로세스 안에서 곧바로 띄운다.
    디스패처가 없을 때만 예전 방식(별도 프로세스)으로 되돌아간다.
    """
    if _DIALOG_READY.is_set():
        picked = _ask_dialog("one", "")
        picked = (picked or "").strip().rstrip("/")
        return picked.replace("/", os.sep) if picked else None

    if getattr(sys, "frozen", False):
        return _pick_folder_via_self()
    code = (
        "import tkinter as tk\n"
        "from tkinter import filedialog\n"
        "r = tk.Tk(); r.withdraw(); r.attributes('-topmost', True)\n"
        "print(filedialog.askdirectory(title='대상 폴더를 선택하세요') or '')\n"
    )
    try:
        out = subprocess.run([sys.executable, "-c", code],
                             capture_output=True, text=True, timeout=300)
        path = out.stdout.strip().rstrip("/")
        return path.replace("/", os.sep) if out.returncode == 0 and path else None
    except (OSError, subprocess.TimeoutExpired):
        return None


# ---------------- HTML (Nike 디자인 시스템) ----------------

PAGE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>FILEFORGE</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
/* ===== DESIGN-nike.md tokens ===== */
:root{
  --ink:#111111; --canvas:#ffffff; --cloud:#f5f5f5;
  --charcoal:#39393b; --mute:#707072; --stone:#9e9ea0;
  --hairline:#cacacb; --hairline-soft:#e5e5e5;
  --sale:#d30005; --success:#007d48;
  --font-display:"Futura","Avenir Next Condensed","Bebas Neue",Impact,"Arial Narrow",sans-serif;
  --font-ui:"Helvetica Neue",Helvetica,"Segoe UI","Malgun Gothic",Arial,sans-serif;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--font-ui);color:var(--ink);background:var(--canvas);
     font-size:16px;line-height:1.5}
button{font-family:var(--font-ui);cursor:pointer;border:none;background:none;color:inherit}
input{font-family:var(--font-ui);font-size:16px;color:var(--ink)}

/* ===== utility bar (36px, soft-cloud) ===== */
.utility{background:var(--cloud);height:36px;display:flex;align-items:center;
  justify-content:flex-end;gap:24px;padding:0 48px;font-size:12px;font-weight:500}

/* ===== primary nav (56px, canvas) ===== */
nav{background:var(--canvas);min-height:56px;display:flex;align-items:center;
  padding:4px 48px;gap:24px;position:sticky;top:0;z-index:20;flex-wrap:wrap;
  box-shadow:inset 0 -1px 0 var(--hairline-soft)}
.wordmark{font-family:var(--font-display);font-size:22px;font-weight:700;
  text-transform:uppercase;letter-spacing:.5px;white-space:nowrap}
.nav-tabs{display:flex;gap:24px;flex:1;justify-content:center}
.nav-tab{font-size:16px;font-weight:500;padding:15px 2px;border-bottom:2px solid transparent;
  white-space:nowrap}
.nav-tab.active{border-bottom-color:var(--ink)}
.search-pill{background:var(--cloud);border:2px solid transparent;border-radius:24px;
  height:40px;padding:8px 16px;width:170px;min-width:120px;flex-shrink:1;
  outline:none;transition:all .12s}
.search-pill::placeholder{color:var(--mute)}
.search-pill:focus{background:var(--canvas);border-color:var(--ink);
  box-shadow:0 0 0 6px var(--cloud)}

/* ===== buttons: pill vocabulary only ===== */
.btn{border-radius:9999px;font-weight:500;font-size:16px;
  transition:transform .12s,opacity .12s;display:inline-flex;align-items:center;
  justify-content:center;white-space:nowrap}
.btn:active{transform:scale(.9);opacity:.6}  /* Nike tap-collapse */
.btn-primary{background:var(--ink);color:var(--canvas);padding:0 32px;height:48px}
.btn-secondary{background:var(--cloud);color:var(--ink);padding:0 32px;height:48px}
.btn-on-image{background:var(--canvas);color:var(--ink);padding:0 24px;height:44px}
.btn-sm{height:40px;padding:0 20px;font-size:14px}
.btn:disabled{opacity:.35;cursor:default}
.btn:disabled:active{transform:none}

/* ===== campaign hero (ink tile, Futura burn-in) ===== */
.hero{background:var(--ink);color:var(--canvas);padding:64px 48px;
  transition:padding .25s}
.hero h1{font-family:var(--font-display);font-weight:500;text-transform:uppercase;
  font-size:clamp(48px,8vw,96px);line-height:.9}
.hero .sub{color:var(--stone);font-size:16px;margin-top:18px;max-width:560px}
.hero .cta-row{margin-top:30px;display:flex;gap:12px;align-items:center}
.hero .path{color:var(--stone);font-size:14px;font-weight:500;word-break:break-all}
body.loaded .hero{padding:30px 48px}
body.loaded .hero h1{font-size:clamp(28px,3.4vw,44px)}
body.loaded .hero .sub{display:none}
body.loaded .hero .cta-row{margin-top:12px}

/* ===== content ===== */
main{max-width:1440px;margin:0 auto;padding:0 48px 120px}
section{padding-top:48px}
h2{font-size:32px;font-weight:500;line-height:1.2;text-transform:uppercase}
.section-head{display:flex;align-items:baseline;justify-content:space-between;
  gap:24px;flex-wrap:wrap;border-bottom:1px solid var(--hairline);padding-bottom:12px}
.section-head .count{text-align:right}
.section-head .count{font-size:14px;font-weight:500;color:var(--mute)}

/* ===== rule chips (filter-chip / filter-chip-active) ===== */
.chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:18px}
.chip{border:1px solid var(--hairline);border-radius:9999px;padding:8px 16px;
  font-size:16px;font-weight:500;background:var(--canvas);color:var(--ink);
  transition:transform .12s}
.chip:active{transform:scale(.9)}
.chip.active{background:var(--ink);color:var(--canvas);border-color:var(--ink)}

/* ===== rule panels: flat, hairline-divided ===== */
.panel{display:none;border-bottom:1px solid var(--hairline-soft);padding:24px 0}
.panel.show{display:block}
.panel h3{font-size:16px;font-weight:500;line-height:1.75;margin-bottom:8px}
.row{display:flex;gap:12px;flex-wrap:wrap;align-items:center}
#panel-rebuild .row + .row{margin-top:12px}
.field{background:var(--cloud);border:2px solid transparent;border-radius:24px;
  height:40px;padding:8px 16px;outline:none;width:220px;transition:all .12s}
.field:focus{background:var(--canvas);border-color:var(--ink);box-shadow:0 0 0 6px var(--cloud)}
.field.narrow{width:90px}
label.opt{display:inline-flex;align-items:center;gap:8px;font-size:14px;font-weight:500}
select.field{width:auto;cursor:pointer}
.hint{font-size:12px;font-weight:500;color:var(--mute);margin-top:8px}

/* ===== preview table: flat rows, hairline dividers ===== */
.thead,.frow{display:grid;grid-template-columns:48px 1fr 1fr 130px;gap:12px;
  align-items:center;padding:12px 0}
.thead{font-size:12px;font-weight:500;color:var(--mute);text-transform:uppercase;
  border-bottom:1px solid var(--hairline)}
.frow{border-bottom:1px solid var(--hairline-soft);cursor:pointer}
.frow .old{font-weight:500}
.frow .new{color:var(--charcoal)}
.frow .st{font-size:14px;font-weight:500}
.frow.changed .st{color:var(--success)}
.frow.conflict .st,.frow.conflict .new{color:var(--sale)}
.frow.excluded{color:var(--stone)}
.frow.excluded .old,.frow.excluded .new{color:var(--stone)}
.empty{padding:48px 0;color:var(--mute);font-size:16px}

/* per-file text column (Excel식 채우기) — '파일별 텍스트' 규칙이 켜질 때만 표시 */
body.textcol .thead, body.textcol .frow{grid-template-columns:48px 1fr 190px 1fr 130px}
.col-text{display:none}
body.textcol .col-text{display:block}
.col-text .cell{position:relative;display:block}
.tcell{width:100%;background:var(--cloud);border:1px solid transparent;border-radius:8px;
  height:34px;padding:6px 12px;font-size:14px;color:var(--ink);outline:none}
.tcell:focus{background:var(--canvas);border-color:var(--ink)}
.fillh{position:absolute;right:-3px;bottom:-3px;width:11px;height:11px;background:var(--ink);
  border:2px solid var(--canvas);border-radius:2px;cursor:crosshair;z-index:2}
.frow.fill-target{background:var(--cloud)}
.frow.fill-target .tcell{background:var(--canvas)}

/* swatch-dot include toggle (concentric ring = included) */
.dot{width:12px;height:12px;border-radius:50%;background:var(--ink);
  box-shadow:0 0 0 2px var(--canvas),0 0 0 4px var(--ink)}
.frow.excluded .dot{background:var(--canvas);border:1px solid var(--hairline);box-shadow:none}

/* 첫 칸: 순서 변경 손잡이 + 포함 점 */
.firstcell{display:flex;align-items:center;gap:6px;margin-left:6px}
.firstcell .dot{margin-left:0}
.grip{color:var(--stone);font-size:13px;line-height:1;cursor:grab;user-select:none;
  padding:0 1px;flex:none;letter-spacing:-2px}
.grip:active{cursor:grabbing}
.frow.dragging{opacity:.45}

/* 미리보기 도구 (전체 선택/해제) */
.prev-tools{display:flex;gap:8px;align-items:center}
.prev-hint{font-size:12px;font-weight:500;color:var(--mute);margin-top:8px}

/* ===== 키워드별 정리: 동적 입력 필드 ===== */
.kwrow{display:flex;gap:8px;align-items:center;margin-bottom:8px}
.kwrow .kwnum{width:22px;color:var(--mute);font-size:14px;font-weight:500;text-align:right}
.kwrow .field{flex:1;max-width:380px}
.kwdel{width:34px;height:34px;border-radius:9999px;background:var(--cloud);color:var(--ink);
  font-size:16px;flex:none;transition:transform .12s}
.kwdel:active{transform:scale(.9)}
#kwAdd{cursor:ns-resize}
/* 키워드별 지정값으로 변경: A → B 매핑 줄 */
.maprow{display:flex;gap:8px;align-items:center;margin-bottom:8px}
.maprow .mapnum{width:22px;color:var(--mute);font-size:14px;font-weight:500;text-align:right}
.maprow .field{flex:1;max-width:240px}
.maprow .maparrow{color:var(--mute);font-weight:500;flex:none}
.mapdel{width:34px;height:34px;border-radius:9999px;background:var(--cloud);color:var(--ink);
  font-size:16px;flex:none;transition:transform .12s}
.mapdel:active{transform:scale(.9)}

/* ===== PDF · 변환 도구: 선택 가능한 파일 목록 ===== */
.pdf-listhead{display:flex;justify-content:space-between;align-items:center;
  border-bottom:1px solid var(--hairline);padding:12px 0;margin-top:18px}
.pdfrow{display:flex;align-items:center;gap:14px;padding:12px 0;
  border-bottom:1px solid var(--hairline-soft);cursor:pointer}
.pdfdot{width:18px;height:18px;border-radius:50%;border:2px solid var(--hairline);
  flex:none;transition:all .12s}
.pdfrow.sel .pdfdot{border-color:var(--ink);background:var(--ink);
  box-shadow:inset 0 0 0 3px var(--canvas)}
.pdfrow .pname{font-weight:500}
.pdfrow.sel .pname{color:var(--ink)}
.pdfrange{flex:none;width:200px;margin-left:auto;padding:8px 14px;font-size:13px}
.pdf-actions{margin-top:24px;display:flex;gap:12px;align-items:center;flex-wrap:wrap}
#pdfResult{font-size:14px;font-weight:500}
#pdfResult.ok{color:var(--success)}
#pdfResult.warn{color:var(--sale)}
#pdfOptions{margin-top:18px}
#pdfOptions .row{margin-bottom:6px}

/* ===== sticky action bar ===== */
.actionbar{position:fixed;left:0;right:0;bottom:0;background:var(--canvas);
  box-shadow:inset 0 1px 0 var(--hairline-soft);padding:12px 48px;
  display:flex;align-items:center;gap:12px;z-index:30}
.actionbar .status{flex:1;font-size:14px;font-weight:500;color:var(--mute)}
.actionbar .status b{color:var(--ink)}
.actionbar .status .ok{color:var(--success)}
.actionbar .status .warn{color:var(--sale)}

/* ===== confirm dialog (geo-selector pattern) ===== */
.scrim{position:fixed;inset:0;background:rgba(17,17,17,.5);display:none;
  align-items:center;justify-content:center;z-index:50}
.scrim.show{display:flex}
.dialog{background:var(--canvas);padding:48px;max-width:480px;width:90%}
.dialog h3{font-size:24px;font-weight:500;line-height:1.2;margin-bottom:12px}
.dialog p{color:var(--mute);margin-bottom:30px}
.dialog .row{justify-content:flex-end}

footer{border-top:1px solid var(--hairline);margin-top:48px;padding:24px 48px;
  font-size:9px;font-weight:500;line-height:1.75;color:var(--mute)}

@media (max-width:960px){
  nav,.hero,main,.actionbar,footer{padding-left:24px;padding-right:24px}
  .utility{padding:0 24px}
  .thead,.frow{grid-template-columns:40px 1fr 1fr 90px}
  body.textcol .thead, body.textcol .frow{grid-template-columns:40px 1fr 130px 1fr 90px}
  .nav-tabs{justify-content:flex-start;order:3;flex-basis:100%}
  .nav-tab{padding:8px 2px 10px}
}
/* ===== 폴더 정리: 폴더 트리 미리보기 ===== */
#orgTreeWrap{margin-top:26px}
.tree{font-family:var(--font-ui);font-size:14px;line-height:1.45;
  border:1px solid var(--hairline-soft);border-radius:10px;background:var(--cloud);
  padding:14px 18px;max-height:440px;overflow:auto}
.tree .empty{border:none;background:none;padding:6px 2px;color:var(--mute)}
.tree .troot{display:flex;align-items:center;gap:8px;font-weight:700;
  padding-bottom:8px;margin-bottom:6px;border-bottom:1px solid var(--hairline-soft)}
.tree .tchildren{margin-left:9px;padding-left:16px;border-left:1px solid var(--hairline)}
.tree .tfolder{display:flex;align-items:center;gap:8px;padding:4px 0;font-weight:600}
.tree .tname{color:var(--ink)}
.tree .tcount{color:var(--mute);font-weight:500;font-size:12px}
.tree .tfile{display:flex;align-items:center;gap:8px;padding:2px 0;color:var(--charcoal)}
.tree .tmore{color:var(--mute);font-size:12px;padding:2px 0 2px 2px}
.tree .ticon{flex:none;width:18px;text-align:center}
.tree .tbadge{font-size:11px;font-weight:700;border-radius:4px;padding:0 6px;line-height:17px}
.tree .tnew{color:var(--success);border:1px solid var(--success)}
.tree .texist{color:var(--mute);border:1px solid var(--hairline);font-weight:600}
.tree .tbadge{margin-left:auto}
.tree .tfiles{margin-left:24px;border-left:1px solid var(--hairline-soft);
  padding-left:14px;margin-top:2px;margin-bottom:8px}
.tree .tsub{display:flex;align-items:center;gap:8px;padding:2px 0;color:var(--mute);font-size:13px}

/* ===== 폴더 만들기 탭 ===== */
.mkpanel{display:none;padding:24px 0;border-bottom:1px solid var(--hairline-soft)}
.mkpanel.show{display:block}
.mkpanel h3{font-size:16px;font-weight:500;line-height:1.75;margin-bottom:8px}
.area{background:var(--cloud);border:2px solid transparent;border-radius:16px;
  padding:12px 16px;outline:none;width:100%;max-width:520px;min-height:132px;resize:vertical;
  font-family:var(--font-ui);font-size:15px;line-height:1.6;color:var(--ink);transition:all .12s}
.area:focus{background:var(--canvas);border-color:var(--ink);box-shadow:0 0 0 6px var(--cloud)}
.area.half{max-width:320px}
.mk-two{display:flex;gap:18px;flex-wrap:wrap}
.mk-two .col{display:flex;flex-direction:column;gap:8px}
.mk-two .col .lbl{font-size:12px;font-weight:500;color:var(--mute);text-transform:uppercase}
#mkCommon{padding:24px 0;border-bottom:1px solid var(--hairline-soft)}
#mkTreeWrap{margin-top:26px}
.mk-actions{margin-top:18px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}

/* 폴더 아이콘 (이름 변경 미리보기에서 폴더 구분) */
.ficon{margin-right:6px}

/* 자동저장 표시 · 최근 작업 이력 */
.utility .saved{color:var(--mute)}
.utility button{font-size:12px;font-weight:500;text-decoration:underline;text-underline-offset:3px}
#histBox{margin-top:48px;border-top:1px solid var(--hairline);padding-top:18px}
#histBox summary{font-size:14px;font-weight:500;cursor:pointer;list-style:none}
#histBox summary::-webkit-details-marker{display:none}
#histBox summary::before{content:"▸ ";color:var(--mute)}
#histBox[open] summary::before{content:"▾ "}
.hrow{display:grid;grid-template-columns:150px 110px 1fr auto;gap:12px;padding:8px 0;
  border-bottom:1px solid var(--hairline-soft);font-size:13px}
.hrow .htime{color:var(--mute);font-weight:500}
.hrow .hkind{font-weight:500}
.hrow .gobtn{align-self:center}
@media (max-width:960px){          /* 좁은 화면: 버튼을 아랫줄로 내린다 */
  .hrow{grid-template-columns:110px 90px 1fr}
  .hrow .gobtn{grid-column:1/-1;justify-self:start;margin:4px 0 0}
}

/* ===== 작업한 폴더로 이동 ===== */
.gobtn{border:1px solid var(--hairline);background:var(--canvas);color:var(--ink);
  border-radius:999px;height:26px;padding:0 12px;font-size:12px;font-weight:500;
  cursor:pointer;white-space:nowrap;vertical-align:middle;margin-left:4px}
.gobtn:hover{background:var(--cloud)}
.recent-wrap{position:relative}
.recent-menu{display:none;position:absolute;right:0;top:calc(100% + 6px);z-index:60;
  min-width:min(300px,calc(100vw - 48px));max-width:min(560px,calc(100vw - 48px));
  background:var(--canvas);
  border:1px solid var(--hairline);border-radius:12px;padding:6px;
  box-shadow:0 12px 32px rgba(0,0,0,.14)}
.recent-menu.show{display:block}
.recent-item{display:block;width:100%;text-align:left;background:none;border:0;cursor:pointer;
  padding:9px 12px;border-radius:8px;font-size:13px;color:var(--ink);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.recent-item:hover{background:var(--cloud)}
.recent-empty{padding:10px 12px;font-size:13px;color:var(--mute)}

/* ===== 키워드별 지정값으로 변경: 여러 폴더에 한꺼번에 적용 ===== */
.multibox{margin-top:22px;padding-top:18px;border-top:1px solid var(--hairline)}
.multihead{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:8px}
.multihead b{font-size:14px}
#mapFolderList{margin-top:12px}
.mfrow{display:grid;grid-template-columns:1fr auto auto;gap:12px;align-items:center;
  padding:8px 0;border-bottom:1px solid var(--hairline-soft);font-size:13px}
.mfpath{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mfcount{white-space:nowrap;font-weight:500}
.multiapply{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-top:14px}
@media (max-width:960px){
  .mfrow{grid-template-columns:1fr auto}
  .mfrow .gobtn{grid-column:1/-1;justify-self:start;margin:4px 0 0}
}
@media (max-width:960px){
  /* 좁은 화면에선 nav 가 줄바꿈돼 버튼이 가운데로 온다. 버튼 오른쪽 끝에 맞추면
     메뉴가 화면 밖으로 넘치므로, 화면 좌우 여백에 맞춰 편다. */
  .recent-menu{position:fixed;left:24px;right:24px;min-width:0;max-width:none}
}
</style>
</head>
<body>

<div class="utility"><span>FileForge — 파일 자동화 도구</span>
  <span class="saved" id="savedAt">자동저장 대기 중</span>
  <button onclick="clearDraft()" title="입력한 규칙·목록을 모두 비웁니다">입력 초기화</button></div>

<nav>
  <span class="wordmark">FileForge</span>
  <div class="nav-tabs">
    <button class="nav-tab active" data-tab="rename" onclick="switchTab('rename')">이름 변경</button>
    <button class="nav-tab" data-tab="organize" onclick="switchTab('organize')">폴더 정리</button>
    <button class="nav-tab" data-tab="make" onclick="switchTab('make')">폴더 만들기</button>
    <button class="nav-tab" data-tab="pdf" onclick="switchTab('pdf')">PDF · 변환</button>
  </div>
  <input id="extFilter" class="search-pill" placeholder="확장자 필터 (jpg,png)"
         oninput="debounceReload()">
  <button class="btn btn-primary btn-sm" onclick="pickFolder()">폴더 선택</button>
  <div class="recent-wrap">
    <button class="btn btn-secondary btn-sm" onclick="toggleRecent(event)"
            title="최근에 작업한 폴더를 탐색기에서 엽니다">최근 폴더 ▾</button>
    <div class="recent-menu" id="recentMenu"></div>
  </div>
</nav>

<div class="hero">
  <h1 id="heroTitle">Forge<br>Every File.</h1>
  <div class="sub">규칙을 고르면 결과가 즉시 미리보기로 나타납니다. 실행 전에는 어떤 파일도 바뀌지 않고, 실행 후에도 한 번에 되돌릴 수 있습니다.</div>
  <div class="cta-row">
    <button class="btn btn-on-image" onclick="pickFolder()">폴더 선택하기</button>
    <span class="path" id="heroPath"></span>
  </div>
</div>

<main>
  <!-- ===== 이름 변경 탭 ===== -->
  <section id="tab-rename">
    <div class="section-head"><h2>Rules</h2><span class="count">칩을 눌러 규칙을 켜고 끄세요 — 위에서 아래 순서로 적용됩니다</span></div>
    <div class="row" id="targetRow" style="margin-top:18px">
      <span class="hint">변경 대상</span>
      <select class="field" id="renameTarget" onchange="onTargetChange()">
        <option value="files">파일만</option>
        <option value="dirs">폴더만</option>
        <option value="both">파일 + 폴더</option>
      </select>
      <span class="hint">아래 <b>모든 규칙</b>이 폴더 이름에도 똑같이 적용됩니다. 폴더는 확장자가 없어서 이름 전체가 바뀌고, 미리보기에서 📁 로 표시됩니다.</span>
    </div>
    <div class="chips">
      <button class="chip" data-rule="rebuild" onclick="toggleRule(this)">이름 새로 입력</button>
      <button class="chip" data-rule="mapname" onclick="toggleRule(this)">키워드별 지정값으로 변경</button>
      <button class="chip" data-rule="replace" onclick="toggleRule(this)">찾기 / 바꾸기</button>
      <button class="chip" data-rule="clean" onclick="toggleRule(this)">자동 정리</button>
      <button class="chip" data-rule="affix" onclick="toggleRule(this)">접두사 / 접미사</button>
      <button class="chip" data-rule="pertext" onclick="toggleRule(this)">파일별 텍스트</button>
      <button class="chip" data-rule="date" onclick="toggleRule(this)">날짜 붙이기</button>
      <button class="chip" data-rule="seq" onclick="toggleRule(this)">순번 매기기</button>
      <button class="chip" data-rule="case" onclick="toggleRule(this)">대소문자</button>
    </div>

    <div class="panel" id="panel-rebuild">
      <h3>이름 새로 입력 — 기존 이름을 지우고 통째로 새 값을 넣습니다</h3>
      <div class="row">
        <select class="field" id="rebuildMode" onchange="onRebuildModeChange()">
          <option value="number">숫자 순번 (001, 002, 003…)</option>
          <option value="common">공통 이름 (겹치면 _1, _2, _3…)</option>
          <option value="perfile">파일별 입력 (직접 / 드래그 / 붙여넣기)</option>
        </select>
        <span class="hint">이 규칙을 켜면 확장자만 남기고 이름이 새 값으로 바뀝니다. 다른 규칙(접두사·날짜 등)과 함께 쓰면 그 위에 덧붙여집니다.</span>
      </div>
      <div class="row" id="rebuildNumberOpts">
        <input class="field" id="rebuildBase" placeholder="앞에 붙일 이름 (선택)" oninput="recompute()">
        <span class="hint">시작 번호</span>
        <input class="field narrow" id="rebuildStart" type="number" value="1" min="0" oninput="recompute()">
        <span class="hint">자릿수</span>
        <input class="field narrow" id="rebuildPad" type="number" value="3" min="1" max="6" oninput="recompute()">
      </div>
      <div class="row" id="rebuildCommonOpts" style="display:none">
        <input class="field" id="rebuildCommon" placeholder="모든 파일에 넣을 이름 (예: 보고서)" oninput="recompute()">
        <span class="hint">파일이 여러 개면 겹치지 않도록 뒤에 <b>_1, _2, _3…</b> 이 자동으로 붙습니다.</span>
      </div>
      <div class="row" id="rebuildPerfileOpts" style="display:none">
        <span class="hint">아래 미리보기의 <b>‘붙일 텍스트’</b> 칸에 파일마다 새 이름을 입력하세요. 칸 오른쪽 아래 ■ 손잡이를 아래로 드래그하면 같은 값이 연속으로 채워지고, 복사·붙여넣기(엑셀에서 여러 줄 붙여넣기 포함)로 한꺼번에 채울 수 있습니다.</span>
      </div>
    </div>
    <div class="panel" id="panel-mapname">
      <h3>키워드별 지정값으로 변경 — 기존 이름이 왼쪽 값과 <b>정확히 같으면</b> 오른쪽 값으로 바꿉니다</h3>
      <p class="hint" style="margin-bottom:12px">확장자는 그대로 유지됩니다. 예) 왼쪽 <b>1</b> → 오른쪽 <b>1_이력서</b> 로 두면 <b>1.pdf</b> 가 <b>1_이력서.pdf</b> 로 바뀝니다.
        빈 칸·매칭 안 되는 파일은 그대로 둡니다. 입력칸에서 <b>Enter</b>를 누르면 같은 칸의 <b>아랫줄</b>로 내려가고
        (마지막 줄이면 새 줄이 생깁니다), <b>Ctrl+Enter</b>는 그 자리 아래에 줄을 하나 끼워 넣습니다.</p>
      <div id="mapFields"></div>
      <button class="btn btn-secondary btn-sm" onclick="addMapRow()">＋ 줄 추가</button>

      <div class="multibox">
        <div class="multihead">
          <b>여러 폴더에 한꺼번에 적용</b>
          <button class="btn btn-secondary btn-sm" onclick="pickManyFolders()">여러 폴더 선택</button>
          <button class="gobtn" id="mapClear" onclick="clearMapFolders()" style="display:none">선택 해제</button>
        </div>
        <p class="hint">폴더 선택창에서 <b>드래그</b>하거나 <b>Ctrl</b>을 누른 채 클릭하면 여러 폴더를 한 번에 고를 수 있습니다.
          고른 모든 폴더에서 위 규칙에 해당하는 파일 이름이 함께 바뀝니다.
          (위쪽 미리보기 표는 지금 선택된 폴더 하나만 보여줍니다)</p>
        <div id="mapFolderList"></div>
        <div class="multiapply" id="mapApplyRow" style="display:none">
          <button class="btn btn-primary btn-sm" id="mapApplyBtn" onclick="askApplyMap()">여러 폴더에 적용</button>
          <span class="hint" id="mapApplyInfo"></span>
        </div>
      </div>
    </div>
    <div class="panel" id="panel-replace">
      <h3>찾기 / 바꾸기</h3>
      <div class="row">
        <input class="field" id="findText" placeholder="찾을 텍스트" oninput="recompute()">
        <input class="field" id="replaceText" placeholder="바꿀 텍스트" oninput="recompute()">
        <label class="opt"><input type="checkbox" id="useRegex" onchange="recompute()"> 정규식</label>
      </div>
    </div>
    <div class="panel" id="panel-clean">
      <h3>자동 정리</h3>
      <div class="row">
        <label class="opt"><input type="checkbox" id="cleanSpaces" checked onchange="recompute()"> 연속 공백 정리</label>
        <label class="opt"><input type="checkbox" id="cleanSpecial" onchange="recompute()"> 특수문자 제거</label>
        <label class="opt"><input type="checkbox" id="cleanUnderscore" onchange="recompute()"> 공백 → 밑줄(_)</label>
      </div>
    </div>
    <div class="panel" id="panel-affix">
      <h3>접두사 / 접미사</h3>
      <div class="row">
        <input class="field" id="prefix" placeholder="이름 앞에 붙일 텍스트" oninput="recompute()">
        <input class="field" id="suffix" placeholder="이름 뒤에 붙일 텍스트" oninput="recompute()">
      </div>
    </div>
    <div class="panel" id="panel-pertext">
      <h3>파일별 텍스트</h3>
      <div class="row">
        <select class="field" id="perTextPos" onchange="recompute()">
          <option value="front">이름 앞에</option><option value="back">이름 뒤에</option>
        </select>
        <span class="hint">아래 미리보기의 <b>‘붙일 텍스트’</b> 칸에 파일마다 입력하세요. 칸 오른쪽 아래 ■ 손잡이를 아래로 드래그하면 같은 값이 연속으로 채워집니다. 복사·붙여넣기(여러 줄 붙여넣기 포함)도 됩니다.</span>
      </div>
    </div>
    <div class="panel" id="panel-date">
      <h3>날짜 붙이기</h3>
      <div class="row">
        <select class="field" id="dateSource" onchange="recompute()">
          <option value="today">오늘 날짜</option><option value="mtime">파일 수정일</option>
        </select>
        <select class="field" id="dateFormat" onchange="recompute()">
          <option>YYYYMMDD</option><option>YYYY-MM-DD</option><option>YYMMDD</option><option>YYYY.MM.DD</option>
        </select>
        <select class="field" id="datePos" onchange="recompute()">
          <option value="front">앞에</option><option value="back">뒤에</option>
        </select>
      </div>
    </div>
    <div class="panel" id="panel-seq">
      <h3>순번 매기기</h3>
      <div class="row">
        <input class="field narrow" id="seqStart" type="number" value="1" min="0" oninput="recompute()">
        <span class="hint">시작 번호</span>
        <input class="field narrow" id="seqPad" type="number" value="2" min="1" max="6" oninput="recompute()">
        <span class="hint">자릿수</span>
        <select class="field" id="seqPos" onchange="recompute()">
          <option value="front">앞에</option><option value="back">뒤에</option>
        </select>
      </div>
    </div>
    <div class="panel" id="panel-case">
      <h3>대소문자 (영문)</h3>
      <div class="row">
        <select class="field" id="caseMode" onchange="recompute()">
          <option value="lower">소문자로</option><option value="upper">대문자로</option>
          <option value="title">단어 첫 글자만 대문자</option>
        </select>
        <label class="opt"><input type="checkbox" id="extLower" checked onchange="recompute()"> 확장자는 항상 소문자</label>
      </div>
    </div>
  </section>

  <!-- ===== 폴더 정리 탭 ===== -->
  <section id="tab-organize" style="display:none">
    <div class="section-head"><h2>Organize</h2><span class="count">선택한 방식의 하위 폴더로 파일을 이동합니다</span></div>
    <div class="chips" id="orgChips">
      <button class="chip active" data-org="category" onclick="pickOrg(this)">카테고리별</button>
      <button class="chip" data-org="ext" onclick="pickOrg(this)">확장자별</button>
      <button class="chip" data-org="month" onclick="pickOrg(this)">날짜별 (년-월)</button>
      <button class="chip" data-org="day" onclick="pickOrg(this)">날짜별 (년-월-일)</button>
      <button class="chip" data-org="keyword" onclick="pickOrg(this)">키워드별</button>
    </div>

    <div id="kwBox" style="display:none;margin-top:18px">
      <p class="hint" style="margin-bottom:10px">입력한 <b>키워드가 파일명에 포함된 파일</b>끼리 그 키워드 이름의 폴더로 모읍니다.
        (여러 키워드에 해당하면 위쪽 키워드가 우선)</p>
      <div id="kwFields"></div>
      <button class="btn btn-secondary btn-sm" id="kwAdd"
              onmousedown="startKwDrag(event)" title="클릭하면 1개 추가 · 누른 채로 아래로 드래그하면 여러 개 추가">＋ 키워드 추가</button>
      <span class="hint" id="kwAddHint" style="margin-left:10px"></span>
    </div>

    <p class="hint" id="orgHint" style="margin-top:18px">카테고리별: 문서 / 엑셀·데이터 / 이미지 / 영상 / 압축 등으로 자동 분류 · 이동도 실행 취소로 원위치됩니다.</p>

    <div id="orgTreeWrap">
      <div class="section-head"><h2>폴더 구조 미리보기</h2><span class="count" id="orgTreeInfo"></span></div>
      <div id="orgTree" class="tree"><div class="empty">폴더를 선택하면 생성될 폴더 구조가 표시됩니다.</div></div>
    </div>
  </section>

  <!-- ===== 폴더 만들기 탭 ===== -->
  <section id="tab-make" style="display:none">
    <div class="section-head"><h2>Make Folders</h2><span class="count">선택한 폴더 안에 새 폴더를 한 번에 만듭니다 — 이미 있는 폴더는 건드리지 않습니다</span></div>
    <div class="chips" id="mkChips">
      <button class="chip active" data-mk="list" onclick="pickMk(this)">목록으로 만들기</button>
      <button class="chip" data-mk="number" onclick="pickMk(this)">번호로 만들기</button>
      <button class="chip" data-mk="date" onclick="pickMk(this)">날짜로 만들기</button>
      <button class="chip" data-mk="combo" onclick="pickMk(this)">조합으로 만들기 (A × B)</button>
      <button class="chip" data-mk="fromfiles" onclick="pickMk(this)">파일 이름으로 만들기</button>
    </div>

    <div class="mkpanel show" id="mk-list">
      <h3>목록으로 만들기 — 한 줄에 폴더 이름 하나씩</h3>
      <p class="hint" style="margin-bottom:10px">엑셀에서 이름이 적힌 열을 복사해 그대로 붙여넣어도 됩니다.
        <b>/</b> 를 쓰면 하위 폴더까지 한 번에 만들어집니다. (예: <b>2026/01월</b>)</p>
      <textarea class="area" id="mkList" placeholder="1_기획&#10;2_디자인&#10;3_개발" oninput="renderMake()"></textarea>
      <div class="mk-actions">
        <button class="btn btn-secondary btn-sm" onclick="mkSample()">예시로 채우기</button>
        <button class="btn btn-secondary btn-sm" onclick="mkDedup()">중복 줄 정리</button>
        <button class="btn btn-secondary btn-sm" onclick="mkFromClipboard()">붙여넣기</button>
        <button class="btn btn-secondary btn-sm" onclick="$('mkList').value='';renderMake()">비우기</button>
      </div>
    </div>

    <div class="mkpanel" id="mk-number">
      <h3>번호로 만들기 — 같은 이름 + 일련번호</h3>
      <div class="row">
        <input class="field" id="mkNumBase" placeholder="앞에 붙일 이름 (예: 사진_)" oninput="renderMake()">
        <input class="field" id="mkNumTail" placeholder="뒤에 붙일 이름 (선택)" oninput="renderMake()">
      </div>
      <div class="row" style="margin-top:12px">
        <span class="hint">시작 번호</span>
        <input class="field narrow" id="mkNumStart" type="number" value="1" min="0" oninput="renderMake()">
        <span class="hint">자릿수</span>
        <input class="field narrow" id="mkNumPad" type="number" value="2" min="1" max="6" oninput="renderMake()">
        <span class="hint">개수</span>
        <input class="field narrow" id="mkNumCount" type="number" value="10" min="1" max="500" oninput="renderMake()">
        <span class="hint">한 번에 최대 500개까지 만들 수 있습니다.</span>
      </div>
    </div>

    <div class="mkpanel" id="mk-date">
      <h3>날짜로 만들기 — 기간을 정하면 날짜 폴더가 쭉 만들어집니다</h3>
      <div class="row">
        <span class="hint">시작</span><input class="field" id="mkDateFrom" type="date" onchange="renderMake()">
        <span class="hint">종료</span><input class="field" id="mkDateTo" type="date" onchange="renderMake()">
        <select class="field" id="mkDateUnit" onchange="renderMake()">
          <option value="day">하루마다</option><option value="week">일주일마다</option>
          <option value="month" selected>한 달마다</option>
        </select>
      </div>
      <div class="row" style="margin-top:12px">
        <span class="hint">형식</span>
        <select class="field" id="mkDateFmt" onchange="renderMake()">
          <option value="YYYY-MM">2026-01</option>
          <option value="YYYY년 MM월">2026년 01월</option>
          <option value="YYYY-MM-DD">2026-01-05</option>
          <option value="YYYYMMDD">20260105</option>
          <option value="MM월">01월</option>
        </select>
        <input class="field" id="mkDatePre" placeholder="앞에 붙일 이름 (선택)" oninput="renderMake()">
      </div>
    </div>

    <div class="mkpanel" id="mk-combo">
      <h3>조합으로 만들기 — A 목록 × B 목록</h3>
      <p class="hint" style="margin-bottom:12px">예) A에 <b>영업팀·관리팀</b>, B에 <b>1월·2월</b> 을 넣으면 4개 조합이 만들어집니다.</p>
      <div class="mk-two">
        <div class="col"><span class="lbl">A 목록 (한 줄에 하나)</span>
          <textarea class="area half" id="mkComboA" placeholder="영업팀&#10;관리팀" oninput="renderMake()"></textarea></div>
        <div class="col"><span class="lbl">B 목록 (한 줄에 하나)</span>
          <textarea class="area half" id="mkComboB" placeholder="1월&#10;2월" oninput="renderMake()"></textarea></div>
      </div>
      <div class="row" style="margin-top:12px">
        <select class="field" id="mkComboMode" onchange="renderMake()">
          <option value="nested">A 폴더 안에 B 폴더 (영업팀/1월)</option>
          <option value="flat">한 폴더 이름으로 합치기 (영업팀_1월)</option>
        </select>
        <span class="hint">이어붙일 기호</span>
        <input class="field narrow" id="mkComboSep" value="_" oninput="renderMake()">
      </div>
    </div>

    <div class="mkpanel" id="mk-fromfiles">
      <h3>파일 이름으로 만들기 — 폴더 안 파일 이름을 그대로 폴더로</h3>
      <p class="hint" style="margin-bottom:12px">상단에서 폴더를 먼저 선택하세요. 파일이 많을 때 <b>같은 이름 규칙끼리 묶어</b> 폴더를 만들고, 원하면 파일을 그 안으로 옮겨줍니다.</p>
      <div class="row">
        <span class="hint">폴더 이름 기준</span>
        <select class="field" id="mkFfMode" onchange="renderMake()">
          <option value="whole">파일 이름 전체 (확장자 제외)</option>
          <option value="delim">기호 앞부분까지</option>
          <option value="firstN">앞에서 N글자</option>
        </select>
        <input class="field narrow" id="mkFfDelim" value="_" oninput="renderMake()" title="기호 앞부분까지 기준일 때 사용">
        <input class="field narrow" id="mkFfN" type="number" value="4" min="1" max="40" oninput="renderMake()" title="앞에서 N글자 기준일 때 사용">
      </div>
      <div class="row" style="margin-top:12px">
        <label class="opt"><input type="checkbox" id="mkFfMove" onchange="renderMake()"> 만든 폴더로 파일도 함께 옮기기</label>
        <span class="hint">옮긴 파일도 <b>실행 취소</b>로 한 번에 원위치됩니다.</span>
      </div>
    </div>

    <div id="mkCommon">
      <h3 style="font-size:16px;font-weight:500;margin-bottom:12px">공통 옵션</h3>
      <div class="row">
        <span class="hint">상위 폴더</span>
        <input class="field" id="mkParent" placeholder="예: 2026년_자료 (비우면 선택한 폴더 바로 아래)" oninput="renderMake()" style="width:320px">
      </div>
      <div class="row" style="margin-top:12px;align-items:flex-start">
        <span class="hint" style="margin-top:10px">하위 폴더 공통 생성</span>
        <input class="field" id="mkSubs" placeholder="예: 01_원본, 02_작업, 03_완료 (쉼표로 구분)" oninput="renderMake()" style="width:420px">
      </div>
      <p class="hint" style="margin-top:10px">만들어지는 모든 폴더 안에 같은 하위 폴더를 함께 만들어 둡니다. 필요 없으면 비워 두세요.</p>
    </div>

    <div id="mkTreeWrap">
      <div class="section-head"><h2>만들어질 폴더 미리보기</h2><span class="count" id="mkInfo"></span></div>
      <div id="mkTree" class="tree"><div class="empty">위에서 폴더 이름을 입력하면 만들어질 구조가 표시됩니다.</div></div>
    </div>
  </section>

  <!-- ===== PDF · 변환 탭 ===== -->
  <section id="tab-pdf" style="display:none">
    <div class="section-head"><h2>PDF · 변환</h2><span class="count">작업을 고르고, 아래 목록에서 파일을 선택하세요</span></div>
    <div class="chips" id="pdfOps">
      <button class="chip active" data-op="merge" onclick="pickPdfOp(this)">PDF 병합</button>
      <button class="chip" data-op="split" onclick="pickPdfOp(this)">PDF 분할</button>
      <button class="chip" data-op="img2pdf" onclick="pickPdfOp(this)">이미지 → PDF</button>
      <button class="chip" data-op="pdf2img" onclick="pickPdfOp(this)">PDF → 이미지</button>
      <button class="chip" data-op="img2img" onclick="pickPdfOp(this)">이미지 형식 변환</button>
      <button class="chip" data-op="compress" onclick="pickPdfOp(this)">PDF 용량 줄이기</button>
      <button class="chip" data-op="excel2pdf" onclick="pickPdfOp(this)">엑셀 → PDF</button>
      <button class="chip" data-op="hwp2pdf" onclick="pickPdfOp(this)">한글 → PDF</button>
      <button class="chip" data-op="xlsmerge" onclick="pickPdfOp(this)">엑셀 시트 합치기</button>
      <button class="chip" data-op="xlsappend" onclick="pickPdfOp(this)">엑셀 행 통합</button>
    </div>

    <div id="pdfOptions"></div>

    <div class="pdf-listhead">
      <span class="prev-tools" id="pdfAllWrap">
        <button class="btn btn-secondary btn-sm" onclick="pdfSelectAll()">전체 선택</button>
        <button class="btn btn-secondary btn-sm" onclick="pdfDeselectAll()">전체 해제</button>
      </span>
      <span class="hint" id="pdfListInfo"></span>
    </div>
    <div id="pdfList"><div class="empty">상단의 <b>폴더 선택</b>으로 폴더를 지정하세요.</div></div>

    <div class="pdf-actions">
      <button class="btn btn-primary" id="pdfRun" onclick="runPdfOp()" disabled>실행</button>
      <span id="pdfResult"></span>
    </div>
  </section>

  <!-- ===== 미리보기 ===== -->
  <section id="previewSection">
    <div class="section-head"><h2>Preview</h2>
      <span class="prev-tools">
        <button class="btn btn-secondary btn-sm" onclick="selectAll()">전체 선택</button>
        <button class="btn btn-secondary btn-sm" onclick="deselectAll()">전체 해제</button>
        <span class="count" id="previewCount"></span>
      </span>
    </div>
    <div class="prev-hint">행을 클릭하면 포함/제외가 전환됩니다 · 행 위를 드래그하면 연속된 파일을 한 번에 선택/해제할 수 있습니다 · 왼쪽 <b>⠿</b> 손잡이를 잡고 드래그하면 순서를 바꿀 수 있습니다(순번 매기기에 반영).</div>
    <div class="thead"><span style="margin-left:6px">포함</span><span>현재 이름</span><span class="col-text">붙일 텍스트</span><span>변경 후</span><span>상태</span></div>
    <div id="rows"><div class="empty">폴더를 선택하면 파일 목록이 표시됩니다. 행을 클릭하면 제외/포함이 전환됩니다.</div></div>
  </section>
  <details id="histBox">
    <summary>최근 작업 이력 (이 컴퓨터에만 저장됩니다 · 최대 30건)</summary>
    <div id="histRows"></div>
  </details>
</main>

<div class="actionbar">
  <span class="status" id="statusBar">대기 중</span>
  <button class="btn btn-secondary" id="btnOpen" onclick="openFolder()" disabled title="탐색기에서 이 폴더 열기">폴더 열기</button>
  <button class="btn btn-secondary" id="btnUndo" onclick="doUndo()" disabled>실행 취소 (Ctrl+Z)</button>
  <button class="btn btn-primary" id="btnRun" onclick="askExecute()" disabled>변경 실행 (Ctrl+Enter)</button>
</div>

<div class="scrim" id="scrim">
  <div class="dialog">
    <h3 id="dlgTitle"></h3>
    <p id="dlgBody"></p>
    <div class="row">
      <button class="btn btn-secondary" id="dlgCancel" onclick="closeDialog()">취소</button>
      <button class="btn btn-primary" id="dlgOk">실행</button>
    </div>
  </div>
</div>

<footer>© 2026 FileForge — Windows Edition. 모든 변경은 실행 취소가 가능하며, 충돌이 감지된 파일은 자동으로 제외됩니다.</footer>

<script>
const INVALID = /[\\/:*?"<>|]/;
const CATEGORIES = {
  "문서":["doc","docx","hwp","hwpx","pdf","txt","rtf","odt","md"],
  "엑셀·데이터":["xls","xlsx","xlsm","csv","tsv","ods"],
  "프레젠테이션":["ppt","pptx","key","odp"],
  "이미지":["jpg","jpeg","png","gif","bmp","tif","tiff","webp","heic","svg","raw"],
  "영상":["mp4","avi","mkv","mov","wmv","flv","webm","m4v"],
  "음악":["mp3","wav","flac","m4a","wma","aac","ogg"],
  "압축":["zip","rar","7z","tar","gz","egg","alz"],
  "실행·설치":["exe","msi","bat","cmd","apk","dmg"],
};

let state = {folder:"", files:[], allNames:[], dirs:[], tab:"rename", org:"category",
             preview:[], canUndo:false, perText:{}, nApply:0, nConflict:0,
             pdfOp:"merge", pdfFiles:[], pdfSelected:new Set(), pdfRanges:{}, keywords:[], mapRules:[],
             mapFolders:[], mapScan:[],
             renameTarget:"files", loadedInclude:"", mkMode:"list", mkPlan:null};
const $ = id => document.getElementById(id);

async function api(path, body){
  const r = await fetch(path,{method:"POST",headers:{"Content-Type":"application/json"},
                              body:JSON.stringify(body||{})});
  return r.json();
}

async function pickFolder(){
  const res = await api("/api/pick");
  if(res.path){ state.folder = res.path; await reload(); }
}

let reloadTimer;
function debounceReload(){ clearTimeout(reloadTimer); reloadTimer=setTimeout(()=>reload(),300); }

// 지금 화면에 필요한 목록 종류 — '이름 변경' 탭에서만 폴더를 함께 불러온다
function neededInclude(){
  return state.tab==="rename" ? (state.renameTarget || "files") : "files";
}

async function reload(silent){
  if(!state.folder) return;
  const exts = $("extFilter").value.split(",").map(s=>s.trim().toLowerCase().replace(/^\./,"")).filter(Boolean);
  const include = neededInclude();
  const res = await api("/api/list",{folder:state.folder, exts, include});
  if(res.error){
    if(silent){   // 시작할 때 되살린 폴더가 없어졌을 뿐이니 부드럽게 안내
      state.folder = ""; $("btnOpen").disabled = true;
      setStatus("이전에 쓰던 폴더를 찾을 수 없습니다. 위쪽 <b>폴더 선택</b>으로 다시 지정해 주세요.", false, true);
    }else setStatus(res.error, true);
    return;
  }
  res.files.forEach(f=>f.excluded=false);
  state.files = res.files; state.allNames = res.allNames; state.dirs = res.dirs || []; state.perText = {};
  state.loadedInclude = include;
  document.body.classList.add("loaded");
  $("heroTitle").innerHTML = "Forge Every File.";
  const nDir = state.files.filter(f=>f.isDir).length, nFile = state.files.length - nDir;
  $("heroPath").textContent = state.folder + "  ·  "
    + (include==="files" ? `파일 ${nFile}개` : `파일 ${nFile}개 · 폴더 ${nDir}개`);
  $("btnOpen").disabled = false;
  savePrefs();
  recompute();
  if(state.tab==="pdf") loadPdfFiles();
}

async function switchTab(tab){
  state.tab = tab;
  document.querySelectorAll(".nav-tab").forEach(b=>b.classList.toggle("active", b.dataset.tab===tab));
  $("tab-rename").style.display = tab==="rename" ? "" : "none";
  $("tab-organize").style.display = tab==="organize" ? "" : "none";
  $("tab-make").style.display = tab==="make" ? "" : "none";
  $("tab-pdf").style.display = tab==="pdf" ? "" : "none";
  $("previewSection").style.display = (tab==="pdf" || tab==="make") ? "none" : "";
  document.querySelector(".actionbar").style.display = tab==="pdf" ? "none" : "";
  $("btnRun").textContent = tab==="organize" ? "정리 실행 (Ctrl+Enter)"
                          : tab==="make" ? "폴더 만들기 (Ctrl+Enter)" : "변경 실행 (Ctrl+Enter)";
  savePrefs();
  if(tab==="pdf"){ renderPdfOptions(); loadPdfFiles(); return; }
  // 탭마다 필요한 목록이 달라서(폴더 포함 여부) 다르면 다시 불러온다
  if(state.folder && neededInclude() !== state.loadedInclude){ await reload(); return; }
  recompute();
}

// 이름 변경 대상(파일/폴더) 전환
async function onTargetChange(){
  state.renameTarget = $("renameTarget").value;
  const dirsOnly = state.renameTarget === "dirs";
  $("extFilter").disabled = dirsOnly;                       // 폴더는 확장자가 없음
  $("extFilter").style.opacity = dirsOnly ? ".35" : "";
  $("extFilter").placeholder = dirsOnly ? "폴더는 확장자 없음" : "확장자 필터 (jpg,png)";
  savePrefs();
  if(state.folder) await reload(); else recompute();
}

function toggleRule(chip){
  chip.classList.toggle("active");
  $("panel-"+chip.dataset.rule).classList.toggle("show", chip.classList.contains("active"));
  // '키워드별 지정값으로 변경'을 처음 켜면 빈 줄 10개를 보여준다
  if(chip.dataset.rule==="mapname" && chip.classList.contains("active") && state.mapRules.length===0)
    initMapRows(10);
  recompute();
}
function ruleOn(name){
  const c = document.querySelector('.chip[data-rule="'+name+'"]');
  return c && c.classList.contains("active");
}
function pickOrg(chip){
  document.querySelectorAll("#orgChips .chip").forEach(c=>c.classList.remove("active"));
  chip.classList.add("active");
  state.org = chip.dataset.org;
  const kw = state.org==="keyword";
  $("kwBox").style.display = kw ? "" : "none";
  $("orgHint").style.display = kw ? "none" : "";
  if(kw && state.keywords.length===0){ state.keywords=["",""]; renderKwFields(); }
  recompute();
}

// ----- 키워드별 정리: 동적 입력 필드 -----
function sanitizeFolder(s){ return s.trim().replace(/[\\/:*?"<>|]/g,"_"); }

function renderKwFields(){
  let h = "";
  state.keywords.forEach((k,i)=>{
    h += `<div class="kwrow"><span class="kwnum">${i+1}</span>
      <input class="field" value="${escA(k)}" placeholder="예: 계약, 2024, 회의록…"
             oninput="onKwInput(${i}, this.value)">
      <button class="kwdel" onclick="removeKeyword(${i})" title="삭제">×</button></div>`;
  });
  $("kwFields").innerHTML = h;
}
function onKwInput(i, val){ state.keywords[i] = val; recompute(); }

// ----- 키워드별 지정값으로 변경: A → B 매핑 입력 (Enter로 줄 추가) -----
function initMapRows(n){
  state.mapRules = [];
  for(let i=0;i<n;i++) state.mapRules.push({a:"", b:""});
  renderMapFields();
}
function renderMapFields(){
  let h = "";
  state.mapRules.forEach((m,i)=>{
    h += `<div class="maprow"><span class="mapnum">${i+1}</span>
      <input class="field" value="${escA(m.a)}" placeholder="기존 이름 (예: 1)"
             oninput="onMapInput(${i},'a',this.value)" onkeydown="onMapKey(event,${i},'a')">
      <span class="maparrow">→</span>
      <input class="field" value="${escA(m.b)}" placeholder="바꿀 이름 (예: 1_이력서)"
             oninput="onMapInput(${i},'b',this.value)" onkeydown="onMapKey(event,${i},'b')">
      <button class="mapdel" onclick="removeMapRow(${i})" title="삭제">×</button></div>`;
  });
  $("mapFields").innerHTML = h;
}
function onMapInput(i, key, val){ state.mapRules[i][key] = val; recompute(); scheduleMapScan(); }
// i번째 줄의 'a'(왼쪽)/'b'(오른쪽) 칸으로 커서를 옮긴다
function focusMapCell(i, key){
  const row = $("mapFields").querySelectorAll(".maprow")[i];
  if(!row) return;
  const inp = row.querySelectorAll("input")[key === "b" ? 1 : 0];
  if(inp) inp.focus();
}
function onMapKey(e, i, key){
  if(e.key !== "Enter") return;
  // 한글 등 조합 입력 중의 Enter 는 '조합 확정'용이다. 여기서 가로채면 마지막 글자가
  // 확정되지 않은 채 넘어갈 수 있으므로, 확정용 Enter 는 그대로 흘려보낸다.
  if(e.isComposing || e.keyCode === 229) return;
  e.preventDefault();
  // Ctrl+Enter 는 전역에서 '실행'에 걸려 있다. 입력칸 안에서는 줄 추가로만 쓰고
  // 전역 핸들러까지 올라가지 않도록 여기서 막는다.
  e.stopPropagation();
  if(e.ctrlKey || e.metaKey){          // Ctrl+Enter: 이 줄 바로 아래에 새 줄
    addMapRow(i + 1, key);
    return;
  }
  if(i + 1 >= state.mapRules.length)   // Enter: 아랫줄로. 마지막 줄이면 이어서 칠 수 있게 늘린다
    addMapRow(i + 1, key);
  else
    focusMapCell(i + 1, key);
}
function addMapRow(at, key){
  const idx = (typeof at === "number") ? at : state.mapRules.length;
  state.mapRules.splice(idx, 0, {a:"", b:""});
  renderMapFields(); recompute();
  focusMapCell(idx, key || "a");
}
// ----- 키워드별 지정값으로 변경: 여러 폴더에 한꺼번에 적용 -----
async function pickManyFolders(){
  const res = await api("/api/pick_many");
  const paths = res.paths || [];
  if(!paths.length) return;                 // 취소했거나 고른 게 없음
  state.mapFolders = paths;
  await scanMapFolders();
}
function clearMapFolders(){
  state.mapFolders = []; state.mapScan = [];
  renderMapFolders();
}
let mapScanTimer;
function scheduleMapScan(){
  if(!state.mapFolders.length) return;
  clearTimeout(mapScanTimer);
  mapScanTimer = setTimeout(()=>scanMapFolders(), 300);   // 타이핑이 멈추면 다시 계산
}
async function scanMapFolders(){
  if(!state.mapFolders.length){ renderMapFolders(); return; }
  const res = await api("/api/map_scan", {folders:state.mapFolders, rules:state.mapRules});
  state.mapScan = res.folders || [];
  renderMapFolders();
}
function mapJobs(){
  return (state.mapScan || []).reduce((all, s)=>all.concat(s.jobs || []), []);
}
function openMapFolder(i){ openPath(state.mapFolders[i]); }
function renderMapFolders(){
  const box = $("mapFolderList"), row = $("mapApplyRow"), clear = $("mapClear");
  if(!box) return;
  if(!state.mapFolders.length){
    box.innerHTML = ""; row.style.display = "none"; clear.style.display = "none"; return;
  }
  clear.style.display = "";
  const scan = state.mapScan || [];
  box.innerHTML = state.mapFolders.map((f,i)=>{
    const s = scan[i] || {}, hit = (s.jobs || []).length;
    const info = s.missing
      ? '<span class="mfcount warn">폴더를 찾을 수 없음</span>'
      : `<span class="mfcount ${hit ? "ok" : "hint"}">${hit}개 바뀜</span>`;
    return `<div class="mfrow"><span class="mfpath" title="${escA(f)}">${esc(f)}</span>${info}`
      + `<button class="gobtn" onclick="openMapFolder(${i})">탐색기에서 열기</button></div>`;
  }).join("");
  const jobs = mapJobs();
  row.style.display = "";
  $("mapApplyBtn").disabled = !jobs.length;
  $("mapApplyInfo").innerHTML = jobs.length
    ? `폴더 ${state.mapFolders.length}개에서 <b class="ok">${jobs.length}개</b> 파일 이름이 바뀝니다.`
    : "규칙에 해당하는 파일이 없습니다. 위 표에 왼쪽·오른쪽 값을 입력하세요.";
}
function askApplyMap(){
  const jobs = mapJobs();
  if(!jobs.length) return;
  const sample = jobs.slice(0,5)
    .map(j=>`${esc(j.src)} <span class="hint">→</span> <b>${esc(j.dst)}</b>`).join("<br>");
  showDialog("여러 폴더에 적용",
    `폴더 <b>${state.mapFolders.length}개</b>에서 <b>${jobs.length}개</b> 파일의 이름을 바꿉니다.<br><br>`
      + sample + (jobs.length > 5 ? `<br><span class="hint">… 외 ${jobs.length-5}개</span>` : "")
      + `<br><br>실행한 뒤 <b>실행 취소</b>로 한 번에 되돌릴 수 있습니다.`,
    `${jobs.length}개 이름 바꾸기`, applyMapFolders);
}
async function applyMapFolders(){
  const jobs = mapJobs();
  if(!jobs.length) return;
  const nFolders = state.mapFolders.length;
  const res = await api("/api/execute", {folders:state.mapFolders, jobs, mode:"rename"});
  if(res.error){ setStatus(res.error, true); return; }
  state.canUndo = res.canUndo; $("btnUndo").disabled = !res.canUndo;
  await scanMapFolders();
  if(state.folder) await reload();          // 지금 보고 있는 폴더 목록도 갱신
  let msg = `<b class="ok">폴더 ${nFolders}개에서 ${res.done}개 이름 변경 완료</b> · 실행 취소 가능`;
  if(res.errors.length) msg += ` · <span class="warn">실패 ${res.errors.length}건: ${esc(res.errors[0])}</span>`;
  msg += ` <button class="gobtn" onclick="openMapFolder(0)">탐색기에서 열기</button>`;
  setStatus(msg, false, true);
  addHistory("여러 폴더 이름 변경",
             `폴더 ${nFolders}개 · ${res.done}개 성공`
             + (res.errors.length ? ` · 실패 ${res.errors.length}건` : ""),
             state.mapFolders[0]);
}

function removeMapRow(i){
  state.mapRules.splice(i,1);
  if(!state.mapRules.length) state.mapRules.push({a:"", b:""});
  renderMapFields(); recompute();
}
function removeKeyword(i){ state.keywords.splice(i,1); if(!state.keywords.length) state.keywords=[""]; renderKwFields(); recompute(); }
function addKeyword(){ state.keywords.push(""); renderKwFields(); recompute();
  const f = $("kwFields").querySelectorAll(".field"); if(f.length) f[f.length-1].focus(); }

let kwDrag = null;
function startKwDrag(e){
  e.preventDefault();
  state.keywords.push("");           // 누르는 순간 1개 추가 (클릭=1개)
  renderKwFields(); recompute();
  kwDrag = {startY:e.clientY, base:state.keywords.length};
  document.addEventListener("mousemove", onKwDragMove);
  document.addEventListener("mouseup", onKwDragUp);
}
function onKwDragMove(e){
  const extra = Math.max(0, Math.floor((e.clientY - kwDrag.startY) / 30));  // 30px마다 1개 더
  const want = kwDrag.base + extra;
  let changed = false;
  while(state.keywords.length < want){ state.keywords.push(""); changed = true; }
  $("kwAddHint").textContent = extra>0 ? `+${extra}개 더 추가` : "";
  if(changed){ renderKwFields(); recompute(); }
}
function onKwDragUp(){
  document.removeEventListener("mousemove", onKwDragMove);
  document.removeEventListener("mouseup", onKwDragUp);
  $("kwAddHint").textContent = "";
  kwDrag = null;
  const f = $("kwFields").querySelectorAll(".field"); if(f.length) f[f.length-1].focus();
}

function fmtDate(d, fmt){
  const p = n => String(n).padStart(2,"0");
  const Y=d.getFullYear(), M=p(d.getMonth()+1), D=p(d.getDate());
  return {"YYYYMMDD":`${Y}${M}${D}`, "YYYY-MM-DD":`${Y}-${M}-${D}`,
          "YYMMDD":`${String(Y).slice(2)}${M}${D}`, "YYYY.MM.DD":`${Y}.${M}.${D}`}[fmt];
}

// 폴더는 확장자가 없으므로 이름 전체가 stem — 파일만 마지막 점 뒤를 확장자로 본다
function splitName(name, isDir){
  if(isDir) return [name, ""];
  const i = name.lastIndexOf(".");
  return i>0 ? [name.slice(0,i), name.slice(i)] : [name, ""];
}

function buildNewName(f, seq){
  const name = f.name, mtime = f.mtime;
  let [stem, ext] = splitName(name, f.isDir);

  if(ruleOn("rebuild")){
    const mode = $("rebuildMode").value;
    if(mode==="number"){
      const base = $("rebuildBase").value;
      const num  = (parseInt($("rebuildStart").value)||0) + seq;
      const ns   = String(num).padStart(Math.max(1,parseInt($("rebuildPad").value)||1),"0");
      stem = base + ns;
    }else if(mode==="common"){
      const base = $("rebuildCommon").value;
      // 이름을 입력했을 때만 교체. 파일이 여러 개면 겹치지 않도록 _1, _2, _3… 을 붙인다.
      if(base) stem = state.includedCount>1 ? base+"_"+(seq+1) : base;
    }else{ // perfile — 입력한 값이 있을 때만 통째 교체 (빈 칸은 기존 이름 유지)
      const t = state.perText[name];
      if(t) stem = t;
    }
  }
  if(ruleOn("mapname")){
    // 기존 이름(확장자 제외)이 왼쪽 값과 정확히 같으면 오른쪽 값으로 통째 교체. 위쪽 줄이 우선.
    const key = stem.trim().toLowerCase();
    for(const m of state.mapRules){
      const a = (m.a||"").trim(), b = (m.b||"").trim();
      if(a && b && a.toLowerCase()===key){ stem = b; break; }
    }
  }
  if(ruleOn("replace") && $("findText").value){
    try{
      if($("useRegex").checked) stem = stem.replace(new RegExp($("findText").value,"g"), $("replaceText").value);
      else stem = stem.split($("findText").value).join($("replaceText").value);
    }catch(e){ /* 정규식 입력 중 오류는 무시 */ }
  }
  if(ruleOn("clean")){
    if($("cleanSpecial").checked) stem = stem.replace(/[!@#$%^&()\[\]{};'`~+=,]/g," ");
    if($("cleanSpaces").checked)  stem = stem.replace(/\s+/g," ").trim();
    if($("cleanUnderscore").checked) stem = stem.replace(/ /g,"_");
  }
  if(ruleOn("case")){
    const m = $("caseMode").value;
    if(m==="lower") stem = stem.toLowerCase();
    else if(m==="upper") stem = stem.toUpperCase();
    else stem = stem.replace(/\w\S*/g, t=>t[0].toUpperCase()+t.slice(1).toLowerCase());
  }
  if(ruleOn("affix")) stem = $("prefix").value + stem + $("suffix").value;
  if(ruleOn("pertext")){
    const t = state.perText[name] || "";
    if(t) stem = $("perTextPos").value==="front" ? t+stem : stem+t;
  }
  if(ruleOn("date")){
    const d = $("dateSource").value==="today" ? new Date() : new Date(mtime*1000);
    const ds = fmtDate(d, $("dateFormat").value);
    stem = $("datePos").value==="front" ? ds+"_"+stem : stem+"_"+ds;
  }
  if(ruleOn("seq")){
    const num = (parseInt($("seqStart").value)||0) + seq;
    const ns = String(num).padStart(Math.max(1,parseInt($("seqPad").value)||1),"0");
    stem = $("seqPos").value==="front" ? ns+"_"+stem : stem+"_"+ns;
  }
  if(ruleOn("case") && $("extLower").checked) ext = ext.toLowerCase();
  return stem + ext;
}

function targetFolder(name, mtime){
  const i = name.lastIndexOf(".");
  const ext = i>0 ? name.slice(i+1).toLowerCase() : "";
  if(state.org==="ext") return ext ? ext.toUpperCase() : "확장자없음";
  if(state.org==="category"){
    for(const [cat,exts] of Object.entries(CATEGORIES)) if(exts.includes(ext)) return cat;
    return "기타";
  }
  if(state.org==="keyword"){
    const low = name.toLowerCase();
    for(const raw of state.keywords){
      const k = raw.trim();
      if(k && low.includes(k.toLowerCase())) return sanitizeFolder(k);
    }
    return null;  // 어떤 키워드에도 안 맞으면 이동 안 함
  }
  const d = new Date(mtime*1000), p = n=>String(n).padStart(2,"0");
  const ym = d.getFullYear()+"-"+p(d.getMonth()+1);
  return state.org==="day" ? ym+"-"+p(d.getDate()) : ym;
}

function computePreview(){
  const organize = state.tab==="organize";
  state.preview = [];
  state.includedCount = state.files.filter(f=>!f.excluded).length;
  let seq = 0;
  for(const f of state.files){
    if(f.excluded){ state.preview.push({disp:"—", st:"제외", target:null}); continue; }
    if(organize){
      const sub = targetFolder(f.name, f.mtime);
      if(sub === null) state.preview.push({disp:"—", st:"해당 없음", target:null});
      else state.preview.push({disp:sub+"/"+f.name, st:"이동", target:sub+"/"+f.name});
    }else{
      const nn = buildNewName(f, seq++);
      const stem = splitName(nn, f.isDir)[0];
      if(INVALID.test(nn) || !stem.trim()) state.preview.push({disp:nn, st:"⚠ 사용불가", target:null});
      else if(nn===f.name) state.preview.push({disp:nn, st:"변경 없음", target:null});
      else state.preview.push({disp:nn, st:"변경", target:nn});
    }
  }
  // 충돌 검사: 결과끼리 중복(대소문자 무시) + 폴더에 이미 존재
  const seen = {};
  state.preview.forEach((p,idx)=>{
    if(!p.target) return;
    const k = p.target.toLowerCase();
    if(k in seen){ p.st="⚠ 중복"; p.target=null;
                   const o=state.preview[seen[k]]; o.st="⚠ 중복"; o.target=null; }
    else seen[k]=idx;
  });
  if(!organize){
    const sources = new Set(state.files.map(f=>f.name.toLowerCase()));
    const existing = new Set(state.allNames.map(n=>n.toLowerCase()));
    state.preview.forEach(p=>{
      if(p.target && existing.has(p.target.toLowerCase()) && !sources.has(p.target.toLowerCase())){
        p.st="⚠ 이미 존재"; p.target=null;
      }
    });
  }
  let nApply=0, nConflict=0;
  state.files.forEach((f,i)=>{
    const p = state.preview[i];
    if(f.excluded) return;
    if(p.st.startsWith("⚠")) nConflict++;
    else if(p.target) nApply++;
  });
  state.nApply = nApply; state.nConflict = nConflict;
}

function rowClass(i){
  const f = state.files[i], p = state.preview[i];
  if(f.excluded) return "excluded";
  if(p.st.startsWith("⚠")) return "conflict";
  if(p.target) return "changed";
  return "";
}

function updateStatusBar(){
  if(state.tab==="make") return;
  const nDir = state.files.filter(f=>f.isDir).length;
  const label = nDir ? `파일·폴더 ${state.files.length}개` : `파일 ${state.files.length}개`;
  $("previewCount").textContent = label;
  $("btnRun").disabled = state.nApply===0;
  let msg = `${nDir?"항목":"파일"} <b>${state.files.length}</b>개 · 적용 대상 <b class="ok">${state.nApply}</b>개`;
  if(state.nConflict) msg += ` · <span class="warn">⚠ 충돌 ${state.nConflict}개 자동 제외</span>`;
  setStatus(msg, false, true);
}

// '파일별 텍스트' 칸을 보여줄지 여부 (pertext 규칙 또는 이름 새로 입력=파일별 모드)
function rebuildPerfileOn(){
  return ruleOn("rebuild") && $("rebuildMode") && $("rebuildMode").value==="perfile";
}
function onRebuildModeChange(){
  const m = $("rebuildMode").value;
  $("rebuildNumberOpts").style.display  = m==="number"  ? "" : "none";
  $("rebuildCommonOpts").style.display  = m==="common"  ? "" : "none";
  $("rebuildPerfileOpts").style.display = m==="perfile" ? "" : "none";
  recompute();
}

// 전체 재계산 + 표 전체 다시 그리기 (규칙 토글·폴더 로드·포함 전환 등)
function recompute(){
  if(state.tab==="make"){ renderMake(); return; }   // 폴더 만들기 탭은 별도 미리보기
  document.body.classList.toggle("textcol", state.tab!=="organize" && (ruleOn("pertext") || rebuildPerfileOn()));
  computePreview();
  render();
  if(state.tab==="organize") renderOrgTree();
}

// ----- 폴더 정리: 폴더 구조 미리보기 -----
function renderOrgTree(){
  const el = $("orgTree"); if(!el) return;
  const info = $("orgTreeInfo");
  if(!state.folder){
    el.innerHTML = '<div class="empty">폴더를 선택하면 생성될 폴더 구조가 표시됩니다.</div>';
    if(info) info.textContent = ""; return;
  }
  const groups = new Map();
  let skipped = 0;
  state.files.forEach(f=>{
    if(f.excluded) return;
    const sub = targetFolder(f.name, f.mtime);
    if(sub === null){ skipped++; return; }
    if(!groups.has(sub)) groups.set(sub, []);
    groups.get(sub).push(f.name);
  });
  if(groups.size === 0){
    el.innerHTML = '<div class="empty">이동할 파일이 없습니다. (조건에 맞는 파일이 없거나 모두 제외됨)</div>';
    if(info) info.textContent = ""; return;
  }
  const dirsLower = new Set((state.dirs||[]).map(d=>d.toLowerCase()));
  const names = [...groups.keys()].sort((a,b)=>a.localeCompare(b,'ko'));
  const rootName = state.folder.replace(/[\\\/]+$/,'').split(/[\\\/]/).pop() || state.folder;
  let moved = 0, LIM = 8;
  let h = `<div class="troot"><span class="ticon">📂</span><span class="tname">${esc(rootName)}</span></div><div class="tchildren">`;
  names.forEach(sub=>{
    const files = groups.get(sub); moved += files.length;
    const exists = dirsLower.has(sub.toLowerCase());
    h += `<div class="tfolder"><span class="ticon">📁</span><span class="tname">${esc(sub)}</span>`
       + `<span class="tcount">${files.length}개</span>`
       + (exists ? `<span class="tbadge texist">기존</span>` : `<span class="tbadge tnew">신규</span>`)
       + `</div><div class="tfiles">`;
    files.slice(0, LIM).forEach(n=>{
      h += `<div class="tfile"><span class="ticon">📄</span><span>${esc(n)}</span></div>`;
    });
    if(files.length > LIM) h += `<div class="tmore">+${files.length - LIM}개 더</div>`;
    h += `</div>`;
  });
  h += `</div>`;
  el.innerHTML = h;
  if(info){
    let t = `${groups.size}개 폴더 · ${moved}개 파일 이동`;
    if(skipped) t += ` · ${skipped}개 해당 없음`;
    info.textContent = t;
  }
}

// 입력칸 포커스를 유지하기 위해, 표를 다시 만들지 않고 결과 칸만 갱신
function refreshComputed(){
  computePreview();
  document.querySelectorAll("#rows .frow").forEach(r=>{
    const i = parseInt(r.dataset.index);
    const p = state.preview[i]; if(!p) return;
    r.className = "frow " + rowClass(i);
    const nw = r.querySelector(".new"); if(nw) nw.textContent = p.disp;
    const st = r.querySelector(".st"); if(st) st.textContent = p.st;
  });
  updateStatusBar();
}

function render(){
  const box = $("rows");
  if(!state.files.length){
    box.innerHTML = '<div class="empty">' + (state.folder ?
      "조건에 맞는 파일이 없습니다." :
      "폴더를 선택하면 파일 목록이 표시됩니다. 행을 클릭하면 제외/포함이 전환됩니다.") + '</div>';
    $("previewCount").textContent=""; $("btnRun").disabled=true;
    setStatus(state.folder ? "파일 0개" : "대기 중"); return;
  }
  let html="";
  state.files.forEach((f,i)=>{
    const p = state.preview[i];
    const tv = state.perText[f.name] || "";
    html += `<div class="frow ${rowClass(i)}" data-index="${i}" onmousedown="startRowToggle(event,${i})">
      <span class="firstcell"><span class="grip" title="드래그하여 순서 변경" onmousedown="startReorder(event,${i})">⠿</span><span class="dot"></span></span>
      <span class="old">${f.isDir?'<span class="ficon" title="폴더">📁</span>':''}${esc(f.name)}</span>
      <span class="col-text"><span class="cell"><input class="tcell" value="${escA(tv)}" placeholder="텍스트" oninput="onPerTextInput(${i},this.value)" onpaste="onPerTextPaste(event,${i})" onclick="event.stopPropagation()"><span class="fillh" onmousedown="startFill(event,${i})" title="드래그하여 같은 값 채우기"></span></span></span>
      <span class="new">${(f.isDir&&p.target)?'<span class="ficon">📁</span>':''}${esc(p.disp)}</span>
      <span class="st">${p.st}</span></div>`;
  });
  box.innerHTML = html;
  updateStatusBar();
}

function esc(s){ return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
function escA(s){ return esc(s).replace(/"/g,"&quot;"); }

// ----- 전체 선택 / 전체 해제 -----
function selectAll(){ if(!state.files.length) return; state.files.forEach(f=>f.excluded=false); recompute(); }
function deselectAll(){ if(!state.files.length) return; state.files.forEach(f=>f.excluded=true); recompute(); }

// ----- 행 드래그로 연속 선택/해제 (포함 ↔ 제외) -----
let rowDrag = null;
function startRowToggle(e, i){
  if(e.button !== 0) return;
  // 손잡이·입력칸·채우기 핸들 위에서는 동작하지 않음 (각자 기능 수행)
  if(e.target.closest('.grip,.tcell,.fillh,input,select,button,a')) return;
  e.preventDefault();
  const val = !state.files[i].excluded;   // 시작 행의 반대 상태를 끌고 다니며 칠함
  rowDrag = {val};
  state.files[i].excluded = val;
  recompute();
  document.addEventListener("mousemove", onRowToggleMove);
  document.addEventListener("mouseup", onRowToggleUp);
}
function onRowToggleMove(e){
  if(!rowDrag) return;
  const t = document.elementFromPoint(e.clientX, e.clientY);
  const row = t && t.closest && t.closest('.frow');
  if(!row) return;
  const idx = parseInt(row.dataset.index);
  if(isNaN(idx)) return;
  if(state.files[idx].excluded !== rowDrag.val){
    state.files[idx].excluded = rowDrag.val;
    recompute();
  }
}
function onRowToggleUp(){
  document.removeEventListener("mousemove", onRowToggleMove);
  document.removeEventListener("mouseup", onRowToggleUp);
  rowDrag = null;
}

// ----- 손잡이 드래그로 순서 변경 (순번 매기기에 반영) -----
let reorder = null;
function startReorder(e, i){
  if(e.button !== 0) return;
  e.preventDefault(); e.stopPropagation();
  reorder = {from:i};
  document.body.style.userSelect = "none";
  document.addEventListener("mousemove", onReorderMove);
  document.addEventListener("mouseup", onReorderUp);
}
function onReorderMove(e){
  if(!reorder) return;
  const t = document.elementFromPoint(e.clientX, e.clientY);
  const row = t && t.closest && t.closest('.frow');
  if(!row) return;
  const to = parseInt(row.dataset.index);
  if(isNaN(to) || to === reorder.from) return;
  const arr = state.files;
  const [m] = arr.splice(reorder.from, 1);
  arr.splice(to, 0, m);
  reorder.from = to;
  recompute();   // 순번 재계산 + 표 다시 그리기
  const moved = document.querySelector(`#rows .frow[data-index="${to}"]`);
  if(moved) moved.classList.add("dragging");
}
function onReorderUp(){
  document.removeEventListener("mousemove", onReorderMove);
  document.removeEventListener("mouseup", onReorderUp);
  document.body.style.userSelect = "";
  reorder = null;
  recompute();
}

// ----- 파일별 텍스트: 입력 / 붙여넣기 / 드래그 채우기 -----
function onPerTextInput(i, value){
  state.perText[state.files[i].name] = value;
  refreshComputed();  // 표를 다시 만들지 않아 입력 포커스 유지
}
function onPerTextPaste(e, i){
  const text = (e.clipboardData || window.clipboardData).getData("text");
  if(text && /\r?\n/.test(text)){   // 엑셀 등에서 여러 줄을 붙이면 연속된 행에 한 줄씩 채움
    e.preventDefault();
    const lines = text.split(/\r?\n/);
    for(let k=0; k<lines.length && i+k<state.files.length; k++)
      state.perText[state.files[i+k].name] = lines[k];
    recompute();
  }
}
let fillState = null;
function startFill(e, i){
  e.preventDefault(); e.stopPropagation();
  fillState = {src:i, value:state.perText[state.files[i].name]||"", lo:i, hi:i};
  document.body.style.userSelect = "none";
  document.addEventListener("mousemove", onFillMove);
  document.addEventListener("mouseup", onFillUp);
}
function onFillMove(e){
  const el = document.elementFromPoint(e.clientX, e.clientY);
  const row = el && el.closest(".frow");
  if(row){
    const idx = parseInt(row.dataset.index);
    if(!isNaN(idx)){ fillState.lo = Math.min(fillState.src, idx); fillState.hi = Math.max(fillState.src, idx); }
  }
  document.querySelectorAll("#rows .frow").forEach(r=>{
    const i = parseInt(r.dataset.index);
    r.classList.toggle("fill-target", i>=fillState.lo && i<=fillState.hi && i!==fillState.src);
  });
}
function onFillUp(){
  document.removeEventListener("mousemove", onFillMove);
  document.removeEventListener("mouseup", onFillUp);
  document.body.style.userSelect = "";
  if(fillState){
    for(let i=fillState.lo; i<=fillState.hi; i++)
      state.perText[state.files[i].name] = fillState.value;
    fillState = null;
    recompute();
  }
}
function setStatus(msg, isErr, isHtml){
  const el = $("statusBar");
  if(isHtml) el.innerHTML = msg; else el.textContent = msg;
  el.style.color = isErr ? "var(--sale)" : "";
}

// 확인 창 (제목·본문·버튼 문구를 상황에 맞게 바꿔 쓴다)
function showDialog(title, bodyHtml, okLabel, onOk, cancelLabel, onCancel){
  $("dlgTitle").textContent = title;
  $("dlgBody").innerHTML = bodyHtml;
  $("dlgOk").textContent = okLabel || "실행";
  $("dlgCancel").textContent = cancelLabel || "취소";
  $("dlgOk").onclick = ()=>{ closeDialog(); if(onOk) onOk(); };
  $("dlgCancel").onclick = ()=>{ closeDialog(); if(onCancel) onCancel(); };
  $("scrim").classList.add("show");
}

function askExecute(){
  if(state.tab==="make") return askMake();
  const organize = state.tab==="organize";
  const jobs = [];
  state.files.forEach((f,i)=>{
    const p = state.preview[i];
    if(p.target && !f.excluded) jobs.push({src:f.name, dst:p.target});
  });
  if(!jobs.length) return;
  const nDir = state.files.filter((f,i)=>f.isDir && state.preview[i].target && !f.excluded).length;
  const nFile = jobs.length - nDir;
  const what = [nFile ? `파일 ${nFile}개` : "", nDir ? `폴더 ${nDir}개` : ""].filter(Boolean).join(" · ");
  showDialog(organize ? "폴더 정리 실행" : "이름 변경 실행",
             `${what}에 적용합니다. 실행 후에도 <b>실행 취소</b>로 한 번에 되돌릴 수 있습니다.`,
             organize ? "정리 실행" : "변경 실행",
             ()=>doExecute(jobs, organize));
}
function closeDialog(){ $("scrim").classList.remove("show"); }

async function doExecute(jobs, organize){
  const res = await api("/api/execute",{folder:state.folder, jobs,
                                        mode:organize?"organize":"rename"});
  if(res.error){ setStatus(res.error, true); return; }
  state.canUndo = res.canUndo; $("btnUndo").disabled = !res.canUndo;
  await reload();
  let msg = `<b class="ok">${res.done}개 ${organize?"이동":"이름 변경"} 완료</b> · 실행 취소 가능`;
  if(res.errors.length) msg += ` · <span class="warn">실패 ${res.errors.length}건: ${esc(res.errors[0])}</span>`;
  msg += ` <button class="gobtn" onclick="openFolder()">탐색기에서 열기</button>`;
  setStatus(msg, false, true);
  addHistory(organize ? "폴더 정리" : "이름 변경",
             `${res.done}개 성공` + (res.errors.length ? ` · 실패 ${res.errors.length}건` : ""));
}

async function doUndo(){
  const res = await api("/api/undo");
  if(res.error){ setStatus(res.error, true); return; }
  $("btnUndo").disabled = !res.canUndo;
  await reload();
  const parts = [];
  if(res.restored) parts.push(`${res.restored}개를 원래 자리로 되돌림`);
  if(res.removedDirs) parts.push(`새로 만든 폴더 ${res.removedDirs}개 삭제`);
  setStatus(`<b>${parts.join(" · ") || "되돌릴 내용이 없습니다"}</b>`, false, true);
  if(parts.length) addHistory("실행 취소", parts.join(" · "));
}

// 결과를 바로 확인할 수 있게 탐색기(Finder)에서 폴더 열기
async function openPath(folder){
  if(!folder) return;
  const res = await api("/api/open", {folder});
  if(res.error) setStatus(res.error, true);   // 폴더가 지워졌거나 옮겨진 경우
}
async function openFolder(){ return openPath(state.folder); }

// 이력에 남은 '그때 작업한 폴더' 열기 (경로를 HTML 에 심지 않고 번호로 찾는다)
function openHist(i){
  const h = lsGet(HIST_KEY) || [];
  if(h[i]) openPath(h[i].folder);
}

// 이력에서 폴더만 중복 없이 최신순으로 뽑는다
function recentFolders(n){
  const h = lsGet(HIST_KEY) || [];
  const seen = new Set(), out = [];
  for(const r of h){
    const f = (r.folder || "").trim();
    if(!f || seen.has(f)) continue;
    seen.add(f); out.push(f);
    if(out.length >= (n || 8)) break;
  }
  return out;
}
function toggleRecent(e){
  e.stopPropagation();                 // 아래 document 리스너가 곧바로 닫지 않도록
  const menu = $("recentMenu");
  const list = recentFolders(8);
  menu.innerHTML = list.length
    ? list.map((f,i)=>`<button class="recent-item" onclick="openRecent(${i})" title="${escA(f)}">${esc(f)}</button>`).join("")
    : '<div class="recent-empty">아직 작업한 폴더가 없습니다.</div>';
  menu.classList.toggle("show");
}
function openRecent(i){
  const f = recentFolders(8)[i];
  $("recentMenu").classList.remove("show");
  openPath(f);
}
document.addEventListener("click", ()=>{
  const m = $("recentMenu");
  if(m) m.classList.remove("show");   // 바깥을 클릭하면 닫는다
});

// ================= 폴더 만들기 =================
function pickMk(chip){
  document.querySelectorAll("#mkChips .chip").forEach(c=>c.classList.remove("active"));
  chip.classList.add("active");
  state.mkMode = chip.dataset.mk;
  document.querySelectorAll("#tab-make .mkpanel").forEach(p=>p.classList.remove("show"));
  const p = $("mk-"+state.mkMode); if(p) p.classList.add("show");
  savePrefs();
  renderMake();
}

// 한 줄에 하나씩 입력된 값을 배열로
function mkLines(id){
  const el = $(id); if(!el) return [];
  return el.value.split(/[\r\n]+/).map(t=>t.trim()).filter(Boolean);
}
// 폴더 이름으로 쓸 수 없는 글자를 정리 ('/'는 하위 폴더 구분자로 살려 둔다)
function sanitizePath(v){
  return String(v||"").replace(/\\/g,"/").split("/")
    .map(seg=>seg.trim().replace(/[:*?"<>|]/g,"_").replace(/[. ]+$/,"").trim())
    .filter(Boolean).join("/");
}
function mkSample(){
  $("mkList").value = "01_기획\n02_디자인\n03_개발\n04_검수\n05_완료";
  renderMake(); scheduleSave();
}
function mkDedup(){
  const seen = new Set(), out = [];
  mkLines("mkList").forEach(n=>{ const k=n.toLowerCase(); if(!seen.has(k)){ seen.add(k); out.push(n); } });
  $("mkList").value = out.join("\n");
  renderMake(); scheduleSave();
}
async function mkFromClipboard(){
  try{
    const t = await navigator.clipboard.readText();
    if(!t){ setStatus("복사한 내용이 없습니다.", true); return; }
    const cur = $("mkList").value.trim();
    $("mkList").value = (cur ? cur+"\n" : "") + t.replace(/\t/g,"\n").trim();
    renderMake(); scheduleSave();
  }catch(e){
    setStatus("브라우저가 붙여넣기를 막았습니다. 입력칸을 클릭한 뒤 Ctrl+V 를 눌러 주세요.", true);
  }
}

function mkFmtDate(d, fmt){
  const p = n=>String(n).padStart(2,"0");
  const Y=d.getFullYear(), M=p(d.getMonth()+1), D=p(d.getDate());
  return ({"YYYY-MM":`${Y}-${M}`, "YYYY년 MM월":`${Y}년 ${M}월`, "YYYY-MM-DD":`${Y}-${M}-${D}`,
           "YYYYMMDD":`${Y}${M}${D}`, "MM월":`${M}월`})[fmt] || `${Y}-${M}-${D}`;
}

const MK_LIMIT = 500;   // 실수로 수천 개가 만들어지는 것을 막는 안전장치

// 화면 입력값 → 만들 폴더 목록(plan) 계산. 실제 만들기 전 미리보기에도 그대로 쓴다.
function makePlan(){
  const parent = sanitizePath(gv("mkParent",""));
  const subs = String(gv("mkSubs","")).split(/[,\n]/).map(sanitizePath).filter(Boolean);
  const mode = state.mkMode;
  let tops = [], warn = "";
  const groups = new Map();   // 파일 이름으로 만들기: 폴더명 → 그 폴더로 갈 파일들

  if(mode==="list"){
    tops = mkLines("mkList").map(sanitizePath);
  }else if(mode==="number"){
    const base = gv("mkNumBase",""), tail = gv("mkNumTail","");
    const start = parseInt(gv("mkNumStart","1"))||0;
    const pad = Math.min(6, Math.max(1, parseInt(gv("mkNumPad","2"))||1));
    let cnt = Math.max(0, parseInt(gv("mkNumCount","0"))||0);
    if(cnt > MK_LIMIT){ cnt = MK_LIMIT; warn = `한 번에 ${MK_LIMIT}개까지만 만듭니다.`; }
    for(let i=0;i<cnt;i++) tops.push(sanitizePath(base + String(start+i).padStart(pad,"0") + tail));
  }else if(mode==="date"){
    const from = gv("mkDateFrom",""), to = gv("mkDateTo","");
    if(from && to){
      const unit = gv("mkDateUnit","month"), fmt = gv("mkDateFmt","YYYY-MM"), pre = gv("mkDatePre","");
      let d = new Date(from+"T00:00:00"), end = new Date(to+"T00:00:00");
      if(!isNaN(d) && !isNaN(end) && d<=end){
        const seen = new Set();
        while(d<=end && tops.length<MK_LIMIT){
          const nm = sanitizePath(pre + mkFmtDate(d, fmt));
          if(nm && !seen.has(nm)){ seen.add(nm); tops.push(nm); }
          if(unit==="day") d.setDate(d.getDate()+1);
          else if(unit==="week") d.setDate(d.getDate()+7);
          else d.setMonth(d.getMonth()+1);
        }
        if(d<=end) warn = `한 번에 ${MK_LIMIT}개까지만 만듭니다. 기간을 나눠서 실행해 주세요.`;
      }else if(d>end) warn = "종료일이 시작일보다 빠릅니다.";
    }
  }else if(mode==="combo"){
    const A = mkLines("mkComboA"), B = mkLines("mkComboB");
    const nested = gv("mkComboMode","nested")==="nested";
    const sep = gv("mkComboSep","_");
    A.forEach(a=>{
      if(!B.length){ tops.push(sanitizePath(a)); return; }
      B.forEach(b=>{
        tops.push(sanitizePath(nested ? a+"/"+b : a+sep+b));
      });
    });
    if(tops.length > MK_LIMIT){ tops = tops.slice(0, MK_LIMIT); warn = `한 번에 ${MK_LIMIT}개까지만 만듭니다.`; }
  }else if(mode==="fromfiles"){
    const how = gv("mkFfMode","whole"), delim = gv("mkFfDelim","_");
    const n = Math.max(1, parseInt(gv("mkFfN","4"))||1);
    state.files.forEach(f=>{
      if(f.isDir) return;
      const stem = splitName(f.name, false)[0];
      let key = stem;
      if(how==="delim"){ const i = delim ? stem.indexOf(delim) : -1; key = i>0 ? stem.slice(0,i) : stem; }
      else if(how==="firstN") key = stem.slice(0, n);
      key = sanitizePath(key);
      if(!key) return;
      if(!groups.has(key)){ groups.set(key, []); tops.push(key); }
      groups.get(key).push(f.name);
    });
    if(tops.length > MK_LIMIT){ tops = tops.slice(0, MK_LIMIT); warn = `한 번에 ${MK_LIMIT}개까지만 만듭니다.`; }
  }

  // 같은 이름 중복 제거(대소문자 무시)
  const seen = new Set(), items = [];
  tops.forEach(nm=>{ if(!nm) return; const k=nm.toLowerCase(); if(!seen.has(k)){ seen.add(k); items.push(nm); } });

  const dirsLower = new Set((state.dirs||[]).map(d=>d.toLowerCase()));
  const dirs = [], tree = [], moves = [];
  if(parent) dirs.push(parent);
  const moveOn = mode==="fromfiles" && $("mkFfMove") && $("mkFfMove").checked;
  items.forEach(nm=>{
    const full = parent ? parent+"/"+nm : nm;
    dirs.push(full);
    subs.forEach(sb=>dirs.push(full+"/"+sb));
    const files = groups.get(nm) || [];
    if(moveOn) files.forEach(fn=>moves.push({src:fn, dst:full+"/"+fn}));
    tree.push({name:nm, full, subs, files,
               exists: !parent && nm.indexOf("/")<0 && dirsLower.has(nm.toLowerCase())});
  });
  return {dirs, tree, moves, parent, subs, warn, moveOn};
}

// 만들어질 폴더 구조를 트리로 보여준다 (실행 전 확인용)
function renderMake(){
  const plan = makePlan();
  state.mkPlan = plan;
  const el = $("mkTree"), info = $("mkInfo");
  const runBtn = $("btnRun");
  if(!plan.tree.length && !plan.dirs.length){
    el.innerHTML = '<div class="empty">'
      + (state.mkMode==="fromfiles"
          ? "상단의 <b>폴더 선택</b>으로 폴더를 지정하면, 그 안의 파일 이름으로 만들 폴더가 표시됩니다."
          : "위에 폴더 이름을 입력하면 만들어질 구조가 여기에 표시됩니다.") + '</div>';
    info.textContent = "";
    if(runBtn) runBtn.disabled = true;
    setStatus(state.folder ? "만들 폴더를 입력하세요" : "폴더를 먼저 선택하세요");
    return;
  }
  const rootName = state.folder
    ? (state.folder.replace(/[\\\/]+$/,"").split(/[\\\/]/).pop() || state.folder)
    : "선택한 폴더";
  const LIM = 6;
  let h = `<div class="troot"><span class="ticon">📂</span><span class="tname">${esc(rootName)}</span></div><div class="tchildren">`;
  if(plan.parent) h += `<div class="tfolder"><span class="ticon">📁</span><span class="tname">${esc(plan.parent)}</span><span class="tbadge tnew">상위</span></div>`;
  const shown = plan.tree.slice(0, 60);
  shown.forEach(t=>{
    h += `<div class="tfolder" data-full="${escA(t.full)}" style="${plan.parent?'margin-left:18px':''}">`
       + `<span class="ticon">📁</span><span class="tname">${esc(t.name)}</span>`
       + (t.files.length ? `<span class="tcount">${t.files.length}개</span>` : ``)
       + (t.exists ? `<span class="tbadge texist">이미 있음</span>` : `<span class="tbadge tnew">신규</span>`)
       + `</div>`;
    if(t.subs.length || t.files.length){
      h += `<div class="tfiles" style="${plan.parent?'margin-left:42px':''}">`;
      t.subs.forEach(sb=>{ h += `<div class="tsub"><span class="ticon">📁</span><span>${esc(sb)}</span></div>`; });
      if(plan.moveOn) t.files.slice(0,LIM).forEach(fn=>{
        h += `<div class="tfile"><span class="ticon">📄</span><span>${esc(fn)}</span></div>`; });
      if(plan.moveOn && t.files.length>LIM) h += `<div class="tmore">+${t.files.length-LIM}개 더</div>`;
      h += `</div>`;
    }
  });
  if(plan.tree.length > shown.length) h += `<div class="tmore">+${plan.tree.length - shown.length}개 폴더 더</div>`;
  h += `</div>`;
  el.innerHTML = h;

  const newCount = plan.tree.filter(t=>!t.exists).length;
  const exist = plan.tree.length - newCount;
  let t = `총 ${plan.dirs.length}개 폴더 생성`;
  if(exist) t += ` · 이미 있는 ${exist}개는 건너뜀`;
  if(plan.moves.length) t += ` · 파일 ${plan.moves.length}개 이동`;
  info.textContent = t;

  checkExisting(plan);   // 서버에 물어 '이미 있음' 배지를 정확히 표시
  if(runBtn) runBtn.disabled = !state.folder || plan.dirs.length===0;
  let msg = state.folder ? `만들 폴더 <b class="ok">${plan.dirs.length}</b>개`
                         : `<span class="warn">폴더를 먼저 선택하세요</span>`;
  if(plan.moves.length) msg += ` · 파일 <b>${plan.moves.length}</b>개 이동`;
  if(plan.warn) msg += ` · <span class="warn">${esc(plan.warn)}</span>`;
  setStatus(msg, false, true);
}

// 만들 폴더 중 이미 존재하는 것을 서버에 확인해 배지·요약만 살짝 고쳐 그린다
let mkCheckTimer;
function checkExisting(plan){
  clearTimeout(mkCheckTimer);
  if(!state.folder || !plan.tree.length) return;
  const names = plan.tree.map(t=>t.full);
  mkCheckTimer = setTimeout(async ()=>{
    const res = await api("/api/checkdirs", {folder:state.folder, names});
    if(res.error || !res.exists) return;
    if(state.mkPlan !== plan) return;                  // 그 사이 입력이 바뀌었으면 무시
    const have = new Set(res.exists);
    plan.tree.forEach(t=>{ t.exists = have.has(t.full); });
    document.querySelectorAll("#mkTree .tfolder[data-full]").forEach(el=>{
      const on = have.has(el.dataset.full);
      const b = el.querySelector(".tbadge");
      if(!b || b.textContent==="상위") return;
      b.textContent = on ? "이미 있음" : "신규";
      b.className = "tbadge " + (on ? "texist" : "tnew");
    });
    const exist = plan.tree.filter(t=>t.exists).length;
    let t = `총 ${plan.dirs.length}개 폴더 생성`;
    if(exist) t += ` · 이미 있는 ${exist}개는 건너뜀`;
    if(plan.moves.length) t += ` · 파일 ${plan.moves.length}개 이동`;
    $("mkInfo").textContent = t;
  }, 250);
}

function askMake(){
  const plan = state.mkPlan || makePlan();
  if(!state.folder || !plan.dirs.length) return;
  let body = `<b>${plan.dirs.length}개</b> 폴더를 만듭니다. 이미 있는 폴더는 그대로 둡니다.`;
  if(plan.moves.length) body += `<br>파일 <b>${plan.moves.length}개</b>도 새 폴더로 옮깁니다. (원본 파일은 삭제되지 않고 이동만 합니다)`;
  body += `<br>실행 후에도 <b>실행 취소</b>로 한 번에 되돌릴 수 있습니다.`;
  showDialog("폴더 만들기", body, "만들기", ()=>doMake(plan));
}

async function doMake(plan){
  setStatus("폴더를 만드는 중…");
  const res = await api("/api/mkdirs", {folder:state.folder, names:plan.dirs, moves:plan.moves});
  if(res.error){ setStatus(res.error, true); return; }
  $("btnUndo").disabled = !res.canUndo;
  await reload();                      // 목록 갱신 → recompute → renderMake
  let msg = `<b class="ok">폴더 ${res.made}개 생성 완료</b>`;
  if(res.moved) msg += ` · 파일 ${res.moved}개 이동`;
  if(res.skipped) msg += ` · 이미 있어 건너뜀 ${res.skipped}개`;
  msg += ` · 실행 취소 가능`;
  if(res.errors.length) msg += `<br><span class="warn">실패 ${res.errors.length}건: ${esc(res.errors.slice(0,3).join(" / "))}</span>`;
  msg += ` <button class="gobtn" onclick="openFolder()">탐색기에서 열기</button>`;
  setStatus(msg, false, true);
  addHistory("폴더 만들기",
    `${res.made}개 생성` + (res.moved?` · ${res.moved}개 이동`:"") + (res.skipped?` · ${res.skipped}개 건너뜀`:"")
    + (res.errors.length?` · 실패 ${res.errors.length}건`:""));
}

// ================= 임시저장 · 최근 사용 기억 · 이력 =================
const PREF_KEY = "fileforge.prefs.v2";     // 마지막 폴더·탭 등 (묻지 않고 복원)
const DRAFT_KEY = "fileforge.draft.v2";    // 입력한 규칙·목록 (복원 여부를 물어봄)
const HIST_KEY = "fileforge.history.v2";   // 처리 이력

function lsGet(k){ try{ return JSON.parse(localStorage.getItem(k)||"null"); }catch(e){ return null; } }
function lsSet(k,v){ try{ localStorage.setItem(k, JSON.stringify(v)); }catch(e){} }

function savePrefs(){
  lsSet(PREF_KEY, {folder:state.folder, ext:$("extFilter").value, tab:state.tab,
                   target:state.renameTarget, org:state.org, mkMode:state.mkMode});
}
function collectFields(){
  const o = {};
  document.querySelectorAll("#tab-rename input, #tab-rename select, #tab-make input, #tab-make textarea, #tab-make select")
    .forEach(el=>{ if(el.id) o[el.id] = (el.type==="checkbox") ? el.checked : el.value; });
  return o;
}
function draftSnapshot(){
  return {t:Date.now(),
          rules:[...document.querySelectorAll(".chip[data-rule]")].filter(c=>c.classList.contains("active")).map(c=>c.dataset.rule),
          fields:collectFields(), keywords:state.keywords, mapRules:state.mapRules,
          mapFolders:state.mapFolders,
          perText:state.perText, org:state.org, mkMode:state.mkMode, target:state.renameTarget};
}
function draftHasContent(d){
  if(!d) return false;
  if((d.rules||[]).length) return true;
  if((d.keywords||[]).some(k=>(k||"").trim())) return true;
  if((d.mapRules||[]).some(m=>(m.a||"").trim() || (m.b||"").trim())) return true;
  const f = d.fields || {};
  return ["mkList","mkComboA","mkComboB","mkParent","mkSubs","mkNumBase","mkDateFrom"]
           .some(k=>String(f[k]||"").trim());
}
let saveTimer;
function scheduleSave(){ clearTimeout(saveTimer); saveTimer = setTimeout(doSave, 600); }
function doSave(){
  const d = draftSnapshot();
  lsSet(DRAFT_KEY, d); savePrefs();
  const t = new Date(d.t), p = n=>String(n).padStart(2,"0");
  $("savedAt").textContent = draftHasContent(d)
    ? `자동저장됨 · ${p(t.getHours())}:${p(t.getMinutes())}` : "자동저장 대기 중";
}
function applyDraft(d){
  const f = d.fields || {};
  Object.keys(f).forEach(id=>{
    const el = $(id); if(!el) return;
    if(el.type==="checkbox") el.checked = !!f[id]; else el.value = f[id];
  });
  (d.rules||[]).forEach(r=>{
    const c = document.querySelector('.chip[data-rule="'+r+'"]');
    if(c && !c.classList.contains("active")){ c.classList.add("active"); $("panel-"+r).classList.add("show"); }
  });
  state.keywords = d.keywords || [];
  state.mapRules = d.mapRules || [];
  state.mapFolders = d.mapFolders || [];
  state.perText  = d.perText  || {};
  if(state.keywords.length) renderKwFields();
  if(state.mapRules.length) renderMapFields();
  if(state.mapFolders.length) scanMapFolders();      // 폴더가 지워졌는지도 여기서 확인된다
  if(d.target){ state.renameTarget = d.target; if($("renameTarget")) $("renameTarget").value = d.target; }
  if(d.org){ const c = document.querySelector('#orgChips .chip[data-org="'+d.org+'"]'); if(c) pickOrg(c); }
  if(d.mkMode){ const c = document.querySelector('#mkChips .chip[data-mk="'+d.mkMode+'"]'); if(c) pickMk(c); }
  onRebuildModeChange();
}
function clearDraft(){
  showDialog("입력 초기화",
    "입력해 둔 규칙·목록·텍스트를 모두 비웁니다. <b>파일은 전혀 바뀌지 않습니다.</b>",
    "초기화", ()=>{ try{ localStorage.removeItem(DRAFT_KEY); }catch(e){} location.reload(); });
}

function addHistory(kind, summary, folder){
  const h = lsGet(HIST_KEY) || [];
  h.unshift({t:Date.now(), kind, summary, folder:folder || state.folder});
  lsSet(HIST_KEY, h.slice(0,30));
  renderHistory();
}
function renderHistory(){
  const h = lsGet(HIST_KEY) || [];
  const box = $("histRows"); if(!box) return;
  if(!h.length){ box.innerHTML = '<div class="hint" style="padding:10px 0">아직 실행한 작업이 없습니다.</div>'; return; }
  const p = n=>String(n).padStart(2,"0");
  box.innerHTML = h.map((r,i)=>{
    const d = new Date(r.t);
    return `<div class="hrow"><span class="htime">${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}</span>`
      + `<span class="hkind">${esc(r.kind)}</span>`
      + `<span>${esc(r.summary)}<br><span class="hint">${esc(r.folder||"")}</span></span>`
      + (r.folder ? `<button class="gobtn" onclick="openHist(${i})" title="${escA(r.folder)}">탐색기에서 열기</button>`
                  : `<span></span>`)
      + `</div>`;
  }).join("");
}

// ================= 단축키 =================
document.addEventListener("keydown", e=>{
  const dlgOpen = $("scrim").classList.contains("show");
  if(e.key === "Escape"){ if(dlgOpen) closeDialog(); return; }
  if((e.ctrlKey || e.metaKey) && e.key === "Enter"){
    e.preventDefault();
    if(dlgOpen){ $("dlgOk").click(); return; }
    if(state.tab === "pdf"){ if(!$("pdfRun").disabled) runPdfOp(); }
    else if(!$("btnRun").disabled) askExecute();
    return;
  }
  if((e.ctrlKey || e.metaKey) && (e.key === "z" || e.key === "Z") && !e.shiftKey){
    const t = e.target;
    if(t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA")) return;  // 입력 중엔 글자 되돌리기
    if(dlgOpen) return;
    e.preventDefault();
    if(!$("btnUndo").disabled) doUndo();
  }
});

// ================= 시작 =================
async function init(){
  renderHistory();
  const pref = lsGet(PREF_KEY), draft = lsGet(DRAFT_KEY);
  if(pref){
    if(pref.ext) $("extFilter").value = pref.ext;
    if(pref.target){ state.renameTarget = pref.target; $("renameTarget").value = pref.target; }
  }
  document.addEventListener("input", scheduleSave);
  document.addEventListener("change", scheduleSave);

  const startup = async ()=>{
    onTargetChange();                                  // 확장자 필터 활성/비활성 반영
    if(pref && pref.folder){ state.folder = pref.folder; await reload(true); }
    if(pref && pref.tab && pref.tab !== "rename") await switchTab(pref.tab);
    else recompute();
    doSave();
  };

  if(draftHasContent(draft)){
    showDialog("이전 작업 내용이 있습니다",
      "마지막으로 입력했던 규칙·목록이 저장되어 있습니다. 이어서 하시겠습니까?",
      "이어서 하기", async ()=>{ applyDraft(draft); await startup(); },
      "새로 시작", async ()=>{ try{ localStorage.removeItem(DRAFT_KEY); }catch(e){} await startup(); });
  }else{
    await startup();
  }
}
window.addEventListener("DOMContentLoaded", init);

// ================= PDF · 변환 도구 =================
const PDF_EXTS = {
  merge:["pdf"], split:["pdf"], compress:["pdf"], pdf2img:["pdf"],
  img2pdf:["jpg","jpeg","png","bmp","tif","tiff","gif","webp"],
  img2img:["jpg","jpeg","png","bmp","gif","tif","tiff","webp","heic","heif","pdf"],
  excel2pdf:["xlsx","xls","xlsm","ods","csv"],
  hwp2pdf:["hwp","hwpx"],
  xlsmerge:["xlsx","xlsm"], xlsappend:["xlsx","xlsm"],
};
const PDF_EMPTY = {
  merge:"이 폴더에 PDF가 없습니다.", split:"이 폴더에 PDF가 없습니다.",
  compress:"이 폴더에 PDF가 없습니다.", pdf2img:"이 폴더에 PDF가 없습니다.",
  img2pdf:"이 폴더에 이미지가 없습니다.", img2img:"이 폴더에 이미지·PDF가 없습니다.",
  excel2pdf:"이 폴더에 엑셀 파일이 없습니다.",
  hwp2pdf:"이 폴더에 한글 파일(.hwp/.hwpx)이 없습니다.",
  xlsmerge:"이 폴더에 엑셀 파일(.xlsx/.xlsm)이 없습니다.",
  xlsappend:"이 폴더에 엑셀 파일(.xlsx/.xlsm)이 없습니다.",
};
function gv(id, d){ const e=$(id); return e ? e.value : d; }

function pickPdfOp(chip){
  document.querySelectorAll("#pdfOps .chip").forEach(c=>c.classList.remove("active"));
  chip.classList.add("active");
  state.pdfOp = chip.dataset.op;
  setPdfResult("");
  renderPdfOptions();
  loadPdfFiles();
}

function renderPdfOptions(){
  const op = state.pdfOp;
  const orderSel = `<div class="row"><span class="hint">정렬</span>
    <select class="field" id="pdfOrder"><option value="asc">파일명 오름차순</option>
    <option value="desc">파일명 내림차순</option></select></div>`;
  const nameIn = (ph)=>`<div class="row"><span class="hint">저장 이름</span>
    <input class="field" id="pdfOut" placeholder="${ph} (비우면 자동)"></div>`;
  let h = "";
  if(op==="merge")      h = `<p class="hint">선택한 PDF들을 정렬 순서대로 하나로 합칩니다 (2개 이상).<br>각 파일 오른쪽 칸에 <b>1-5, 7-8</b> 처럼 입력하면 그 페이지만 병합합니다. (비우면 전체)</p>` + orderSel + nameIn("병합");
  else if(op==="img2pdf") h = `<p class="hint">선택한 이미지들을 정렬 순서대로 한 PDF로 묶습니다.</p>` + orderSel + nameIn("이미지병합");
  else if(op==="split") h = `<p class="hint">목록에서 PDF <b>1개</b>를 고른 뒤, 아래 두 기능 중 하나를 사용하세요.</p>
    <div class="row"><span class="hint">① 분할 기준</span>
    <input class="field narrow" id="pdfPagesPer" type="number" min="1" value="10">
    <span class="hint">이 페이지 수마다 잘라 새 폴더에 저장</span></div>
    <div class="row"><span class="hint">② 페이지 삭제</span>
    <input class="field" id="pdfDelPages" placeholder="예: 1-3, 5, 8-10" style="width:240px">
    <span class="hint">입력한 페이지·구간을 삭제한 새 PDF로 저장 (입력하면 분할 대신 삭제 실행)</span></div>`;
  else if(op==="pdf2img") h = `<p class="hint">선택한 PDF의 각 페이지를 이미지로 저장합니다. <b>파일명은 그대로 유지</b>되고, 여러 페이지면 <b>이름_1, 이름_2…</b> 로 저장됩니다. (원본 PDF는 그대로)</p>
    <div class="row"><span class="hint">형식</span>
    <select class="field" id="pdfImgFmt"><option value="jpg">JPG</option><option value="jpeg">JPEG (.jpeg)</option><option value="png">PNG</option></select>
    <span class="hint">해상도</span>
    <select class="field" id="pdfImgDpi"><option value="100">낮음 (100dpi · 작은 용량)</option>
    <option value="150" selected>보통 (150dpi)</option><option value="300">높음 (300dpi · 인쇄용)</option></select></div>`;
  else if(op==="img2img") h = `<p class="hint">선택한 <b>이미지·PDF</b>(PNG·BMP·GIF·TIFF·WEBP·JPG·<b>HEIC</b>·<b>PDF</b> 등)를 고른 형식으로 <b>한 번에 변환</b>합니다. <b>파일명은 그대로 유지</b>되고, PDF는 페이지별로 <b>이름_1, 이름_2…</b> 로 저장됩니다. (원본은 그대로 · 투명 배경은 흰색으로 채워짐)</p>
    <div class="row"><span class="hint">변환할 형식</span>
    <select class="field" id="imgFmt"><option value="jpg">JPG</option><option value="jpeg">JPEG (.jpeg)</option><option value="png">PNG</option></select>
    <span class="hint">PDF 해상도</span>
    <select class="field" id="imgDpi"><option value="100">낮음 (100dpi)</option>
    <option value="150" selected>보통 (150dpi)</option><option value="300">높음 (300dpi)</option></select></div>`;
  else if(op==="compress") h = `<p class="hint">선택한 PDF의 용량을 줄여 <b>_압축</b> 파일로 저장합니다. (원본 유지)</p>
    <div class="row"><span class="hint">압축 강도</span>
    <select class="field" id="pdfLevel"><option value="high">강 — 가장 작게</option>
    <option value="medium" selected>중 — 균형</option><option value="low">약 — 화질 우선</option></select></div>`;
  else if(op==="excel2pdf") h = `<p class="hint">선택한 엑셀 파일을 각각 PDF로 변환합니다. (LibreOffice 필요 · 서식 유지)</p>`;
  else if(op==="hwp2pdf") h = `<p class="hint">선택한 한글 문서(.hwp/.hwpx)를 각각 PDF로 변환합니다. 원본은 그대로 두고 같은 이름의 PDF를 만듭니다.<br>
    한글(아래아한글)이 설치돼 있으면 자동으로 사용해 서식을 그대로 재현하고, 없으면 LibreOffice로 변환합니다.</p>`;
  else if(op==="xlsmerge") h = `<p class="hint">선택한 엑셀 파일들의 <b>모든 시트</b>를 하나의 새 워크북으로 모읍니다. (원본 유지 · .xlsx/.xlsm)</p>`
    + `<div class="row"><span class="hint">시트 순서</span>
       <select class="field" id="pdfOrder">
         <option value="asc">시트 이름 오름차순</option>
         <option value="desc">시트 이름 내림차순</option>
         <option value="file">파일·시트 원래 순서</option>
       </select></div>` + nameIn("시트합치기");
  else if(op==="xlsappend") h = `<p class="hint"><b>먼저 클릭한 파일이 1번(헤더 기준)</b>입니다. 선택한 모든 파일의 모든 시트의 행을 위에서 아래로 이어 붙여 하나의 시트로 통합합니다. (원본 유지 · .xlsx/.xlsm)</p>`
    + `<div class="row"><label class="opt"><input type="checkbox" id="pdfSkipHeader" checked>
       각 시트의 첫 행을 헤더로 보고 제외 (헤더는 1번 파일 것만 사용)</label></div>` + nameIn("행통합");
  $("pdfOptions").innerHTML = h;
}

async function loadPdfFiles(){
  state.pdfSelected = new Set();
  state.pdfRanges = {};
  if(!state.folder){
    $("pdfList").innerHTML = '<div class="empty">상단의 <b>폴더 선택</b>으로 폴더를 지정하세요.</div>';
    $("pdfListInfo").textContent = ""; updatePdfRun(); return;
  }
  const res = await api("/api/list", {folder:state.folder, exts:PDF_EXTS[state.pdfOp]});
  if(res.error){ $("pdfList").innerHTML = '<div class="empty">'+esc(res.error)+'</div>'; return; }
  state.pdfFiles = res.files;
  renderPdfList();
}

function renderPdfList(){
  const single = state.pdfOp==="split";
  $("pdfAllWrap").style.display = single ? "none" : "";
  if(!state.pdfFiles.length){
    $("pdfList").innerHTML = '<div class="empty">'+PDF_EMPTY[state.pdfOp]+'</div>';
    $("pdfListInfo").textContent = ""; updatePdfRun(); return;
  }
  let h = "";
  state.pdfFiles.forEach((f,i)=>{
    const on = state.pdfSelected.has(f.name);
    let rangeUi = "";
    if(state.pdfOp==="merge" && on){
      const v = state.pdfRanges[f.name] || "";
      rangeUi = `<input class="field pdfrange" type="text" value="${esc(v)}"
        placeholder="예: 1-5, 7-8 (전체)"
        onclick="event.stopPropagation()" oninput="setPdfRange(${i}, this.value)">`;
    }
    h += `<div class="pdfrow ${on?'sel':''}" onclick="togglePdfRow(${i})">
      <span class="pdfdot"></span><span class="pname">${esc(f.name)}</span>${rangeUi}</div>`;
  });
  $("pdfList").innerHTML = h;
  $("pdfListInfo").textContent = `${state.pdfFiles.length}개 중 ${state.pdfSelected.size}개 선택`;
  updatePdfRun();
}

function togglePdfRow(i){
  const name = state.pdfFiles[i].name;
  if(state.pdfOp==="split"){
    state.pdfSelected = state.pdfSelected.has(name) ? new Set() : new Set([name]);
  }else{
    if(state.pdfSelected.has(name)) state.pdfSelected.delete(name);
    else state.pdfSelected.add(name);
  }
  renderPdfList();
}

function pdfSelectAll(){
  if(state.pdfOp==="split") return;   // 분할은 1개만 선택 가능
  state.pdfSelected = new Set(state.pdfFiles.map(f=>f.name));
  renderPdfList();
}
function pdfDeselectAll(){
  state.pdfSelected = new Set();
  renderPdfList();
}

function setPdfRange(i, val){
  state.pdfRanges[state.pdfFiles[i].name] = val;
}

function updatePdfRun(){
  const n = state.pdfSelected.size;
  const ok = state.pdfOp==="merge" ? n>=2 : n>=1;
  $("pdfRun").disabled = !ok;
}

function setPdfResult(msg, kind){
  const el = $("pdfResult");
  el.className = kind || "";
  el.innerHTML = msg;
}

async function runPdfOp(){
  const op = state.pdfOp, files = [...state.pdfSelected];
  $("pdfRun").disabled = true;
  setPdfResult("처리 중… (PDF 기능 첫 사용 시 라이브러리 설치로 시간이 걸릴 수 있습니다)");
  let res;
  if(op==="merge"){
    const ranges = {};
    files.forEach(n=>{ const r=(state.pdfRanges[n]||"").trim(); if(r) ranges[n]=r; });
    res = await api("/api/merge", {folder:state.folder, files, order:gv("pdfOrder","asc"), out:gv("pdfOut",""), ranges});
  }
  else if(op==="split"){
    const del = (gv("pdfDelPages","")||"").trim();
    if(del) res = await api("/api/pdfdelete", {folder:state.folder, file:files[0], pages:del});
    else    res = await api("/api/split",     {folder:state.folder, file:files[0], pagesPer:gv("pdfPagesPer","10")});
  }
  else if(op==="img2pdf") res = await api("/api/img2pdf", {folder:state.folder, files, order:gv("pdfOrder","asc"), out:gv("pdfOut","")});
  else if(op==="pdf2img") res = await api("/api/pdf2img", {folder:state.folder, files, fmt:gv("pdfImgFmt","jpg"), dpi:gv("pdfImgDpi","150")});
  else if(op==="img2img") res = await api("/api/img2img", {folder:state.folder, files, fmt:gv("imgFmt","jpg"), dpi:gv("imgDpi","150")});
  else if(op==="compress")res = await api("/api/compress",{folder:state.folder, files, level:gv("pdfLevel","medium")});
  else if(op==="excel2pdf")res= await api("/api/excel2pdf",{folder:state.folder, files});
  else if(op==="hwp2pdf")  res= await api("/api/hwp2pdf",  {folder:state.folder, files});
  else if(op==="xlsmerge") res= await api("/api/xlsmerge", {folder:state.folder, files, order:gv("pdfOrder","asc"), out:gv("pdfOut","")});
  else if(op==="xlsappend")res= await api("/api/xlsappend",{folder:state.folder, files, out:gv("pdfOut",""), skipHeader:($("pdfSkipHeader")?$("pdfSkipHeader").checked:true)});

  if(res.error){ setPdfResult("⚠ "+esc(res.error).replace(/\n/g,"<br>"), "warn"); updatePdfRun(); return; }
  setPdfResult(formatPdfResult(op, res), "ok");
  await loadPdfFiles();
}

function kb(n){ return n>=1048576 ? (n/1048576).toFixed(1)+"MB" : Math.round(n/1024)+"KB"; }

function formatPdfResult(op, res){
  if(op==="merge")   return `✓ PDF ${res.count}개(총 ${res.pages}페이지)를 합쳤습니다 → <b>${esc(res.output)}</b>`;
  if(op==="img2pdf") return `✓ 이미지 ${res.count}장을 묶었습니다 → <b>${esc(res.output)}</b>`;
  if(op==="pdf2img"){
    const ok = res.results.filter(r=>!r.error);
    const fail = res.results.filter(r=>r.error);
    let m = `✓ PDF ${ok.length}개 → ${res.fmt.toUpperCase()} 이미지 ${res.total}장 저장`
      + (ok.length? "<br>"+ok.map(r=>`${esc(r.name)} → ${r.count}장`).join("<br>") : "");
    if(fail.length) m += `<br><span class="warn">실패 ${fail.length}건: ${fail.map(r=>`${esc(r.name)}${r.error?" ("+esc(r.error)+")":""}`).join(", ")}</span>`;
    return m;
  }
  if(op==="img2img"){
    const ok = res.results.filter(r=>!r.error);
    const fail = res.results.filter(r=>r.error);
    let m = `✓ ${ok.length}개를 ${res.fmt.toUpperCase()}로 변환했습니다`
      + (ok.length? "<br>"+ok.map(r=>`${esc(r.name)} → ${esc(r.output)}`+((r.count>1)?` 외 ${r.count-1}장`:``)).join("<br>") : "");
    if(fail.length) m += `<br><span class="warn">실패 ${fail.length}건: ${fail.map(r=>`${esc(r.name)}${r.error?" ("+esc(r.error)+")":""}`).join(", ")}</span>`;
    return m;
  }
  if(op==="split"){
    if(res.removed!==undefined)
      return `✓ ${res.removed}개 페이지를 삭제했습니다 (남은 ${res.remaining}p / 원본 ${res.total}p) → <b>${esc(res.output)}</b>`;
    return `✓ ${res.parts}개로 분할했습니다 (총 ${res.total_pages}p) → <b>${esc(res.folder)}/</b> 폴더`;
  }
  if(op==="compress"){
    const ok = res.results.filter(r=>!r.error);
    const parts = ok.map(r=>`${esc(r.name)} ${kb(r.before)}→${kb(r.after)} (${r.saved}%↓)`);
    const fail = res.results.filter(r=>r.error).map(r=>esc(r.name));
    let m = `✓ ${ok.length}개 압축 완료 [${res.engine}]<br>` + parts.join("<br>");
    if(res.engine==="pypdf") m += `<br><span class="hint">※ 더 줄이려면 Ghostscript 설치 권장</span>`;
    if(fail.length) m += `<br>실패: ${fail.join(", ")}`;
    return m;
  }
  if(op==="excel2pdf"){
    const ok = res.results.filter(r=>!r.error);
    const fail = res.results.filter(r=>r.error);
    let m = `✓ ${ok.length}개 변환 완료` + (ok.length? "<br>"+ok.map(r=>esc(r.output)).join("<br>") : "");
    if(fail.length) m += `<br><span class="warn">실패 ${fail.length}건: ${fail.map(r=>esc(r.name)).join(", ")}</span>`;
    return m;
  }
  if(op==="hwp2pdf"){
    const ok = res.results.filter(r=>!r.error);
    const fail = res.results.filter(r=>r.error);
    const eng = res.engine==="hancom" ? "한글" : "LibreOffice";
    let m = `✓ ${ok.length}개 변환 완료 [${eng}]` + (ok.length? "<br>"+ok.map(r=>esc(r.output)).join("<br>") : "");
    if(fail.length) m += `<br><span class="warn">실패 ${fail.length}건: ${fail.map(r=>`${esc(r.name)}${r.error?" ("+esc(r.error)+")":""}`).join(", ")}</span>`;
    return m;
  }
  if(op==="xlsmerge") return `✓ ${res.files}개 파일의 시트 ${res.sheets}개를 합쳤습니다 → <b>${esc(res.output)}</b>`;
  if(op==="xlsappend") return `✓ ${res.files}개 파일에서 데이터 ${res.rows}행을 통합했습니다 → <b>${esc(res.output)}</b>`
    + (res.header ? "" : `<br><span class="hint">※ 헤더 없이 모든 행을 합쳤습니다</span>`);
  return "완료";
}

// ── UI 수명 주기: 탭을 닫으면 백그라운드 서버도 함께 종료 ──
const _TAB_ID = (window.crypto && crypto.randomUUID)
  ? crypto.randomUUID() : (Date.now() + "_" + Math.random());
function _heartbeat(){
  fetch("/api/ping", {method:"POST", headers:{"Content-Type":"application/json"},
                      body:JSON.stringify({tab:_TAB_ID}), keepalive:true}).catch(()=>{});
}
_heartbeat();
// TTL 이 70초라 15초면 충분하다. 탭을 닫을 때는 아래 sendBeacon 으로 즉시 알린다.
setInterval(_heartbeat, 15000);
function _bye(){
  try{ navigator.sendBeacon("/api/quit", JSON.stringify({tab:_TAB_ID})); }catch(e){}
}
window.addEventListener("pagehide", _bye);
window.addEventListener("beforeunload", _bye);
</script>
</body>
</html>
"""


# ---------------- HTTP 서버 ----------------

PAGE_BYTES = PAGE.encode("utf-8")   # 요청마다 90KB를 다시 인코딩하지 않도록 한 번만 변환


class Handler(BaseHTTPRequestHandler):
    # HTTP/1.1 = keep-alive. 예전에는 하트비트마다 TCP 연결을 새로 맺어
    # TIME_WAIT 소켓이 수십 개씩 쌓였다. (모든 응답에 Content-Length 를 붙이고 있음)
    protocol_version = "HTTP/1.1"
    timeout = 30                    # 유휴 연결이 워커 스레드를 계속 붙잡지 않도록

    def log_message(self, *args):
        pass  # 콘솔 로그 생략

    def _json(self, obj, code=200):
        data = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path == "/api/hello":   # 중복 실행 감지용 서명
            return self._json({"app": "fileforge"})
        if self.path in ("/", "/index.html"):
            data = PAGE_BYTES
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        else:
            self.send_error(404)

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length) or b"{}")
        except (ValueError, json.JSONDecodeError):
            return self._json({"error": "잘못된 요청"}, 400)

        # ----- UI 수명 주기: 하트비트 / 닫힘 신호 (파일 작업 LOCK 밖에서 즉시 처리) -----
        if self.path == "/api/ping":
            global _ARMED
            tid = body.get("tab", "")
            with TABS_LOCK:
                _ARMED = True
                if tid:
                    ACTIVE_TABS[tid] = time.time() + HEARTBEAT_TTL
            return self._json({"ok": True})

        if self.path == "/api/quit":
            tid = body.get("tab", "")
            with TABS_LOCK:
                if tid in ACTIVE_TABS:  # 즉시 삭제 대신 유예: 새로고침이면 재연결로 되살아남
                    ACTIVE_TABS[tid] = min(ACTIVE_TABS[tid], time.time() + CLOSE_GRACE)
            return self._json({"ok": True})

        try:
            with LOCK:
                if self.path == "/api/pick":
                    path = pick_folder_native()
                    return self._json({"path": path})

                if self.path == "/api/pick_many":
                    return self._json({"paths": pick_folders_native()})

                if self.path == "/api/map_scan":
                    return self._json({"folders": map_scan(body.get("folders", []),
                                                           body.get("rules", []))})

                if self.path == "/api/list":
                    folder = body.get("folder", "")
                    if not os.path.isdir(folder):
                        return self._json({"error": "폴더를 찾을 수 없습니다: " + folder})
                    files, all_names = list_files(folder, set(body.get("exts", [])),
                                                  body.get("include", "files"))
                    dirs = [n for n in all_names
                            if os.path.isdir(os.path.join(folder, n))]
                    return self._json({"files": files, "allNames": all_names, "dirs": dirs})

                if self.path == "/api/execute":
                    folder = body.get("folder", "")
                    folders = body.get("folders") or None      # 여러 폴더에 한꺼번에 적용할 때
                    if folders:
                        # 실제로 작업이 있는 폴더만 확인한다. 목록에 지워진 폴더가
                        # 섞여 있어도 나머지 폴더 작업까지 막지 않기 위함이다.
                        bad = None
                        for j in body.get("jobs", []):
                            try:
                                base = folders[int(j.get("d", 0))]
                            except (TypeError, ValueError, IndexError, AttributeError):
                                bad = "폴더 지정이 잘못되었습니다."
                                break
                            if not os.path.isdir(base):
                                bad = "폴더를 찾을 수 없습니다: " + base
                                break
                        if bad:
                            return self._json({"error": bad})
                    elif not os.path.isdir(folder):
                        return self._json({"error": "폴더를 찾을 수 없습니다"})
                    done, errors = execute_jobs(folder, body.get("jobs", []),
                                                body.get("mode", "rename"), folders)
                    return self._json({"done": done, "errors": errors,
                                       "canUndo": bool(UNDO_STACK)})

                if self.path == "/api/undo":
                    restored, errors, removed = undo_last()
                    return self._json({"restored": restored, "errors": errors,
                                       "removedDirs": removed,
                                       "canUndo": bool(UNDO_STACK)})

                if self.path == "/api/mkdirs":
                    folder = body.get("folder", "")
                    if not os.path.isdir(folder):
                        return self._json({"error": "폴더를 먼저 선택하세요."})
                    res = make_dirs(folder, body.get("names", []), body.get("moves", []))
                    res["canUndo"] = bool(UNDO_STACK)
                    return self._json(res)

                if self.path == "/api/checkdirs":
                    # 만들 예정인 경로 중 '이미 있는 폴더'를 알려준다 (미리보기 배지용)
                    folder = body.get("folder", "")
                    if not os.path.isdir(folder):
                        return self._json({"exists": []})
                    found = []
                    for rel in (body.get("names", []) or [])[:2000]:
                        try:
                            p = safe_join(folder, (rel or "").strip().strip("/\\"))
                        except ValueError:
                            continue
                        if os.path.isdir(p):
                            found.append(rel)
                    return self._json({"exists": found})

                if self.path == "/api/open":
                    open_in_explorer(body.get("folder", ""))
                    return self._json({"ok": True})

                # ----- PDF · 변환 도구 -----
                folder = body.get("folder", "")
                if self.path in ("/api/merge", "/api/split", "/api/pdfdelete", "/api/img2pdf",
                                 "/api/pdf2img", "/api/img2img", "/api/compress", "/api/excel2pdf",
                                 "/api/hwp2pdf", "/api/xlsmerge", "/api/xlsappend"):
                    if not os.path.isdir(folder):
                        return self._json({"error": "폴더를 먼저 선택하세요."})

                pdf_result = None
                if self.path == "/api/merge":
                    pdf_result = pdf_merge(folder, body.get("files", []),
                                           body.get("order", "asc"), body.get("out", ""),
                                           body.get("ranges", {}))
                elif self.path == "/api/split":
                    pdf_result = pdf_split(folder, body.get("file", ""),
                                           body.get("pagesPer", 1))
                elif self.path == "/api/pdfdelete":
                    pdf_result = pdf_delete_pages(folder, body.get("file", ""),
                                                  body.get("pages", ""))
                elif self.path == "/api/img2pdf":
                    pdf_result = images_to_pdf(folder, body.get("files", []),
                                               body.get("order", "asc"), body.get("out", ""))
                elif self.path == "/api/pdf2img":
                    pdf_result = pdf_to_images(folder, body.get("files", []),
                                               body.get("fmt", "jpg"), body.get("dpi", 150))
                elif self.path == "/api/img2img":
                    pdf_result = images_convert(folder, body.get("files", []),
                                                body.get("fmt", "jpg"), body.get("dpi", 150))
                elif self.path == "/api/compress":
                    pdf_result = compress_pdf(folder, body.get("files", []),
                                              body.get("level", "medium"))
                elif self.path == "/api/excel2pdf":
                    pdf_result = excel_to_pdf(folder, body.get("files", []))
                elif self.path == "/api/hwp2pdf":
                    pdf_result = hwp_to_pdf(folder, body.get("files", []))
                elif self.path == "/api/xlsmerge":
                    pdf_result = excel_merge_sheets(folder, body.get("files", []),
                                                    body.get("order", "asc"), body.get("out", ""))
                elif self.path == "/api/xlsappend":
                    pdf_result = excel_consolidate_rows(folder, body.get("files", []),
                                                        body.get("out", ""), body.get("skipHeader", True))

                if pdf_result is not None:
                    notify_shell_change([folder])   # 새로 만든 파일이 탐색기에 바로 보이도록
                    return self._json(pdf_result)

            self.send_error(404)
        except (ValueError, RuntimeError, OSError) as e:
            self._json({"error": str(e)}, 200)
        except Exception as e:  # PDF/변환 라이브러리 오류도 UI로 안전하게 전달
            self._json({"error": f"처리 중 오류: {e}"}, 200)


class _Server(ThreadingHTTPServer):
    """중복 바인딩을 막는 HTTP 서버.

    윈도우의 SO_REUSEADDR 는 리눅스와 달리 "이미 다른 프로세스가 듣고 있는
    포트에 또 바인딩하는 것"까지 허용한다. HTTPServer 는 이 옵션을 기본으로
    켜기 때문에, FileForge 를 두 번 실행하면 두 인스턴스가 같은 8765 포트에
    동시에 붙어 요청이 어느 쪽으로 갈지 알 수 없는 상태가 됐다."""
    daemon_threads = True      # 요청 처리 스레드가 종료를 막지 않도록
    allow_reuse_address = False
    exclusive = True

    def server_bind(self):
        if self.exclusive and hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            try:
                self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
            except OSError:
                pass
        super().server_bind()


class _ReuseServer(_Server):
    """직전 인스턴스가 남긴 TIME_WAIT 소켓 때문에 바인딩이 막혔을 때만 쓰는 완화 버전.
    (듣고 있는 프로세스가 없다고 먼저 확인한 뒤에만 사용한다)"""
    allow_reuse_address = True
    exclusive = False


def _port_listening(port):
    """이 포트에서 실제로 듣고 있는 프로세스가 있는지 확인."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sk:
        sk.settimeout(0.3)
        return sk.connect_ex((HOST, port)) == 0


def _is_fileforge(port):
    """그 포트를 쓰는 게 이미 실행 중인 FileForge 인지 확인."""
    conn = None
    try:
        conn = http.client.HTTPConnection(HOST, port, timeout=0.5)
        conn.request("GET", "/api/hello")
        res = conn.getresponse()
        return res.status == 200 and json.loads(res.read(200) or b"{}").get("app") == "fileforge"
    except Exception:
        return False
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


def open_server():
    """비어 있는 포트에 서버를 연다.

    반환: (서버, 포트). 이미 FileForge 가 떠 있으면 (None, 그 포트) — 새로 띄우지 않고
    기존 화면만 열어주기 위함이다."""
    last_err = None
    for port in range(BASE_PORT, BASE_PORT + 20):
        if _port_listening(port):
            if _is_fileforge(port):
                return None, port      # 이미 실행 중 → 그 창을 열어준다
            continue                   # 다른 프로그램이 쓰는 포트 → 건너뛴다
        try:
            return _Server((HOST, port), Handler), port
        except OSError as e:
            last_err = e
            try:
                # 듣는 프로세스는 없는데 막혔다면 직전 인스턴스의 TIME_WAIT 소켓 탓이다
                return _ReuseServer((HOST, port), Handler), port
            except OSError as e2:
                last_err = e2
    raise SystemExit(f"사용 가능한 포트를 찾지 못했습니다. ({last_err})")


def _shutdown_now():
    """서버 루프를 멈춰 프로세스 종료 절차를 시작한다."""
    global _SHUTTING
    with TABS_LOCK:
        if _SHUTTING:
            return
        _SHUTTING = True
    STOP.set()
    srv = SERVER
    if srv is not None:
        # serve_forever() 루프를 다른 스레드에서 멈춘다(같은 스레드에서 호출 불가)
        threading.Thread(target=srv.shutdown, daemon=True).start()


def _maybe_shutdown():
    """열린 UI(탭)가 하나도 없으면 서버를 정상 종료시킨다."""
    with TABS_LOCK:
        if _SHUTTING or not _ARMED or ACTIVE_TABS:
            return
    _shutdown_now()


def _watchdog():
    """만료된 탭을 걷어내고, 남은 탭이 없으면 종료를 트리거하는 데몬 스레드."""
    while not STOP.wait(1.0):
        now = time.time()
        with TABS_LOCK:
            for tid in [t for t, exp in ACTIVE_TABS.items() if exp <= now]:
                del ACTIVE_TABS[tid]
        # 브라우저가 한 번도 열리지 않으면 종료 조건(_ARMED)이 영영 안 걸려서
        # 프로세스가 계속 남는다. 그 경우도 스스로 정리하게 한다.
        if not _ARMED and now - _START_TS > ARM_TIMEOUT:
            _shutdown_now()
            return
        _maybe_shutdown()


def main():
    global SERVER, _START_TS
    # frozen exe 특수 모드: 폴더 다이얼로그만 띄우고 결과를 파일로 남긴 뒤 종료
    pick_out = os.environ.get("FILEFORGE_PICK_OUT")
    if pick_out:
        try:
            picked = _show_folder_dialog()
        except Exception as e:  # 진단용: 실패 원인을 결과 파일에 남긴다
            picked = "__ERROR__ " + repr(e)
        try:
            with open(pick_out, "w", encoding="utf-8") as f:
                f.write(picked)
        except OSError:
            pass
        return

    server, port = open_server()
    url = f"http://{HOST}:{port}/"

    if server is None:
        # 이미 FileForge 가 떠 있다 → 두 번째 서버를 띄우지 않고 그 화면만 열어준다
        webbrowser.open(url)
        return

    SERVER = server
    _START_TS = time.time()

    if sys.stdout:  # .pyw(콘솔 없음)로 실행하면 stdout이 없음
        print(f"FileForge 실행 중 → {url}")
        print("브라우저 탭을 닫으면 자동으로 종료됩니다. (Ctrl+C 로도 종료)")

    threading.Thread(target=_watchdog, daemon=True).start()          # UI 종료 감시
    threading.Thread(target=server.serve_forever, daemon=True).start()
    opener = threading.Timer(0.4, lambda: webbrowser.open(url))
    opener.daemon = True
    opener.start()
    try:
        # 메인 스레드는 폴더 선택 다이얼로그 전담. STOP 이 걸리면 루프가 끝난다.
        _dialog_dispatcher()
    except KeyboardInterrupt:
        pass
    finally:
        # 소켓·서버 자원 해제 후, 남은 스레드까지 확실히 정리하고 프로세스 종료
        try:
            server.shutdown()
        except Exception:
            pass
        try:
            server.server_close()
        except Exception:
            pass
        os._exit(0)


if __name__ == "__main__":
    main()
