"""릴리스 안내문 만들기 — 템플릿의 __SIZE__ / __MSG__ / __SHA__ 자리를 실제 값으로 채운다."""
import io
import os

here = os.path.dirname(os.path.abspath(__file__))
text = io.open(os.path.join(here, "release-notes-template.md"), encoding="utf-8").read()

msg = (os.environ.get("COMMIT_MSG") or "").strip().splitlines()
text = (text.replace("__SIZE__", os.environ.get("SIZE_MB", "?"))
            .replace("__MSG__", msg[0] if msg else "수동 빌드")
            .replace("__SHA__", os.environ.get("SHA256", "?")))

io.open(1, "w", encoding="utf-8", closefd=False).write(text)
