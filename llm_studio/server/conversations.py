"""conversations.py

대화 기록을 conversations/ 폴더에 대화당 JSON 파일 하나로 저장한다.
메시지는 OpenAI 규격 그대로 저장하므로(assistant.tool_calls, role=tool 포함)
불러온 이력을 그대로 다음 요청에 넣을 수 있다.
"""

from __future__ import annotations

import json
import time
import uuid
from pathlib import Path


class ConversationStore:
    def __init__(self, base_dir: Path):
        self.dir = base_dir / "conversations"
        self.dir.mkdir(exist_ok=True)

    def _path(self, conv_id: str) -> Path:
        safe = "".join(ch for ch in conv_id if ch.isalnum() or ch in "-_")
        return self.dir / f"{safe}.json"

    def create(self, title: str = "새 대화") -> dict:
        conv = {
            "id": f"{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}",
            "title": title,
            "created": time.time(),
            "updated": time.time(),
            "messages": [],
        }
        self.save(conv)
        return conv

    def save(self, conv: dict) -> None:
        conv["updated"] = time.time()
        self._path(conv["id"]).write_text(
            json.dumps(conv, ensure_ascii=False, indent=1), encoding="utf-8"
        )

    def load(self, conv_id: str) -> dict | None:
        path = self._path(conv_id)
        if not path.exists():
            return None
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None

    def delete(self, conv_id: str) -> bool:
        path = self._path(conv_id)
        if path.exists():
            path.unlink()
            return True
        return False

    def rename(self, conv_id: str, title: str) -> bool:
        conv = self.load(conv_id)
        if conv is None:
            return False
        conv["title"] = title.strip()[:80] or conv["title"]
        self.save(conv)
        return True

    def list(self) -> list[dict]:
        """메시지 본문을 뺀 메타 목록. 최근 수정 순."""
        metas = []
        for path in self.dir.glob("*.json"):
            try:
                conv = json.loads(path.read_text(encoding="utf-8"))
                metas.append({
                    "id": conv["id"],
                    "title": conv.get("title", "(제목 없음)"),
                    "updated": conv.get("updated", 0),
                    "message_count": len(conv.get("messages", [])),
                })
            except (OSError, json.JSONDecodeError, KeyError):
                continue
        return sorted(metas, key=lambda m: m["updated"], reverse=True)

    @staticmethod
    def title_from(text: str) -> str:
        first_line = text.strip().splitlines()[0] if text.strip() else "새 대화"
        return first_line[:40] + ("…" if len(first_line) > 40 else "")
