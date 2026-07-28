"""projects.py

프로젝트별로 **프롬프트와 메모리를 격리**한다 (claude Projects 스타일).

설계: **폴더 = 프로젝트**. `ConversationStore`/`MemoryStore`가 이미 디렉터리 단위라,
프로젝트마다 폴더 하나를 주고 그 안에 대화·메모리를 담는다. 물리적으로 격리돼 삭제·백업이
깔끔하고 스토어 코드는 그대로 재사용한다.

    <data_dir>/
      conversations/   ← "기본" 공간(프로젝트 없음) 대화 — 기존 그대로
      memory.db        ← "기본" 공간 메모리 — 기존 그대로
      projects/
        <id>/project.json      { id, name, prompt, created, updated }
        <id>/conversations/     이 프로젝트의 대화
        <id>/memory.db          이 프로젝트의 메모리(완전 격리)

기존 데이터는 그대로 "기본" 공간으로 산다 — 마이그레이션이 없다. 프로젝트는 순수 추가.
프롬프트 결합은 **대체**: 프로젝트에 프롬프트가 있으면 그것만 쓰고, 비우면 전역으로 폴백.
메모리는 **완전 격리**: 회상·자동저장·비우기가 전부 그 프로젝트 memory.db만 대상으로 한다.
"""

from __future__ import annotations

import json
import shutil
import threading
import time
import uuid
from pathlib import Path

from server.conversations import ConversationStore
from server.memory import MemoryStore


class ProjectManager:
    """프로젝트 폴더와 프로젝트별 (ConversationStore, MemoryStore)를 관리한다.

    project_id가 falsy면 생성자에서 받은 루트 스토어(기본 공간)를 그대로 돌려준다.
    프로젝트 스토어는 지연 생성하고 캐시한다(매 요청 sqlite 재연결 방지).
    """

    def __init__(self, data_dir: Path, default_store: ConversationStore,
                 default_memory: MemoryStore):
        self.dir = data_dir / "projects"
        self.dir.mkdir(exist_ok=True)
        self._default_store = default_store
        self._default_memory = default_memory
        self._cache: dict[str, tuple[ConversationStore, MemoryStore]] = {}
        self._lock = threading.Lock()

    # ---------- 경로 ----------

    @staticmethod
    def _safe(pid: str) -> str:
        return "".join(ch for ch in (pid or "") if ch.isalnum() or ch in "-_")

    def _pdir(self, pid: str) -> Path:
        return self.dir / self._safe(pid)

    def _meta_path(self, pid: str) -> Path:
        return self._pdir(pid) / "project.json"

    # ---------- CRUD ----------

    def create(self, name: str, prompt: str = "") -> dict:
        pid = f"{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:6]}"
        meta = {
            "id": pid,
            "name": (name or "새 프로젝트").strip()[:80] or "새 프로젝트",
            "prompt": prompt or "",
            "created": time.time(),
            "updated": time.time(),
        }
        self._pdir(pid).mkdir(parents=True, exist_ok=True)
        self._save_meta(meta)
        return meta

    def _save_meta(self, meta: dict) -> None:
        meta["updated"] = time.time()
        self._meta_path(meta["id"]).write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def get(self, pid: str) -> dict | None:
        path = self._meta_path(pid)
        if not path.exists():
            return None
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None

    def list(self) -> list[dict]:
        """프로젝트 메타 목록(프롬프트 본문 제외, 대화 수 포함). 최근 갱신 순."""
        out = []
        for meta_path in self.dir.glob("*/project.json"):
            try:
                meta = json.loads(meta_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            conv_dir = meta_path.parent / "conversations"
            out.append({
                "id": meta["id"],
                "name": meta.get("name", "(이름 없음)"),
                "has_prompt": bool(meta.get("prompt", "").strip()),
                "updated": meta.get("updated", 0),
                "conversation_count": len(list(conv_dir.glob("*.json"))) if conv_dir.is_dir() else 0,
            })
        return sorted(out, key=lambda m: m["updated"], reverse=True)

    def update(self, pid: str, name: str | None = None, prompt: str | None = None) -> dict | None:
        meta = self.get(pid)
        if meta is None:
            return None
        if name is not None:
            meta["name"] = name.strip()[:80] or meta["name"]
        if prompt is not None:
            meta["prompt"] = prompt
        self._save_meta(meta)
        return meta

    def delete(self, pid: str) -> bool:
        """프로젝트 폴더를 통째로 지운다(대화·메모리 포함, 🔴 되돌릴 수 없음)."""
        pdir = self._pdir(pid)
        if not pdir.is_dir():
            return False
        with self._lock:
            cached = self._cache.pop(self._safe(pid), None)
        if cached is not None:
            cached[1].close()  # memory.db 잠금 해제(Windows 삭제 차단 방지)
        shutil.rmtree(pdir, ignore_errors=True)
        return not pdir.exists()

    # ---------- 스토어 해석 ----------

    def stores(self, pid: str | None) -> tuple[ConversationStore, MemoryStore]:
        """(대화 스토어, 메모리 스토어)를 돌려준다. pid가 falsy거나 없는 프로젝트면
        기본 공간(루트 스토어)으로 우아하게 폴백한다 — chat이 잘못된 pid로 죽지 않게."""
        if not pid:
            return self._default_store, self._default_memory
        key = self._safe(pid)
        with self._lock:
            if key in self._cache:
                return self._cache[key]
        if not self._meta_path(pid).exists():
            # 존재하지 않는 프로젝트 → 기본 공간으로 폴백(chat이 죽지 않게)
            return self._default_store, self._default_memory
        pdir = self._pdir(pid)
        pdir.mkdir(parents=True, exist_ok=True)
        pair = (ConversationStore(pdir), MemoryStore(pdir))
        with self._lock:
            self._cache.setdefault(key, pair)
            return self._cache[key]

    def base_prompt(self, pid: str | None, global_prompt: str) -> str:
        """이 프로젝트의 기본 시스템 프롬프트. 프로젝트 프롬프트가 있으면 그것(대체),
        비어 있거나 프로젝트가 없으면 전역 프롬프트로 폴백."""
        if pid:
            meta = self.get(pid)
            if meta and meta.get("prompt", "").strip():
                return meta["prompt"]
        return global_prompt
