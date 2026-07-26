"""memory.py

대화를 넘나드는 **장기 기억**을 sqlite 파일 하나(<data_dir>/memory.db)에 저장한다.
conversations.py의 얇은 스토어 스타일을 그대로 따르며, 나중에 그대로 MCP 서버
(memory__remember / memory__recall)로 뽑아낼 수 있게 인터페이스를 얇게 유지한다.

검색은 FTS5의 **trigram 토크나이저**를 쓴다. 기본 unicode61 토크나이저는 한국어를
형태소로 못 쪼개서 "한국어"가 저장된 토큰 "한국어를"과 안 맞는다. trigram은 3글자
단위로 쪼개므로 조사가 붙은 변형까지 매칭된다("한국어를" 질의 → "한국어" 항목 적중).
단 3자 미만 항은 trigram으로 안 걸리고, FTS5 자체가 없는 sqlite 빌드도 있으므로
그럴 때는 LIKE 검색으로 **우아하게 저하**한다 (기능이 죽지 않는다).

모든 저장/질의 텍스트는 NFC로 정규화한다 (한글 자모 합성/분해 표기가 섞이면
같은 글자라도 매칭이 깨지기 때문).
"""

from __future__ import annotations

import sqlite3
import threading
import time
import unicodedata
from pathlib import Path


def _nfc(text: str) -> str:
    """한글 정규화. 합성/분해 표기 차이로 매칭이 깨지는 걸 막는다."""
    return unicodedata.normalize("NFC", text or "").strip()


class MemoryStore:
    """장기 기억 하나를 sqlite로 관리한다. 스레드 안전(단일 연결 + 락)."""

    def __init__(self, data_dir: Path):
        self.path = data_dir / "memory.db"
        self._lock = threading.Lock()
        self._fts = False  # FTS5 사용 가능 여부 (아래에서 실제로 시험해 정한다)
        # 여러 스레드(FastAPI 이벤트 루프, 재시작 executor 등)에서 부르므로 연결을 공유하고
        # 모든 접근을 락으로 감싼다. 볼륨이 작아 성능 문제는 없다.
        self.conn = sqlite3.connect(str(self.path), check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    # ---------- 스키마 ----------

    def _init_schema(self) -> None:
        with self._lock:
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS memories (
                    id INTEGER PRIMARY KEY,
                    content TEXT NOT NULL,
                    kind TEXT DEFAULT 'fact',
                    source_conversation_id TEXT,
                    created REAL,
                    updated REAL
                )
                """
            )
            self._fts = self._try_init_fts()
            self.conn.commit()

    def _try_init_fts(self) -> bool:
        """FTS5(trigram) 가상 테이블과 동기화 트리거를 만든다. 안 되면 False."""
        try:
            self.conn.execute(
                "CREATE VIRTUAL TABLE IF NOT EXISTS memories_fts "
                "USING fts5(content, content='memories', content_rowid='id', tokenize='trigram')"
            )
            # 본문 테이블 변경을 FTS 인덱스에 그대로 반영하는 트리거들.
            self.conn.executescript(
                """
                CREATE TRIGGER IF NOT EXISTS memories_ai AFTER INSERT ON memories BEGIN
                    INSERT INTO memories_fts(rowid, content) VALUES (new.id, new.content);
                END;
                CREATE TRIGGER IF NOT EXISTS memories_ad AFTER DELETE ON memories BEGIN
                    INSERT INTO memories_fts(memories_fts, rowid, content)
                        VALUES ('delete', old.id, old.content);
                END;
                CREATE TRIGGER IF NOT EXISTS memories_au AFTER UPDATE ON memories BEGIN
                    INSERT INTO memories_fts(memories_fts, rowid, content)
                        VALUES ('delete', old.id, old.content);
                    INSERT INTO memories_fts(rowid, content) VALUES (new.id, new.content);
                END;
                """
            )
            return True
        except sqlite3.OperationalError as e:
            # FTS5가 없는 sqlite 빌드 → LIKE 검색으로 저하. 로그만 남기고 계속 동작.
            print(f"[주의] FTS5를 쓸 수 없어 LIKE 검색으로 저하합니다: {e}")
            return False

    # ---------- 쓰기 ----------

    def add(self, content: str, kind: str = "fact",
            source_conversation_id: str | None = None) -> int | None:
        """사실 하나를 저장한다. NFC 정규화 후 동일 내용이 이미 있으면 저장하지 않고
        그 항목의 시각만 갱신한다(간단 dedup). 빈 내용은 무시하고 None을 반환한다."""
        norm = _nfc(content)
        if not norm:
            return None
        now = time.time()
        with self._lock:
            existing = self.conn.execute(
                "SELECT id FROM memories WHERE content = ?", (norm,)
            ).fetchone()
            if existing:
                self.conn.execute(
                    "UPDATE memories SET updated = ? WHERE id = ?", (now, existing["id"])
                )
                self.conn.commit()
                return int(existing["id"])
            cur = self.conn.execute(
                "INSERT INTO memories(content, kind, source_conversation_id, created, updated) "
                "VALUES (?, ?, ?, ?, ?)",
                (norm, kind, source_conversation_id, now, now),
            )
            self.conn.commit()
            return int(cur.lastrowid)

    def update(self, mem_id: int, content: str) -> bool:
        norm = _nfc(content)
        if not norm:
            return False
        with self._lock:
            cur = self.conn.execute(
                "UPDATE memories SET content = ?, updated = ? WHERE id = ?",
                (norm, time.time(), mem_id),
            )
            self.conn.commit()
            return cur.rowcount > 0

    def delete(self, mem_id: int) -> bool:
        with self._lock:
            cur = self.conn.execute("DELETE FROM memories WHERE id = ?", (mem_id,))
            self.conn.commit()
            return cur.rowcount > 0

    # ---------- 읽기 ----------

    # 질의 하나가 만드는 trigram/어근 수 상한. 긴 질의가 거대한 OR 절을 만드는 걸 막는다.
    _MAX_TRIGRAMS = 64
    _MAX_ROOTS = 32

    # 어절 끝에서 떼어낼 흔한 조사·어미 (긴 것부터 — 먼저 매칭되게). 형태소 분석기가
    # 없는 폐쇄망에서 "이름이"↔"이름은"처럼 조사만 다른 어근을 잇기 위한 최소 휴리스틱.
    _JOSA = (
        "으로서", "으로써", "이라고", "이라는", "에서", "에게", "한테", "까지", "부터",
        "보다", "처럼", "마다", "라고", "라는", "이며", "이고", "으로", "은", "는",
        "이", "가", "을", "를", "의", "에", "도", "만", "과", "와", "로", "야", "께", "랑",
    )

    @staticmethod
    def _trigrams(term: str) -> list[str]:
        """어절 하나를 겹치는 3글자 창으로 쪼갠다. 3자 미만이면 빈 리스트."""
        return [term[i:i + 3] for i in range(len(term) - 2)] if len(term) >= 3 else []

    def _query_trigrams(self, terms: list[str]) -> list[str]:
        """질의 어절들에서 trigram 집합을 만든다 (순서 유지 dedup, 상한 적용).

        질의를 trigram으로 분해해야 조사·어미가 붙은 어절("폐쇄망은")이 저장된
        어근("폐쇄망")과 겹치는 trigram으로 매칭된다. 통째로 phrase 매칭하면
        조사 때문에 안 걸린다.
        """
        tris: list[str] = []
        for t in terms:
            tris.extend(self._trigrams(t))
        return list(dict.fromkeys(tris))[: self._MAX_TRIGRAMS]

    def _strip_josa(self, word: str) -> str:
        """어절 끝의 조사를 한 번 떼어 어근을 만든다 (어근이 2자 미만이 되면 안 뗀다)."""
        for j in self._JOSA:
            if word.endswith(j) and len(word) - len(j) >= 2:
                return word[: -len(j)]
        return word

    def _query_roots(self, terms: list[str]) -> list[str]:
        """조사를 뗀 2자 이상 어근 집합. LIKE 부분 문자열 매칭용.

        trigram은 3자 미만 어근("이름")을 못 잡고, 조사가 다르면("이름이"↔"이름은")
        3-gram이 안 겹친다. 어근으로 LIKE 하면 두 경우 다 걸린다.
        """
        roots = []
        for t in terms:
            r = self._strip_josa(t)
            if len(r) >= 2:
                roots.append(r)
        return list(dict.fromkeys(roots))[: self._MAX_ROOTS]

    def search(self, query: str, limit: int = 5) -> list[dict]:
        """질의와 관련된 기억을 최대 limit개 돌려준다.

        두 신호를 합친다: (1) trigram FTS5(있으면 bm25로 랭킹, 3자 이상 겹침에 정밀),
        (2) 조사 뗀 어근 LIKE(조사 변형·2자 어근을 잡는 저하 경로). 관련 없으면 빈 리스트.
        """
        norm = _nfc(query)
        if not norm:
            return []
        terms = [t for t in norm.split() if t]
        tris = self._query_trigrams(terms)
        roots = self._query_roots(terms)

        results: dict = {}  # id -> row (삽입 순서 = 우선순위)
        # 1) 정밀: trigram FTS (bm25 랭킹). FTS5가 없으면 건너뛴다.
        if self._fts and tris:
            for r in self._fts_search(tris, limit):
                results.setdefault(r["id"], r)
        # 2) 보강: 조사 뗀 어근으로 부분 문자열 매칭.
        if roots and len(results) < limit:
            for r in self._like_search(roots, limit):
                results.setdefault(r["id"], r)
        return list(results.values())[:limit]

    def _fts_search(self, tris: list[str], limit: int) -> list[dict]:
        """trigram OR 매칭 + bm25 랭킹. MATCH 파싱 오류 시 빈 리스트."""
        match = " OR ".join('"' + g.replace('"', '""') + '"' for g in tris)
        try:
            with self._lock:
                rows = self.conn.execute(
                    "SELECT m.id, m.content, m.kind, m.created, m.updated "
                    "FROM memories_fts f JOIN memories m ON m.id = f.rowid "
                    "WHERE memories_fts MATCH ? ORDER BY bm25(memories_fts) LIMIT ?",
                    (match, limit),
                ).fetchall()
            return [dict(r) for r in rows]
        except sqlite3.OperationalError:
            return []

    def _like_search(self, needles: list[str], limit: int) -> list[dict]:
        """주어진 조각들을 부분 문자열(LIKE)로 OR 매칭. 최근 갱신 순."""
        if not needles:
            return []
        where = " OR ".join(["content LIKE ?"] * len(needles))
        params = [f"%{n}%" for n in needles] + [limit]
        with self._lock:
            rows = self.conn.execute(
                f"SELECT id, content, kind, created, updated FROM memories "
                f"WHERE {where} ORDER BY updated DESC LIMIT ?",
                params,
            ).fetchall()
        return [dict(r) for r in rows]

    def recent(self, limit: int = 50) -> list[dict]:
        with self._lock:
            rows = self.conn.execute(
                "SELECT id, content, kind, created, updated FROM memories "
                "ORDER BY updated DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [dict(r) for r in rows]

    def all(self) -> list[dict]:
        """관리/조회용 전체 목록. 최근 갱신 순."""
        return self.recent(limit=100000)

    def count(self) -> int:
        with self._lock:
            return int(self.conn.execute("SELECT COUNT(*) FROM memories").fetchone()[0])
