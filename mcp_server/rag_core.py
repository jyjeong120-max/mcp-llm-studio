"""rag_core.py

RAG 공용 코어 — 구성(rag_indexer.py)과 서빙(rag_server.py)이 함께 쓰는
문서 읽기·청킹·임베딩·저장소 로직. 이 모듈 자체는 실행 파일이 아니다.

구성 (전부 로컬 — 폐쇄망 전제, 런타임 외부 HTTP는 localhost 임베딩 서버뿐):
    문서 읽기 : office_server.py의 COM 헬퍼(_document)를 재사용한다. Word가 문서를
                열어 주므로 사내 DRM으로 암호화된 docx도 사용자 PC에서 읽히는 그대로
                읽힌다(바이트 파싱으로는 암호문만 나오는 환경 대응).
    청킹     : 문단 경계 존중, 기본 1000자 / 겹침 200자.
    임베딩   : llama-server --embeddings (OpenAI 호환 /v1/embeddings, 기본 :8001).
                EmbeddingGemma 등 GGUF 임베딩 모델을 CPU(-ngl 0)로 상주시키면 된다.
                질의/문서에 모델 규격 프리픽스를 붙인다(EMBED_QUERY_PREFIX/DOC_TEMPLATE
                — EmbeddingGemma 기본값). **임베딩 서버가 없으면 키워드 인덱스만 만들고,
                검색도 키워드만으로 우아하게 저하한다** — 나중에 서버를 띄우고 reindex하면
                벡터가 붙는다.
    리랭크   : (선택) llama-server --reranking (bge-reranker 등 GGUF, 기본 :8002).
                하이브리드 상위 후보를 질의-문서 관련도로 재정렬해 정밀도를 높인다.
                **서버가 없으면 RRF 순서를 그대로 쓴다** — 임베딩과 같은 우아한 저하.
    저장/검색 : 청크 본문·키워드 인덱스는 sqlite(rag_index.db — FTS5 trigram,
                llm_studio memory.py와 같은 조사 변형 대응 패턴). 벡터는
                **qdrant-client가 있으면 Qdrant 로컬(파일) 모드**(rag_vectors/,
                HNSW 근사 검색 — 서버 프로세스 없음)에, 없으면 sqlite BLOB에 저장하고
                코사인을 직접 계산한다(numpy 있으면 빠르게, 없으면 순수 파이썬).

⚠ Qdrant 로컬 모드는 단일 프로세스 잠금이 있다 — 구성(rag_indexer)과 서빙(rag_server)을
동시에 돌리지 말 것. rag_indexer는 잠겨 있으면 시작을 거부한다(서빙과 다른 저장소에
벡터가 쌓여 검색이 어긋나는 사고 방지). 서빙 쪽은 잠금을 못 잡으면 sqlite BLOB으로
저하하고 rag_status에 사유를 표시한다.
"""

from __future__ import annotations

import json
import math
import os
import re
import sqlite3
import struct
import sys
import threading
import time
import unicodedata
import urllib.error
import urllib.request
from functools import wraps
from pathlib import Path

# office_server의 문서 확보 로직을 그대로 재사용한다 (열려 있으면 그 세션, 아니면
# 백그라운드 읽기 전용으로 열었다 닫음 + 암호 문서 대화상자 억제까지 검증된 경로).
try:
    from office_server import COM_AVAILABLE, OfficeError, _document

    OFFICE_IMPORT_ERROR = ""
except Exception as e:  # noqa: BLE001 — fastmcp/pywin32 부재 등 어떤 실패든 저하로
    COM_AVAILABLE = False
    OFFICE_IMPORT_ERROR = str(e)
    OfficeError = RuntimeError  # type: ignore[assignment,misc]
    _document = None  # type: ignore[assignment]

try:
    import pythoncom  # COM 스레드 초기화용 (인덱싱 경로에서만 필요)
except ImportError:
    pythoncom = None  # type: ignore[assignment]

try:
    import numpy as _np  # 선택 의존성 — 없으면 순수 파이썬 코사인으로 저하
except ImportError:
    _np = None

# 벡터 저장 백엔드: qdrant-client(사내 미러 등록됨)가 있으면 Qdrant 로컬(파일) 모드.
# 없거나 열기 실패(잠금 등)면 sqlite BLOB으로 우아하게 저하한다.
try:
    from qdrant_client import QdrantClient
    from qdrant_client.models import Distance, PointIdsList, PointStruct, VectorParams

    QDRANT_AVAILABLE = True
    QDRANT_IMPORT_ERROR = ""
except ImportError as e:
    QDRANT_AVAILABLE = False
    QDRANT_IMPORT_ERROR = str(e)

# ─────────────────────────────── 설정 (env/CLI로 조정) ───────────────────────────────
# CLI가 덮어쓸 수 있으므로 다른 모듈에서는 `core.DB_PATH`처럼 매번 속성으로 읽을 것
# (from-import로 값을 복사하면 덮어쓴 게 안 보인다).

# 임베딩 서버 (llama-server --embeddings). /v1 까지 포함한 베이스 URL.
EMBED_URL = os.getenv("RAG_EMBED_URL", "http://127.0.0.1:8001/v1")
# 리랭커 서버 (llama-server --reranking, bge-reranker 등 GGUF). /v1 까지 포함한 베이스 URL.
# 서버가 없거나 응답이 없으면 리랭크를 건너뛰고 RRF 순서를 그대로 쓴다(우아한 저하).
RERANK_URL = os.getenv("RAG_RERANK_URL", "http://127.0.0.1:8002/v1")
# 인덱스 파일 위치. 기본은 이 스크립트 옆.
DB_PATH = os.getenv("RAG_DB", str(Path(__file__).with_name("rag_index.db")))
# Qdrant 로컬 모드 데이터 폴더 (qdrant-client 있을 때만 사용).
QDRANT_PATH = os.getenv("RAG_QDRANT", str(Path(__file__).with_name("rag_vectors")))

CHUNK_SIZE = 1000      # 청크 목표 길이(문자)
CHUNK_OVERLAP = 200    # 청크 사이 겹침(문자)
EMBED_BATCH = 16       # 임베딩 요청 한 번에 보낼 청크 수
EMBED_TIMEOUT = 120    # 임베딩 요청 타임아웃(초) — CPU 서빙이면 배치가 느릴 수 있다
DOC_PATTERNS = (".docx", ".doc")  # 인덱싱 대상 확장자
MAX_RESULT_CHARS = 1200           # 검색 결과에서 청크 하나당 보여줄 최대 길이
RRF_K = 60                        # RRF 상수 (관례값)

# 임베딩 프롬프트 — EmbeddingGemma 규격(모델 카드). 질의와 문서에 서로 다른 프리픽스를
# 붙여야 검색 품질이 제대로 나온다(안 붙이면 같은 모델인데 recall이 떨어짐). 다른 모델로
# 바꾸면 이 두 값을 그 모델 포맷으로 교체할 것 — E5: "query: "/"passage: ", BGE-m3: 프리픽스
# 없음. 빈 문자열/템플릿로 두면 프리픽스 없이(raw) 임베딩한다.
EMBED_QUERY_PREFIX = os.getenv("RAG_EMBED_QUERY_PREFIX", "task: search result | query: ")
EMBED_DOC_TEMPLATE = os.getenv("RAG_EMBED_DOC_TEMPLATE", "title: {title} | text: {text}")

# 검색 결과에 붙일 이웃 청크 수(같은 섹션 한정, 앞뒤 각각). 0이면 매칭 청크만. 근거가
# 청크 경계에서 잘리는 걸 막는 small-to-big 확장이다.
CONTEXT_WINDOW = int(os.getenv("RAG_CONTEXT_WINDOW", "1"))
MAX_CONTEXT_CHARS = 2400          # 이웃 청크까지 합친 블록의 최대 표시 길이


class RagError(Exception):
    """도구가 사용자에게 그대로 돌려줄 안내 메시지를 담은 예외."""


def rag_tool(fn):
    """예외를 안내 문자열로 바꾸고, COM이 필요한 경로를 위해 스레드를 초기화한다.

    FastMCP는 동기 도구를 워커 스레드에서 실행한다. Word를 부르는 인덱싱 경로는
    스레드마다 CoInitialize가 필요하므로 매 호출 초기화/해제한다(검색은 COM을
    안 쓰지만 초기화가 무해해서 공통으로 감싼다).
    """

    @wraps(fn)
    def wrapper(*args, **kwargs):
        if pythoncom is not None:
            pythoncom.CoInitialize()
        try:
            return fn(*args, **kwargs)
        except (RagError, OfficeError) as e:
            return str(e)
        except Exception as e:  # noqa: BLE001 — 도구는 항상 문자열을 돌려준다
            return f"작업에 실패했습니다: {type(e).__name__}: {e}"
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()

    return wrapper


def _nfc(text: str) -> str:
    """한글 정규화(NFC). 합성/분해 표기 차이로 매칭이 깨지는 걸 막는다."""
    return unicodedata.normalize("NFC", text or "").strip()


# ─────────────────────────────── 문서 읽기 (Word COM) ───────────────────────────────


def _require_word():
    if _document is None or not COM_AVAILABLE:
        raise RagError(
            "Word 문서를 읽을 수 없습니다. office_server의 COM 헬퍼를 불러오지 못했습니다"
            f"({OFFICE_IMPORT_ERROR or 'pywin32 없음'}). Windows + Office + pywin32 환경에서 "
            "실행하세요. (검색은 이미 만든 인덱스로 계속 동작합니다.)"
        )


def _clean_word_text(raw: str) -> str:
    """Word Content.Text의 COM 특수문자를 일반 텍스트로 정리한다.

    \\r=문단 끝, \\x07=표 셀 구분, \\x0b=수동 줄바꿈, \\x0c=쪽 나눔.
    표 셀은 탭으로 이어 한 행이 한 줄이 되게 한다(청킹 시 표가 흩어지지 않도록).
    """
    text = raw.replace("\r\x07", "\t").replace("\x07", "\t")
    text = text.replace("\r", "\n").replace("\x0b", "\n").replace("\x0c", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    return text


def _extract_word_text(path: str, password: str = "") -> str:
    """Word 문서 전체 본문을 텍스트로 뽑는다 (COM 왕복 1회 — 문단별 순회보다 빠름)."""
    _require_word()
    with _document("word", path, password) as doc:
        raw = doc.Content.Text
    return _clean_word_text(raw or "")


# ─────────────────────────────── 청킹 ───────────────────────────────


def _split_paragraphs(text: str) -> list[str]:
    return [p.strip() for p in text.split("\n") if p.strip()]


# ─────────────────────────────── 헤딩(제목) 휴리스틱 ───────────────────────────────
# doc.Content.Text에는 서식/스타일 정보가 없어(통짜 텍스트) '제목 1' 같은 Word 스타일을
# 알 수 없다. 그래서 번호 체계·길이로 제목을 '추정'한다. 사내 문서 주 체계인 숫자 점
# 표기(1 / 1.2 / 1.2.3)와 장·절 표기를 잡고, 애매하면 헤딩으로 보지 않는다(보수적).
# ⚠ 오탐/누락이 있을 수 있는 휴리스틱이다 — 문서 체계가 다르면 아래 패턴을 조정할 것.

_HEADING_MAX_LEN = 50  # 이보다 길면 제목이 아니라 본문으로 본다
# 숫자 점 표기 뒤에 제목 텍스트가 오는 형태: "1 개요", "3.2 점검주기", "1.2.3. 세부"
_HEADING_NUM = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+(\S.*)$")
_HEADING_JANG = re.compile(r"^제\s*\d+\s*장\b")   # 제1장
_HEADING_JEOL = re.compile(r"^제\s*\d+\s*절\b")   # 제1절
# 서술형 종결로 끝나면 제목이 아니라 순서 목록 본문("1. 먼저 ~한다.")으로 보고 제외한다.
# 명사가 우연히 요/음/다로 끝나는 경우(개요·요약·항목…)를 오탐하지 않도록, '마침표로
# 끝나는 종결(…다.)' 또는 '명확한 서술 어미(…습니다/…한다)'만 본문으로 본다.
_SENTENCE_TAIL = re.compile(r"[다요음함임됨]\.$|(?:니다|습니다|하다|한다|된다|이다|였다|었다)$")


def _heading_level(line: str) -> int:
    """한 줄이 제목이면 계층 레벨(1=최상위)을, 아니면 0을 돌려준다 (텍스트 휴리스틱).

    표 행(탭 포함)·긴 줄·서술형 종결 줄은 제외한다. 숫자 점 표기는 점 개수+1을
    레벨로 삼아 대제목/소제목 깊이를 구분한다.
    """
    s = line.strip()
    if not s or "\t" in s or len(s) > _HEADING_MAX_LEN:
        return 0
    if _HEADING_JANG.match(s):
        return 1
    if _HEADING_JEOL.match(s):
        return 2
    m = _HEADING_NUM.match(s)
    if m and not _SENTENCE_TAIL.search(m.group(2).strip()):
        return m.group(1).count(".") + 1
    return 0


def _chunk_text(
    text: str, size: int = CHUNK_SIZE, overlap: int = CHUNK_OVERLAP
) -> list[tuple[str, str]]:
    """문단 경계를 존중하며 목표 길이로 (섹션경로, 본문) 청크를 만든다.

    제목으로 추정되는 문단을 만나면 (1) 지금 모으던 청크를 닫아 섹션이 청크에
    섞이지 않게 하고 (2) 계층 스택을 갱신해 '대제목 › 소제목' 경로를 만든다.
    제목 줄 자체는 그 섹션 첫 청크 본문의 머리에 넣어 키워드(FTS) 검색으로도
    잡히게 한다. 같은 섹션 안에서 목표를 넘으면 이전 청크의 꼬리(overlap자)를
    겹치고, 목표의 1.5배가 넘는 초장문 문단은 문장 경계로 강제 분할한다.
    """
    paras = _split_paragraphs(_nfc(text))
    chunks: list[tuple[str, str]] = []
    stack: list[tuple[int, str]] = []  # (레벨, 제목) — 현재 섹션 경로
    buf = ""
    buf_path = ""  # 지금 buf가 속한 섹션 경로 (flush 시점에 함께 저장)

    def flush():
        nonlocal buf
        if buf.strip():
            chunks.append((buf_path, buf.strip()))
        buf = ""

    def carry() -> str:
        """직전 청크의 꼬리를 겹침으로 넘긴다."""
        return chunks[-1][1][-overlap:] if (chunks and overlap > 0) else ""

    for p in paras:
        lvl = _heading_level(p)
        if lvl:
            # 제목 경계: 이전 청크를 닫고 계층 스택을 갱신한 뒤, 제목 줄로 새 청크를 연다.
            flush()
            while stack and stack[-1][0] >= lvl:
                stack.pop()
            stack.append((lvl, p))
            buf_path = " › ".join(t for _, t in stack)
            buf = p
            continue
        if len(p) > size * 1.5:
            # 초장문 문단: 지금까지 모은 것을 닫고, 문장 단위로 창을 밀며 자른다.
            flush()
            sentences = re.split(r"(?<=[.。!?다\.])\s+", p) or [p]
            piece = carry()
            for s in sentences:
                if len(piece) + len(s) + 1 > size and piece.strip():
                    chunks.append((buf_path, piece.strip()))
                    piece = piece[-overlap:] if overlap > 0 else ""
                piece = (piece + " " + s).strip() if piece else s
                # 문장 하나가 그 자체로 너무 길면 고정 폭으로 최종 분할
                while len(piece) > size * 1.5:
                    chunks.append((buf_path, piece[:size]))
                    piece = piece[size - overlap:] if overlap > 0 else piece[size:]
            if piece.strip():
                chunks.append((buf_path, piece.strip()))
            continue
        if buf and len(buf) + len(p) + 1 > size:
            flush()
            head = carry()
            buf = (head + "\n" + p) if head else p
        else:
            buf = (buf + "\n" + p) if buf else p
    flush()
    return chunks


def _merge_overlapping(texts: list[str], max_overlap: int = CHUNK_OVERLAP) -> str:
    """이웃 청크들을 겹침(overlap)을 제거하며 하나로 잇는다(검색 결과 컨텍스트 확장용).

    연속 청크는 앞 청크의 꼬리 일부를 공유하므로(청킹의 carry), 뒤 청크에서 그 겹치는
    접두부만큼을 잘라내고 이어붙인다. 겹침이 없으면 줄바꿈으로 잇는다.
    """
    parts = [t for t in texts if t]
    if not parts:
        return ""
    merged = parts[0]
    for t in parts[1:]:
        ov = 0
        limit = min(len(merged), len(t), max_overlap + 40)  # carry + 줄바꿈 여유
        for n in range(limit, 0, -1):
            if merged[-n:] == t[:n]:
                ov = n
                break
        merged += t[ov:] if ov else ("\n" + t)
    return merged


# ─────────────────────────────── 임베딩 클라이언트 ───────────────────────────────


def _normalize(vec: list[float]) -> list[float]:
    n = math.sqrt(sum(v * v for v in vec))
    return [v / n for v in vec] if n > 0 else vec


def _embed_texts(texts: list[str]) -> list[list[float]] | None:
    """임베딩 서버에 배치 요청. 어떤 실패든 None을 돌려 키워드 전용으로 저하한다.

    반환 벡터는 단위 길이로 정규화한다 (검색 때 내적 = 코사인 유사도).
    """
    if not texts:
        return []
    url = EMBED_URL.rstrip("/") + "/embeddings"
    body = json.dumps({"model": "embedding", "input": texts}).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, headers={"Content-Type": "application/json"}, method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=EMBED_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, OSError):
        return None
    items = sorted(data.get("data") or [], key=lambda d: d.get("index", 0))
    vecs = [it.get("embedding") for it in items]
    if len(vecs) != len(texts) or any(not isinstance(v, list) or not v for v in vecs):
        return None
    return [_normalize([float(x) for x in v]) for v in vecs]


def _embed_available() -> bool:
    """임베딩 서버가 응답하는지 1건짜리 요청으로 확인한다."""
    return bool(_embed_texts(["ping"]))


def _embed_query_text(query: str) -> str:
    """질의 임베딩용 프롬프트를 만든다 — EmbeddingGemma 등은 질의 프리픽스가 따로 있다."""
    return (EMBED_QUERY_PREFIX + query) if EMBED_QUERY_PREFIX else query


def _embed_doc_text(content: str, title: str = "") -> str:
    """문서 청크 임베딩용 프롬프트를 만든다.

    title에는 섹션 경로(대제목 › 소제목)를 넣어 하위 청크에도 상위 제목 맥락이 벡터에
    담기게 한다(없으면 'none'). 템플릿이 비어 있으면 프리픽스 없이 제목+본문만 잇는다.
    """
    if not EMBED_DOC_TEMPLATE:
        return f"{title}\n{content}" if title else content
    return EMBED_DOC_TEMPLATE.format(title=title or "none", text=content)


# ─────────────────────────────── 리랭커 클라이언트 ───────────────────────────────
# 하이브리드로 뽑은 상위 후보를 질의-문서 관련도로 재정렬한다(정밀도 향상). 임베딩과
# 같은 패턴(로컬 llama-server, 우아한 저하) — 서버가 없으면 None을 돌려 RRF 순서를 쓴다.


def _rerank(query: str, documents: list[str]) -> list[tuple[int, float]] | None:
    """리랭커 서버(llama-server --reranking)에 질의-문서 관련도를 매긴다.

    반환: (원본 인덱스, 점수) 목록을 점수 내림차순으로. 서버가 없거나 어떤 실패든
    None을 돌려 호출부가 RRF 순서를 그대로 쓰게 한다(우아한 저하).
    """
    if not documents:
        return []
    url = RERANK_URL.rstrip("/") + "/rerank"
    body = json.dumps(
        {"model": "rerank", "query": query, "documents": documents}
    ).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, headers={"Content-Type": "application/json"}, method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=EMBED_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, OSError):
        return None
    results = data.get("results")
    if not isinstance(results, list) or not results:
        return None
    out: list[tuple[int, float]] = []
    for r in results:
        idx = r.get("index")
        score = r.get("relevance_score", r.get("score"))
        if idx is None or score is None:
            continue
        out.append((int(idx), float(score)))
    if not out:
        return None
    out.sort(key=lambda t: -t[1])
    return out


def _rerank_available() -> bool:
    """리랭커 서버가 응답하는지 1건짜리 요청으로 확인한다."""
    return _rerank("ping", ["pong"]) is not None


def _pack_vec(vec: list[float]) -> bytes:
    return struct.pack(f"<{len(vec)}f", *vec)


def _unpack_vec(blob: bytes) -> list[float]:
    return list(struct.unpack(f"<{len(blob) // 4}f", blob))


# ─────────────────────────────── 저장/검색 (sqlite) ───────────────────────────────
# llm_studio memory.py와 같은 뼈대: FTS5 trigram + 조사 제거 LIKE 저하. 여기에
# 청크 벡터(Qdrant 또는 BLOB)를 더해 하이브리드가 된다.

_JOSA = (
    "으로서", "으로써", "이라고", "이라는", "에서", "에게", "한테", "까지", "부터",
    "보다", "처럼", "마다", "라고", "라는", "이며", "이고", "으로", "은", "는",
    "이", "가", "을", "를", "의", "에", "도", "만", "과", "와", "로", "야", "께", "랑",
)
_MAX_TRIGRAMS = 64
_MAX_ROOTS = 32


class _QdrantVectors:
    """Qdrant 로컬(파일) 모드 벡터 백엔드 — HNSW 근사 검색, 서버 프로세스 없음.

    포인트 id는 sqlite chunks.id를 그대로 쓴다(두 저장소의 유일한 연결 고리).
    임베딩 모델을 바꿔 차원이 달라지면 컬렉션을 재생성한다(경고 후 reindex 필요).
    """

    kind = "qdrant"
    COLLECTION = "chunks"

    def __init__(self, path: str):
        self.client = QdrantClient(path=path)

    def dim(self) -> int:
        try:
            info = self.client.get_collection(self.COLLECTION)
            return int(info.config.params.vectors.size)
        except Exception:  # noqa: BLE001 — 컬렉션이 아직 없으면 0
            return 0

    def _ensure(self, dim: int) -> None:
        cur = self.dim()
        if cur == dim:
            return
        if cur:
            print(
                f"[주의] 임베딩 차원 변경({cur}→{dim}) — 벡터 컬렉션을 재생성합니다. "
                "기존 파일들은 reindex로 다시 인덱싱해야 벡터가 붙습니다.",
                file=sys.stderr,
            )
            self.client.delete_collection(self.COLLECTION)
        self.client.create_collection(
            self.COLLECTION,
            vectors_config=VectorParams(size=dim, distance=Distance.COSINE),
        )

    def upsert(self, ids: list[int], vecs: list[list[float]]) -> None:
        if not ids:
            return
        self._ensure(len(vecs[0]))
        self.client.upsert(
            self.COLLECTION,
            points=[PointStruct(id=int(i), vector=list(v)) for i, v in zip(ids, vecs)],
        )

    def delete(self, ids: list[int]) -> None:
        if ids and self.dim():
            self.client.delete(
                self.COLLECTION,
                points_selector=PointIdsList(points=[int(i) for i in ids]),
            )

    def search(self, qvec: list[float], k: int) -> list[tuple[int, float]]:
        if not self.dim():
            return []
        try:
            pts = self.client.query_points(self.COLLECTION, query=list(qvec), limit=k).points
        except AttributeError:  # 구버전 qdrant-client는 query_points가 없다
            pts = self.client.search(self.COLLECTION, query_vector=list(qvec), limit=k)
        return [(int(p.id), float(p.score)) for p in pts]

    def count(self) -> int:
        try:
            return int(self.client.count(self.COLLECTION).count)
        except Exception:  # noqa: BLE001
            return 0

    def clear(self) -> None:
        try:
            self.client.delete_collection(self.COLLECTION)
        except Exception:  # noqa: BLE001 — 컬렉션이 없으면 그만
            pass


class RagStore:
    """인덱스 sqlite 파일 하나(+ 선택적 Qdrant 벡터 폴더)를 관리한다. 스레드 안전."""

    def __init__(self, path: str, qdrant_path: str = ""):
        self.path = path
        self.qdrant_path = qdrant_path
        self._lock = threading.Lock()
        self._fts = False
        self._vec_cache: tuple[int, list, object] | None = None  # (세대, ids, 행렬) — sqlite 백엔드용
        self._generation = 0
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._init_schema()
        # 벡터 백엔드: Qdrant 로컬 → 실패 시 sqlite BLOB (우아한 저하).
        self.vec: _QdrantVectors | None = None
        self.vec_error = "" if QDRANT_AVAILABLE else f"qdrant-client 없음({QDRANT_IMPORT_ERROR})"
        if QDRANT_AVAILABLE and qdrant_path:
            try:
                self.vec = _QdrantVectors(qdrant_path)
            except Exception as e:  # noqa: BLE001 — 다른 프로세스가 잠근 경우 등
                self.vec_error = f"Qdrant 열기 실패({type(e).__name__}: {e}) — sqlite BLOB으로 저하"
                print(f"[주의] {self.vec_error}", file=sys.stderr)

    @property
    def vec_kind(self) -> str:
        return "qdrant" if self.vec is not None else "sqlite"

    def _init_schema(self) -> None:
        with self._lock:
            self.conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS files (
                    id INTEGER PRIMARY KEY,
                    path TEXT NOT NULL UNIQUE,
                    mtime REAL, size INTEGER,
                    indexed_at REAL, chunk_count INTEGER
                );
                CREATE TABLE IF NOT EXISTS chunks (
                    id INTEGER PRIMARY KEY,
                    file_id INTEGER NOT NULL REFERENCES files(id) ON DELETE CASCADE,
                    seq INTEGER NOT NULL,
                    content TEXT NOT NULL,
                    heading TEXT,
                    embedding BLOB
                );
                CREATE INDEX IF NOT EXISTS idx_chunks_file ON chunks(file_id);
                """
            )
            # 기존 DB 마이그레이션: heading 컬럼이 없으면 추가한다(CREATE는 이미 있는
            # 테이블을 건드리지 않으므로). 재인덱싱 전 기존 청크의 heading은 NULL로 남는다.
            cols = {r[1] for r in self.conn.execute("PRAGMA table_info(chunks)")}
            if "heading" not in cols:
                self.conn.execute("ALTER TABLE chunks ADD COLUMN heading TEXT")
            try:
                self.conn.execute(
                    "CREATE VIRTUAL TABLE IF NOT EXISTS chunks_fts "
                    "USING fts5(content, content='chunks', content_rowid='id', tokenize='trigram')"
                )
                self.conn.executescript(
                    """
                    CREATE TRIGGER IF NOT EXISTS chunks_ai AFTER INSERT ON chunks BEGIN
                        INSERT INTO chunks_fts(rowid, content) VALUES (new.id, new.content);
                    END;
                    CREATE TRIGGER IF NOT EXISTS chunks_ad AFTER DELETE ON chunks BEGIN
                        INSERT INTO chunks_fts(chunks_fts, rowid, content)
                            VALUES ('delete', old.id, old.content);
                    END;
                    """
                )
                self._fts = True
            except sqlite3.OperationalError as e:
                print(f"[주의] FTS5를 쓸 수 없어 LIKE 검색으로 저하합니다: {e}", file=sys.stderr)
                self._fts = False
            self.conn.commit()

    # ---------- 쓰기 ----------

    def file_unchanged(self, path: str, mtime: float, size: int) -> bool:
        with self._lock:
            row = self.conn.execute(
                "SELECT mtime, size FROM files WHERE path = ?", (path,)
            ).fetchone()
        return bool(row) and abs(row["mtime"] - mtime) < 1e-6 and row["size"] == size

    def replace_file(self, path: str, mtime: float, size: int,
                     chunks: list[tuple[str, str]], vectors: list[list[float]] | None) -> int:
        """파일 하나의 청크들을 통째로 교체한다(기존 것 삭제 후 삽입).

        chunks는 (섹션경로, 본문) 튜플 목록이다. 벡터는 백엔드에 따라 Qdrant(청크
        id를 포인트 id로) 또는 sqlite BLOB에 둔다.
        """
        use_qdrant = self.vec is not None
        new_ids: list[int] = []
        with self._lock:
            self.conn.execute("PRAGMA foreign_keys = ON")
            old = self.conn.execute("SELECT id FROM files WHERE path = ?", (path,)).fetchone()
            old_chunk_ids: list[int] = []
            if old:
                old_chunk_ids = [
                    int(r["id"]) for r in self.conn.execute(
                        "SELECT id FROM chunks WHERE file_id = ?", (old["id"],)
                    ).fetchall()
                ]
                self.conn.execute("DELETE FROM chunks WHERE file_id = ?", (old["id"],))
                self.conn.execute("DELETE FROM files WHERE id = ?", (old["id"],))
            cur = self.conn.execute(
                "INSERT INTO files(path, mtime, size, indexed_at, chunk_count) "
                "VALUES (?, ?, ?, ?, ?)",
                (path, mtime, size, time.time(), len(chunks)),
            )
            fid = cur.lastrowid
            for i, (heading, content) in enumerate(chunks):
                blob = None
                if not use_qdrant and vectors and i < len(vectors):
                    blob = _pack_vec(vectors[i])
                c = self.conn.execute(
                    "INSERT INTO chunks(file_id, seq, content, heading, embedding) "
                    "VALUES (?, ?, ?, ?, ?)",
                    (fid, i + 1, content, heading or None, blob),
                )
                new_ids.append(int(c.lastrowid))
            self.conn.commit()
            self._generation += 1  # (sqlite 백엔드) 벡터 캐시 무효화
        if use_qdrant:
            self.vec.delete(old_chunk_ids)
            if vectors:
                self.vec.upsert(new_ids, vectors)
        return len(chunks)

    def remove_missing(self, existing_paths: set[str]) -> int:
        """디스크에서 사라진 파일의 인덱스를 정리한다. 반환: 삭제한 파일 수."""
        gone_chunk_ids: list[int] = []
        with self._lock:
            self.conn.execute("PRAGMA foreign_keys = ON")
            rows = self.conn.execute("SELECT id, path FROM files").fetchall()
            gone = [r for r in rows if r["path"] not in existing_paths]
            for r in gone:
                gone_chunk_ids.extend(
                    int(c["id"]) for c in self.conn.execute(
                        "SELECT id FROM chunks WHERE file_id = ?", (r["id"],)
                    ).fetchall()
                )
                self.conn.execute("DELETE FROM chunks WHERE file_id = ?", (r["id"],))
                self.conn.execute("DELETE FROM files WHERE id = ?", (r["id"],))
            if gone:
                self.conn.commit()
                self._generation += 1
        if self.vec is not None:
            self.vec.delete(gone_chunk_ids)
        return len(gone)

    def clear(self) -> tuple[int, int]:
        with self._lock:
            nf = self.conn.execute("SELECT COUNT(*) FROM files").fetchone()[0]
            nc = self.conn.execute("SELECT COUNT(*) FROM chunks").fetchone()[0]
            self.conn.execute("DELETE FROM chunks")
            self.conn.execute("DELETE FROM files")
            self.conn.commit()
            self._generation += 1
        if self.vec is not None:
            self.vec.clear()
        return int(nf), int(nc)

    # ---------- 통계 ----------

    def stats(self) -> dict:
        with self._lock:
            nf = self.conn.execute("SELECT COUNT(*) FROM files").fetchone()[0]
            nc = self.conn.execute("SELECT COUNT(*) FROM chunks").fetchone()[0]
            nv = self.conn.execute(
                "SELECT COUNT(*) FROM chunks WHERE embedding IS NOT NULL"
            ).fetchone()[0]
            dim_row = self.conn.execute(
                "SELECT LENGTH(embedding) AS l FROM chunks WHERE embedding IS NOT NULL LIMIT 1"
            ).fetchone()
        if self.vec is not None:
            nv = self.vec.count()
            dim = self.vec.dim()
        else:
            dim = (dim_row["l"] // 4) if dim_row else 0
        return {
            "files": int(nf), "chunks": int(nc), "with_vector": int(nv),
            "dim": dim, "fts": self._fts,
        }

    def list_files(self, limit: int = 200) -> list[dict]:
        with self._lock:
            rows = self.conn.execute(
                "SELECT path, chunk_count, indexed_at FROM files ORDER BY path LIMIT ?",
                (limit,),
            ).fetchall()
        return [dict(r) for r in rows]

    # ---------- 검색 ----------

    def _load_vectors(self):
        """(ids, 행렬) — numpy가 있으면 (N×dim) ndarray, 없으면 list[list[float]]."""
        cached = self._vec_cache
        if cached and cached[0] == self._generation:
            return cached[1], cached[2]
        with self._lock:
            rows = self.conn.execute(
                "SELECT id, embedding FROM chunks WHERE embedding IS NOT NULL"
            ).fetchall()
        ids = [r["id"] for r in rows]
        vecs = [_unpack_vec(r["embedding"]) for r in rows]
        matrix = _np.array(vecs, dtype="float32") if (_np is not None and vecs) else vecs
        self._vec_cache = (self._generation, ids, matrix)
        return ids, matrix

    def vector_search(self, query_vec: list[float], top_k: int) -> list[tuple[int, float]]:
        """정규화된 질의 벡터로 코사인 상위 top_k개 (chunk_id, 점수)."""
        if self.vec is not None:
            return self.vec.search(query_vec, top_k)
        ids, matrix = self._load_vectors()
        if not ids:
            return []
        if _np is not None:
            scores = matrix @ _np.array(query_vec, dtype="float32")
            order = _np.argsort(-scores)[:top_k]
            return [(ids[int(i)], float(scores[int(i)])) for i in order]
        scored = []
        for cid, vec in zip(ids, matrix):
            s = sum(a * b for a, b in zip(vec, query_vec))
            scored.append((cid, s))
        scored.sort(key=lambda t: -t[1])
        return scored[:top_k]

    @staticmethod
    def _trigrams(term: str) -> list[str]:
        return [term[i:i + 3] for i in range(len(term) - 2)] if len(term) >= 3 else []

    def _strip_josa(self, word: str) -> str:
        for j in _JOSA:
            if word.endswith(j) and len(word) - len(j) >= 2:
                return word[: -len(j)]
        return word

    def keyword_search(self, query: str, top_k: int) -> list[tuple[int, float]]:
        """trigram FTS5(bm25) + 조사 제거 LIKE 저하. (chunk_id, 순위 기반 점수)."""
        terms = [t for t in _nfc(query).split() if t]
        tris: list[str] = []
        for t in terms:
            tris.extend(self._trigrams(t))
        tris = list(dict.fromkeys(tris))[:_MAX_TRIGRAMS]
        roots = list(dict.fromkeys(
            r for r in (self._strip_josa(t) for t in terms) if len(r) >= 2
        ))[:_MAX_ROOTS]

        ordered: list[int] = []
        if self._fts and tris:
            match = " OR ".join('"' + g.replace('"', '""') + '"' for g in tris)
            try:
                with self._lock:
                    rows = self.conn.execute(
                        "SELECT rowid FROM chunks_fts WHERE chunks_fts MATCH ? "
                        "ORDER BY bm25(chunks_fts) LIMIT ?",
                        (match, top_k),
                    ).fetchall()
                ordered.extend(int(r["rowid"]) for r in rows)
            except sqlite3.OperationalError:
                pass
        if roots and len(ordered) < top_k:
            where = " OR ".join(["content LIKE ?"] * len(roots))
            params = [f"%{n}%" for n in roots] + [top_k]
            with self._lock:
                rows = self.conn.execute(
                    f"SELECT id FROM chunks WHERE {where} LIMIT ?", params
                ).fetchall()
            for r in rows:
                if int(r["id"]) not in ordered:
                    ordered.append(int(r["id"]))
        # 순위 → 점수 (RRF 합산용이라 절대값 의미는 없다)
        return [(cid, 1.0 / (rank + 1)) for rank, cid in enumerate(ordered[:top_k])]

    def fetch_chunks(self, chunk_ids: list[int]) -> dict[int, dict]:
        if not chunk_ids:
            return {}
        qmarks = ",".join("?" * len(chunk_ids))
        with self._lock:
            rows = self.conn.execute(
                f"SELECT c.id, c.seq, c.content, c.heading, c.file_id, f.path FROM chunks c "
                f"JOIN files f ON f.id = c.file_id WHERE c.id IN ({qmarks})",
                chunk_ids,
            ).fetchall()
        return {int(r["id"]): dict(r) for r in rows}

    def fetch_context(self, file_id: int, seq: int, window: int) -> list[dict]:
        """같은 파일에서 seq 앞뒤 window개 청크를 seq 순으로 돌려준다(이웃 컨텍스트 확장용)."""
        if window <= 0:
            return []
        with self._lock:
            rows = self.conn.execute(
                "SELECT seq, content, heading FROM chunks "
                "WHERE file_id = ? AND seq BETWEEN ? AND ? ORDER BY seq",
                (file_id, seq - window, seq + window),
            ).fetchall()
        return [dict(r) for r in rows]


_store: RagStore | None = None
_store_lock = threading.Lock()


def get_store() -> RagStore:
    global _store
    with _store_lock:
        if _store is None or _store.path != DB_PATH or _store.qdrant_path != QDRANT_PATH:
            _store = RagStore(DB_PATH, QDRANT_PATH)
        return _store


# ─────────────────────────────── 상태 요약 (양쪽 공용) ───────────────────────────────


def status_text() -> str:
    """인덱스 상태 요약 — rag_server의 rag_status 도구와 rag_indexer --status가 공용."""
    store = get_store()
    s = store.stats()
    embed_ok = _embed_available()
    lines = [
        f"인덱스 파일: {DB_PATH}",
        f"파일 {s['files']}개 / 청크 {s['chunks']}개 "
        f"(벡터 있는 청크 {s['with_vector']}개{', ' + str(s['dim']) + '차원' if s['dim'] else ''})",
        f"키워드 인덱스(FTS5): {'사용 가능' if s['fts'] else '없음 — LIKE로 저하'}",
        f"임베딩 서버({EMBED_URL}): {'연결됨' if embed_ok else '연결 안 됨 — 검색이 키워드 전용으로 동작'}",
        f"리랭커 서버({RERANK_URL}): {'연결됨' if _rerank_available() else '연결 안 됨 — RRF 순위를 그대로 사용'}",
        (
            f"벡터 저장소: Qdrant 로컬 ({store.qdrant_path})"
            if store.vec_kind == "qdrant"
            else "벡터 저장소: sqlite BLOB"
            + (f" — {store.vec_error}" if store.vec_error else "")
            + ("" if _np is not None else " (numpy 없음 — 순수 파이썬 코사인, 사내 미러에 있으면 설치 권장)")
        ),
        f"Word 읽기(COM): {'가능' if (COM_AVAILABLE and _document is not None) else '불가 — ' + (OFFICE_IMPORT_ERROR or 'pywin32 없음')}",
    ]
    if s["files"]:
        lines.append("")
        lines.append("인덱싱된 파일(최대 20개):")
        for f in store.list_files(limit=20):
            lines.append(f"  - {os.path.basename(f['path'])} (청크 {f['chunk_count']}개)")
    else:
        lines.append("")
        lines.append("인덱스가 비어 있습니다 — rag_indexer.py(run_rag_indexer.bat)로 문서 폴더를 인덱싱하세요.")
    return "\n".join(lines)
