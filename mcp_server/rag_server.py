"""rag_server.py

RAG **서빙** MCP 서버 — 읽기 전용(🟢). 인덱스 구성(인덱싱·삭제)은 rag_indexer.py가
담당한다 (구성·서빙 분리 — 이유는 rag_indexer.py docstring 참고). 저장소·검색
로직은 rag_core.py에 있고 여기는 MCP 도구 껍데기만 둔다.

도구:
    search_docs — 하이브리드 검색(벡터+키워드, RRF 융합). 임베딩 서버가 없으면
                  키워드 전용으로 우아하게 저하한다.
    rag_status  — 인덱스 상태 조회 (파일/청크/벡터 수, 임베딩 서버 연결 여부).

사용:
    python rag_server.py                     # MCP 서버 (stdio)
    python rag_server.py --transport http    # n8n 등 네트워크용, :8090
    python rag_server.py --search "휴가 규정"  # 검색 시험 (CLI, MCP 없이)

llm_studio 장착 (코드 변경 불필요): 데이터 폴더의 mcp_servers.json에 등록만 하면
도구가 docs__search_docs 등으로 모델에 노출된다.
    {"mcpServers": {"docs": {"command": "python",
                             "args": ["C:\\경로\\mcp_server\\rag_server.py"]}}}

⚠ 이 서버가 떠 있는 동안 rag_indexer를 돌리면 Qdrant 로컬 잠금 때문에 인덱서가
시작을 거부한다 — 인덱싱할 때는 이 서버를 잠시 내릴 것. 반대로 인덱서가 도는 중에
이 서버를 띄우면 벡터 검색이 sqlite BLOB으로 저하한다(rag_status에 사유 표시).
"""

from __future__ import annotations

import argparse
import atexit
import os
import sys

from fastmcp import FastMCP

import rag_core as core
import rag_llama
from rag_core import MAX_RESULT_CHARS, RRF_K, RagError, rag_tool

mcp = FastMCP(
    name="docs",
    instructions=(
        "사내 Word 문서(RAG 인덱스)를 검색하는 MCP 서버입니다. 문서 내용에 관한 "
        "질문을 받으면 search_docs로 관련 청크를 찾아 근거로 답하세요(출처 파일명을 "
        "함께 알려주면 좋습니다). 인덱스가 비었거나 상태가 궁금하면 rag_status를 "
        "먼저 호출하세요. 이 서버는 읽기 전용입니다 — 인덱스 구성(문서 추가/삭제)은 "
        "관리자가 rag_indexer CLI로 합니다."
    ),
)


@mcp.tool()
@rag_tool
def rag_status() -> str:
    """RAG 인덱스 상태를 조회합니다 — 파일/청크 수, 벡터 유무, 임베딩 서버 연결. (🟢 읽기)

    검색이 이상하거나 인덱스가 비어 보일 때 가장 먼저 호출하세요.
    """
    return core.status_text()


@mcp.tool()
@rag_tool
def search_docs(query: str, top_k: int = 5) -> str:
    """인덱싱된 사내 문서에서 질의와 관련된 내용을 찾습니다. (🟢 읽기)

    벡터(의미) 검색과 키워드 검색을 함께 수행해 RRF로 합치고, 리랭커가 있으면 상위
    후보를 관련도로 재정렬합니다. 임베딩·리랭커 서버가 없으면 각각 우아하게 저하합니다
    (키워드 전용 / RRF 순서). 결과에는 출처 파일명·섹션 경로가 붙고, 매칭 청크의 같은
    섹션 이웃을 함께 보여줍니다 — 답변할 때 출처를 함께 알려주세요.

    Args:
        query: 찾을 내용(자연어 질문 그대로도 됩니다).
        top_k: 돌려줄 청크 수(기본 5, 최대 20).
    """
    q = core._nfc(query)
    if not q:
        raise RagError("query가 비어 있습니다. 찾을 내용을 지정하세요.")
    k = max(1, min(int(top_k), 20))
    store = core.get_store()
    if store.stats()["chunks"] == 0:
        return "인덱스가 비어 있습니다. rag_indexer.py(run_rag_indexer.bat)로 문서 폴더를 먼저 인덱싱하세요."

    # 두 경로에서 넉넉히(k*3) 뽑아 RRF로 합친다.
    pool = k * 3
    vec_hits: list[tuple[int, float]] = []
    qv = core._embed_texts([core._embed_query_text(q)])  # 질의 프리픽스(모델 포맷) 부착
    mode = "키워드 전용 (임베딩 서버 없음)"
    if qv:
        vec_hits = store.vector_search(qv[0], pool)
        mode = "하이브리드 (벡터+키워드)"
    kw_hits = store.keyword_search(q, pool)

    fused: dict[int, float] = {}
    for hits in (vec_hits, kw_hits):
        for rank, (cid, _score) in enumerate(hits):
            fused[cid] = fused.get(cid, 0.0) + 1.0 / (RRF_K + rank + 1)
    if not fused:
        return f"'{q}'와 관련된 내용을 찾지 못했습니다. 다른 표현으로 다시 시도해 보세요."

    # 리랭크 후보: 융합 상위 N개(k보다 넉넉히)를 뽑아 본문을 가져온다. 리랭커가 있으면
    # 질의-청크 관련도로 재정렬하고, 없으면 RRF 순서를 그대로 쓴다(우아한 저하).
    rerank_pool = max(k * 4, 20)
    ranked = sorted(fused.items(), key=lambda t: -t[1])[:rerank_pool]
    chunk_map = store.fetch_chunks([cid for cid, _ in ranked])
    present = [(cid, chunk_map[cid]) for cid, _ in ranked if cid in chunk_map]
    rr_docs = [
        (c["heading"] + "\n" + c["content"]) if c.get("heading") else c["content"]
        for _, c in present
    ]
    rr = core._rerank(q, rr_docs)
    if rr is not None:
        order = [present[idx] for idx, _s in rr if idx < len(present)]
        mode += " + 리랭크"
    else:
        order = present

    # 이웃 확장을 켜면 인접 청크가 같은 블록으로 겹칠 수 있다 — 이미 낸 블록과 같은
    # 내용은 건너뛰고 다음 후보로 채워 서로 다른 결과 k개를 유지한다.
    cap = core.MAX_CONTEXT_CHARS if core.CONTEXT_WINDOW > 0 else MAX_RESULT_CHARS
    out = [f"검색: {q}  [{mode}]", ""]
    seen: set[str] = set()
    i = 0
    for _cid, c in order:
        if i >= k:
            break
        content = _expand_context(store, c)
        if content in seen:
            continue
        seen.add(content)
        if len(content) > cap:
            content = content[:cap] + " …(생략)"
        i += 1
        # 섹션 경로(대제목 › 소제목)가 있으면 출처에 함께 표시한다(구버전 인덱스는 없음).
        heading = c.get("heading")
        loc = f" › {heading}" if heading else ""
        out.append(f"[{i}] {os.path.basename(c['path'])}{loc} (청크 {c['seq']})")
        out.append(content)
        out.append("")
    return "\n".join(out).rstrip()


def _expand_context(store, c: dict) -> str:
    """매칭 청크에 같은 섹션의 이웃 청크(앞뒤)를 붙여 근거가 청크 경계에서 잘리지 않게 한다.

    이웃은 같은 섹션 경로(heading)만 골라 seq 순으로 겹침을 제거하며 잇는다. 컨텍스트
    확장이 꺼져 있거나(CONTEXT_WINDOW=0) 이웃이 없으면 매칭 청크 본문만 돌려준다.
    """
    if core.CONTEXT_WINDOW <= 0:
        return c["content"]
    neigh = store.fetch_context(c["file_id"], c["seq"], core.CONTEXT_WINDOW)
    same = [
        n["content"] for n in neigh
        if (n.get("heading") or None) == (c.get("heading") or None)
    ]
    return core._merge_overlapping(same) if same else c["content"]


# ─────────────────────────────── CLI / 서버 기동 ───────────────────────────────


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Word 문서 RAG 검색 MCP 서버 (서빙 전용 — 인덱싱은 rag_indexer.py)"
    )
    parser.add_argument("--search", metavar="QUERY", help="검색을 시험하고 종료 (MCP 서버를 띄우지 않음)")
    parser.add_argument("--top-k", type=int, default=5, help="--search 결과 수 (기본 5)")
    parser.add_argument("--db", default=None, help=f"인덱스 파일 경로 (기본 {core.DB_PATH})")
    parser.add_argument("--qdrant", default=None,
                        help=f"Qdrant 로컬 데이터 폴더 (기본 {core.QDRANT_PATH})")
    parser.add_argument("--embed-url", default=None,
                        help=f"임베딩 서버 /v1 베이스 URL (기본 {core.EMBED_URL})")
    parser.add_argument("--rerank-url", default=None,
                        help=f"리랭커 서버 /v1 베이스 URL (기본 {core.RERANK_URL}, 없으면 리랭크 건너뜀)")
    # 서빙 시 임베딩(:8001)·리랭커(:8002) llama-server를 백그라운드로 자동 기동한다.
    # 모델(rag_embed/·rag_rerank/의 gguf)이나 실행파일이 없으면 조용히 저하한다.
    parser.add_argument("--no-embed-server", action="store_true",
                        help="임베딩 서버를 자동 기동하지 않음 (이미 떠 있으면 그것만 사용)")
    parser.add_argument("--no-rerank-server", action="store_true",
                        help="리랭커 서버를 자동 기동하지 않음 (이미 떠 있으면 그것만 사용)")
    parser.add_argument("--embed-model", default=None,
                        help="임베딩 GGUF 경로 (기본: rag_embed/의 유일한 .gguf 자동 선택)")
    parser.add_argument("--rerank-model", default=None,
                        help="리랭커 GGUF 경로 (기본: rag_rerank/의 유일한 .gguf 자동 선택)")
    parser.add_argument("--server-bin", default=None,
                        help="llama-server 실행 파일 (기본: LLAMA_SERVER_BIN > rag_embed/rag_rerank > PATH)")
    parser.add_argument(
        "--transport", choices=["stdio", "http", "sse"],
        default=os.getenv("RAG_MCP_TRANSPORT", "stdio"),
        help="stdio(기본): 로컬 클라이언트가 직접 실행. http/sse: n8n 등 네트워크 접속.",
    )
    parser.add_argument("--host", default=os.getenv("RAG_MCP_HOST", "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.getenv("RAG_MCP_PORT", "8090")))
    args = parser.parse_args()

    # 종료 시 Qdrant 로컬 클라이언트를 명시적으로 닫는다 — 안 그러면 종료 중 __del__이
    # 뒤늦게 돌며 "ImportError: sys.meta_path is None / msvcrt halted" 잡음을 남긴다.
    atexit.register(core.close_store)

    if args.db:
        core.DB_PATH = os.path.abspath(os.path.expanduser(args.db))
    if args.qdrant:
        core.QDRANT_PATH = os.path.abspath(os.path.expanduser(args.qdrant))
    if args.embed_url:
        core.EMBED_URL = args.embed_url
    if args.rerank_url:
        core.RERANK_URL = args.rerank_url

    if args.search:
        # CLI 모드: 도구 함수를 직접 호출한다 (rag_tool 데코레이터가 예외를 처리).
        # 단발 진단이라 자동 기동은 하지 않는다(이미 떠 있는 서버가 있으면 그대로 사용).
        print(search_docs(args.search, top_k=args.top_k))
        sys.exit(0)

    # 서빙 시작 전, 임베딩(:8001)·리랭커(:8002) llama-server를 백그라운드로 자동 기동한다.
    # 모델 로드(수십 초)가 MCP 핸드셰이크를 막지 않게 별도 스레드에서 띄우고, 프로세스 종료
    # 시 우리가 띄운 것만 정리한다(atexit). 준비 전 검색은 우아하게 저하했다가 준비되면
    # 자동으로 하이브리드·리랭크가 붙는다. 모델/실행파일이 없으면 조용히 저하한다.
    backends = []
    if not args.no_embed_server:
        backends.append(rag_llama.make_embed_server(
            args.embed_model, args.server_bin, degrade_note="키워드 전용으로 검색합니다"))
    if not args.no_rerank_server:
        backends.append(rag_llama.make_rerank_server(args.rerank_model, args.server_bin))
    if backends:
        rag_llama.start_in_background(backends)

    if args.transport in ("http", "sse"):
        path = "/mcp/" if args.transport == "http" else "/sse/"
        print(f"RAG MCP 서버 시작 ({args.transport}) — http://{args.host}:{args.port}{path}",
              file=sys.stderr)
        mcp.run(transport=args.transport, host=args.host, port=args.port)
    else:
        # stdio: stdout은 MCP 프로토콜 채널 — 로그는 stderr로.
        print("RAG MCP 서버 시작 (stdio)", file=sys.stderr)
        mcp.run(transport="stdio")
