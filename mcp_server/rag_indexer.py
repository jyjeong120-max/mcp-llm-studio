r"""rag_indexer.py

RAG 인덱스 **구성** CLI — 서빙(rag_server.py, MCP)과 의도적으로 분리돼 있다.
폴더/파일 인덱싱(🟡: 원본 문서는 읽기만 하고 인덱스 파일에만 씀)과
인덱스 삭제(🔴: --clear는 --yes 없이는 프리뷰만)를 담당한다.

구성·서빙 분리 이유:
    1. Qdrant 로컬(파일) 모드는 단일 프로세스 잠금 — 서빙 중 인덱싱하면 잠금을 못
       잡아 sqlite로 저하한 채 벡터가 서빙과 **다른 저장소에** 쌓인다(검색 어긋남).
       그래서 이 인덱서는 Qdrant를 못 잡으면 시작을 거부한다: 인덱싱 전에
       rag_server(MCP)를 내리고 실행할 것.
    2. 서빙 MCP는 읽기 전용(🟢)만 노출돼 모델이 인덱스를 건드릴 수 없다.

인덱싱 대상은 Word(.docx/.doc)·PowerPoint(.pptx/.ppt)·PDF(.pdf)다. Word/PPT는 office
COM으로, PDF는 pdf_server의 추출 체인(direct→word_com→reader_print)으로 읽는다.

임베딩 서버 자동 기동: 폴더/파일 인덱싱 시, 이미 임베딩 서버가 떠 있지 않으면 루트
rag_embed/ 폴더의 .gguf를 llama-server(--embeddings, CPU)로 잠깐 띄웠다가 인덱싱이
끝나면 종료한다(우리가 띄운 것만). 모델/실행파일이 없으면 키워드 전용으로 저하한다.

사용 (루트의 run_rag_indexer.bat 이 이 스크립트를 부른다):
    python rag_indexer.py C:\docs               # 폴더 인덱싱 (증분 — 변경된 파일만)
    python rag_indexer.py C:\docs --reindex     # 전부 다시 (임베딩 포함)
    python rag_indexer.py --file C:\docs\a.pptx # 파일 하나만 다시
    python rag_indexer.py --status              # 인덱스 상태 확인
    python rag_indexer.py --clear               # 삭제 프리뷰 (실행 안 함)
    python rag_indexer.py --clear --yes         # 인덱스 전체 삭제
    python rag_indexer.py C:\docs --no-embed-server        # 임베딩 서버 자동 기동 안 함
    python rag_indexer.py C:\docs --embed-model C:\m.gguf  # 임베딩 모델 직접 지정

Word/PPT는 office_server와 같은 제약: Windows + Office, 사용자가 로그인한 세션에서 실행.
"""

from __future__ import annotations

import argparse
import os
import sys
import time

import rag_core as core
import rag_llama
from rag_core import DOC_PATTERNS, EMBED_BATCH, OfficeError, RagError, RagStore


# ─────────────────────────────── 인덱싱 파이프라인 ───────────────────────────────


def _iter_doc_files(folder: str) -> list[str]:
    """폴더(재귀)에서 인덱싱 대상 파일 목록(Word/PPT/PDF). Office 임시 파일(~$)은 제외."""
    found = []
    for root, _dirs, names in os.walk(folder):
        for name in names:
            if name.startswith("~$"):
                continue
            if os.path.splitext(name)[1].lower() in DOC_PATTERNS:
                found.append(os.path.abspath(os.path.join(root, name)))
    return sorted(found)


def _index_one_file(store: RagStore, path: str, password: str = "",
                    reindex: bool = False, use_embed: bool | None = None) -> str:
    """파일 하나를 인덱싱한다. 반환: 한 줄 결과 요약."""
    st = os.stat(path)
    if not reindex and store.file_unchanged(path, st.st_mtime, st.st_size):
        return "변경 없음 — 건너뜀"
    text = core._extract_text(path, password)  # 확장자에 맞는 추출기(Word/PPT/PDF)로 분기
    chunks = core._chunk_text(text)  # [(섹션경로, 본문), ...]
    if not chunks:
        store.replace_file(path, st.st_mtime, st.st_size, [], None)
        return "본문 없음 — 청크 0개"
    # 임베딩에는 섹션 경로를 제목(title)으로 붙여, 하위 청크에도 상위 제목 맥락이 벡터에
    # 담기게 한다(본문엔 그 섹션 제목 줄만 있고 상위 대제목은 없을 수 있으므로).
    # EmbeddingGemma 등은 문서 프롬프트 포맷이 정해져 있어 core 헬퍼로 생성한다.
    embed_texts = [core._embed_doc_text(c, h) for h, c in chunks]
    vectors: list[list[float]] | None = None
    if use_embed is not False:
        vectors = []
        for i in range(0, len(embed_texts), EMBED_BATCH):
            batch = core._embed_texts(embed_texts[i:i + EMBED_BATCH], warn=True)
            if batch is None:
                vectors = None  # 서버 죽음/오류 → 이 파일은 키워드 전용으로 저장
                break
            vectors.extend(batch)
    store.replace_file(path, st.st_mtime, st.st_size, chunks, vectors)
    vec_note = f"벡터 {len(vectors)}개" if vectors else "벡터 없음(키워드 전용)"
    return f"청크 {len(chunks)}개, {vec_note}"


def index_folder(folder: str, reindex: bool = False, prune: bool = True, password: str = "") -> str:
    """폴더(하위 포함)의 문서(Word/PPT/PDF)를 모두 인덱싱하고 결과 요약을 돌려준다.

    원본 문서는 읽기만 하고 인덱스 파일에만 쓴다. 수정 시각·크기가 같은 파일은
    건너뛴다(증분). prune=True면 폴더에서 사라진 파일을 인덱스에서도 정리한다.
    """
    root = os.path.abspath(os.path.expanduser(folder))
    if not os.path.isdir(root):
        raise RagError(f"'{root}' 폴더가 없습니다. 경로를 확인하세요.")
    files = _iter_doc_files(root)
    if not files:
        exts = "/".join(DOC_PATTERNS)
        return f"'{root}' 아래에 인덱싱할 문서({exts})가 없습니다."
    # Word/PPT는 COM이 필요하다(PDF는 direct 백엔드로 COM 없이도 됨) — COM이 없으면
    # 해당 파일들은 개별적으로 실패 처리되고 나머지(PDF direct)는 계속 진행한다.
    if any(os.path.splitext(f)[1].lower() in core.WORD_PATTERNS + core.PPT_PATTERNS
           for f in files):
        core._require_word()

    store = core.get_store()
    embed_ok = core._embed_available()
    results: list[str] = []
    ok = skipped = failed = 0
    start = time.time()
    for path in files:
        try:
            note = _index_one_file(store, path, password, reindex, use_embed=embed_ok)
            if note.startswith("변경 없음"):
                skipped += 1
            else:
                ok += 1
                results.append(f"  - {os.path.basename(path)}: {note}")
        except (RagError, OfficeError) as e:
            failed += 1
            results.append(f"  ✗ {os.path.basename(path)}: {e}")
        except Exception as e:  # noqa: BLE001 — 한 파일 실패가 전체를 멈추지 않게
            failed += 1
            results.append(f"  ✗ {os.path.basename(path)}: {type(e).__name__}: {e}")
    pruned = store.remove_missing({os.path.abspath(f) for f in files}) if prune else 0

    s = store.stats()
    head = [
        f"인덱싱 완료: {root}  ({time.time() - start:.1f}초)",
        f"  처리 {ok}개 / 건너뜀(변경 없음) {skipped}개 / 실패 {failed}개"
        + (f" / 정리(삭제된 파일) {pruned}개" if pruned else ""),
        f"  임베딩: {'벡터 생성함' if embed_ok else '서버 없음 — 키워드 인덱스만 생성'}",
        f"  누적: 파일 {s['files']}개, 청크 {s['chunks']}개 (벡터 {s['with_vector']}개)",
    ]
    body = results[:100]
    if len(results) > 100:
        body.append(f"  … (이하 {len(results) - 100}개 생략)")
    return "\n".join(head + ([""] + body if body else []))


def index_file(path: str, password: str = "") -> str:
    """문서(Word/PPT/PDF) 하나를 (다시) 인덱싱하고 결과 요약을 돌려준다."""
    p = os.path.abspath(os.path.expanduser(path))
    if not os.path.isfile(p):
        raise RagError(f"'{p}' 파일이 없습니다. 경로를 확인하세요.")
    if os.path.splitext(p)[1].lower() not in DOC_PATTERNS:
        exts = "/".join(DOC_PATTERNS)
        raise RagError(f"인덱싱 대상 형식({exts})이 아닙니다: {os.path.basename(p)}")
    store = core.get_store()
    note = _index_one_file(store, p, password, reindex=True, use_embed=None)
    return f"인덱싱 완료: {os.path.basename(p)} — {note}"


def _require_qdrant_or_exit(store: RagStore) -> None:
    """qdrant-client가 있는데 백엔드를 못 열었다면(대개 서빙이 잠금 보유) 중단한다.

    저하한 채 인덱싱하면 벡터가 서빙(Qdrant)과 다른 곳(sqlite BLOB)에 쌓여
    검색이 어긋난다 — 조용한 불일치보다 명시적 거부가 안전하다.
    qdrant-client 자체가 없는 환경은 양쪽 다 sqlite라 일관되므로 그대로 진행한다.
    """
    if core.QDRANT_AVAILABLE and store.vec_kind != "qdrant":
        print(f"[중단] Qdrant 벡터 저장소를 열지 못했습니다: {store.vec_error}", file=sys.stderr)
        print(
            "rag_server(MCP 서빙)가 떠 있으면 종료한 뒤 다시 실행하세요. "
            "(저하한 채 인덱싱하면 서빙과 다른 저장소에 벡터가 쌓여 검색이 어긋납니다)",
            file=sys.stderr,
        )
        sys.exit(2)


# ─────────────────────────────── CLI ───────────────────────────────


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Word 문서 RAG 인덱스 구성 CLI (서빙은 rag_server.py)"
    )
    parser.add_argument("folder", nargs="?", help="인덱싱할 폴더 (하위 폴더 포함)")
    parser.add_argument("--file", metavar="DOCX", help="파일 하나만 (다시) 인덱싱")
    parser.add_argument("--reindex", action="store_true", help="변경 없어도 전부 다시 (임베딩 포함)")
    parser.add_argument("--no-prune", action="store_true", help="사라진 파일을 인덱스에서 정리하지 않음")
    parser.add_argument("--password", default="", help="문서 공통 열기 암호")
    parser.add_argument("--status", action="store_true", help="인덱스 상태만 출력하고 종료")
    parser.add_argument("--clear", action="store_true", help="인덱스 전체 삭제 (--yes 없으면 프리뷰만)")
    parser.add_argument("--yes", action="store_true", help="--clear를 실제로 실행")
    parser.add_argument("--db", default=None, help=f"인덱스 파일 경로 (기본 {core.DB_PATH})")
    parser.add_argument("--qdrant", default=None,
                        help=f"Qdrant 로컬 데이터 폴더 (기본 {core.QDRANT_PATH})")
    parser.add_argument("--embed-url", default=None,
                        help=f"임베딩 서버 /v1 베이스 URL (기본 {core.EMBED_URL})")
    parser.add_argument("--embed-model", default=None,
                        help="임베딩 GGUF 경로 (기본: 루트 rag_embed/의 유일한 .gguf 자동 선택)")
    parser.add_argument("--server-bin", default=None,
                        help="llama-server 실행 파일 (기본: LLAMA_SERVER_BIN > rag_embed/ > PATH)")
    parser.add_argument("--no-embed-server", action="store_true",
                        help="임베딩 서버를 자동 기동하지 않음 (이미 떠 있으면 그것만 사용)")
    args = parser.parse_args()

    if args.db:
        core.DB_PATH = os.path.abspath(os.path.expanduser(args.db))
    if args.qdrant:
        core.QDRANT_PATH = os.path.abspath(os.path.expanduser(args.qdrant))
    if args.embed_url:
        core.EMBED_URL = args.embed_url

    if core.pythoncom is not None:  # CLI 메인 스레드 COM 초기화 (Word 읽기용)
        core.pythoncom.CoInitialize()

    try:
        if args.status:
            print(core.status_text())
            return

        if args.clear:
            store = core.get_store()
            s = store.stats()
            if not args.yes:
                print(
                    f"⚠️ 프리뷰 — 아직 삭제하지 않았습니다: 인덱스 전체 삭제\n"
                    f"  인덱스 파일: {core.DB_PATH}\n"
                    f"  삭제 대상: 파일 {s['files']}개, 청크 {s['chunks']}개 (원본 문서는 안전)\n"
                    "실제로 삭제하려면 --yes 를 함께 지정하세요."
                )
                return
            _require_qdrant_or_exit(store)
            nf, nc = store.clear()
            print(f"인덱스를 비웠습니다 (파일 {nf}개, 청크 {nc}개 삭제). 원본 문서는 그대로입니다.")
            return

        if args.file or args.folder:
            _require_qdrant_or_exit(core.get_store())
            # 임베딩 서버를 자동 기동(rag_embed의 모델) → 인덱싱 → 종료. 모델/바이너리가
            # 없거나 --no-embed-server면 키워드 전용으로 저하한다(spawn=False면 재사용만).
            with rag_llama.make_embed_server(
                    args.embed_model, args.server_bin, spawn=not args.no_embed_server,
                    degrade_note="키워드 전용으로 인덱싱합니다"):
                if args.file:
                    print(index_file(args.file, password=args.password))
                else:
                    print(index_folder(args.folder, reindex=args.reindex,
                                       prune=not args.no_prune, password=args.password))
            return

        parser.print_help()
    except (RagError, OfficeError) as e:
        print(f"[오류] {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        # Qdrant 로컬 클라이언트를 살아 있을 때 명시적으로 닫는다 — 안 그러면 종료 중
        # __del__이 뒤늦게 돌며 "ImportError: sys.meta_path is None" 잡음을 남긴다.
        core.close_store()
        if core.pythoncom is not None:
            core.pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
