# rag_docs — RAG 학습(인덱싱)용 문서 투입 폴더

RAG로 검색하고 싶은 Word 문서(.docx/.doc)를 **이 폴더에 넣고** 인덱서를 돌린다.

```bat
mcp_server\run_rag_indexer.bat rag_docs            :: 증분 인덱싱
mcp_server\run_rag_indexer.bat rag_docs --reindex  :: 임베딩까지 전체 재구성
mcp_server\run_rag_indexer.bat --status            :: 인덱스 상태
```

- 인덱싱 코드는 `mcp_server/`에 있다(`rag_core.py`·`rag_indexer.py`·`rag_server.py`). `rag_core`가 같은 폴더의 `office_server`를 Word COM 재사용 용도로 import하므로 **코드는 mcp_server를 떠나지 않는다** — 이 폴더는 "학습에 넣을 문서"만 담는 데이터 폴더다.
- 인덱싱 산출물(`rag_index.db`, `rag_vectors/`)은 `mcp_server/`에 생기고 gitignore된다. 재인덱싱으로 언제든 복구된다.
- 서빙(`run_rag_server.bat`, 검색 MCP)이 Qdrant 잠금을 쥐고 있으면 인덱서가 시작을 거부한다(exit 2). **인덱싱할 때는 서빙을 잠시 내릴 것.**

> ⚠ **이 폴더에 넣은 실제 문서는 git에 커밋되지 않는다**(`.gitignore`가 `rag_docs/`의 README를 제외한 전부를 제외). 공개 저장소에 사내 문서가 딸려 올라가지 않게 하려는 의도적 설계다.
