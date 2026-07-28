# rag_embed — 임베딩 모델 반입 폴더

RAG 인덱서(`mcp_server/run_rag_indexer.bat`)가 벡터를 만들 때 쓰는 **임베딩 모델**과
(선택) **llama-server 실행 파일**을 여기에 둡니다. 큰 파일이라 git에는 올라가지 않고
(`.gitignore`), 이 README만 추적됩니다 — USB 등으로 반입하세요.

## 넣을 것

1. **임베딩 GGUF 1개** — 예: `embeddinggemma-300m-q8_0.gguf`
   - 이 폴더에 `.gguf`가 **딱 하나면 인덱서가 자동으로 선택**합니다.
   - 여러 개면 `run_rag_indexer.bat ..\rag_docs --embed-model rag_embed\<파일>.gguf`로 지정.
2. **`llama-server.exe`** (선택) — 이 폴더에 두면 인덱서가 자동으로 찾습니다.
   - 이미 `LLAMA_SERVER_BIN` 환경변수나 PATH에 있으면 안 넣어도 됩니다.
   - llm_studio가 쓰는 것과 같은 실행 파일입니다(llama.cpp Windows 빌드).

## 동작

`run_rag_indexer.bat ..\rag_docs`를 실행하면 인덱서가:

1. 임베딩 서버가 이미 떠 있는지 확인 → 있으면 그대로 사용
2. 없으면 이 폴더의 `.gguf`를 `llama-server --embeddings -ngl 0`(CPU)으로 **자동 기동**
3. `rag_docs/`의 Word·PPT·PDF를 읽어 인덱싱
4. 우리가 띄운 임베딩 서버를 **종료**하고 끝냄

모델이나 실행 파일이 없으면 죽지 않고 **키워드 전용 인덱스**만 만듭니다(나중에 모델을
넣고 `--reindex`하면 벡터가 붙습니다).

## 참고

- 포트는 `RAG_EMBED_URL`(기본 `http://127.0.0.1:8001/v1`)에서 가져옵니다.
- 모델 로드 대기 한도는 `RAG_LLAMA_LOAD_TIMEOUT`(기본 300초)로 조정.
- 자동 기동을 끄려면 `--no-embed-server` (이미 떠 있는 서버만 사용).
- **검색 서버(`run_rag_server.bat`)도 뜰 때 이 임베딩 서버를 자동 기동**합니다 — 검색 시점에
  질의를 임베딩해야 벡터·교차언어 검색이 되기 때문입니다(백그라운드 기동, 서버 내리면 함께 종료).
- `llama-server.exe`는 이 폴더(하위 재귀)와 `rag_rerank/`, PATH에서 함께 찾습니다 — 한 곳에만 두면 됩니다.
