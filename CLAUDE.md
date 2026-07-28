# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

두 층으로 이루어져 있고, **둘의 성격이 완전히 다르다**.

1. **반입물 (repo가 추적하는 것)** — `mcp_server/`의 MCP 서버들과 그 실행용 `run_*.bat`, `llm_studio/`(앱 + `serve_llm.py`), RAG 학습 문서 투입 폴더 `rag_docs/`, 루트의 `install_requirements.bat`. 사내 폐쇄망에서 실제로 돌릴 코드다. 여기가 이 저장소의 본체다. (폴더 규약: 서버 코드와 실행 bat은 `mcp_server/`에, LLM 서빙은 `llm_studio/`에, 오프라인 설치와 공용 `venv/`만 루트에 둔다.)
2. **강의 자료 (`Examples/`, `.gitignore`로 제외됨)** — LangChain/LangGraph 한국어 코스("AITF")의 랩 노트북과 봇 하니스. **git에 올라가지 않으므로 clone한 곳에는 존재하지 않는다.** 개발 PC에만 있는 참고 자료이고, 루트 코드가 여기에 의존하지 않는다.

`Examples/`를 수정하는 작업은 커밋되지 않는다는 점을 항상 염두에 둘 것.

## 배포 경로 (이 저장소의 존재 이유)

```
개인 PC (개발·커밋)  →  GitHub  →  회사 PC (clone)  →  폐쇄망 (실행)
```

폐쇄망에는 **인터넷이 없고**, 사내 PyPI 미러만 있다. 이 제약이 아래 모든 규칙의 근거다.

- **인터넷을 전제한 코드를 새로 넣지 말 것.** 런타임에 나가는 HTTP는 전부 localhost여야 한다. 패키지·모델·바이너리를 실행 중에 내려받는 코드는 폐쇄망에서 무조건 죽는다.
- **새 의존성은 사내 미러에 있는지 먼저 의심할 것.** 추가하면 반드시 `requirements.txt`에 명시하고, 없을 때 어떻게 되는지(우아한 저하 or 실패)를 주석에 적는다. `fastmcp`는 특히 위험한 축에 속한다 (개발 PC 검증 버전 3.4.4).
- **대용량 파일은 git에 넣지 않는다.** `.gguf` 모델, `llama-server.exe`+CUDA DLL, PyInstaller 산출물은 `.gitignore`로 막혀 있다. GitHub는 파일당 100MB 제한이 있고, 이것들은 USB 등 별도 매체로 옮긴다.
- 폐쇄망에서는 소스 직접 실행과 exe 설치 **두 경로를 모두 지원한다** (아래 llm_studio 참고).

## 구성 요소

### `mcp_server/` — MCP 서버 모음

MCP 서버(office/outlook/catia/rag/ansys/pdf)와 스모크 테스트, **그리고 그 실행용 bat이 전부 이 폴더에 있다**. `run_office_server.bat`·`run_outlook_server.bat`·`run_catia_server.bat`·`run_rag_server.bat`·`run_ansys_server.bat`·`run_pdf_server.bat`이 각각을 띄우는 실행 파일이다 — **인자 없이(더블클릭) 실행하면 HTTP 트랜스포트**로 뜨고, 인자를 주면 그대로 전달한다(`run_office_server.bat --transport stdio`). RAG 인덱스 구성 CLI는 `run_rag_indexer.bat`이 따로 있다. bat들은 자기 폴더로 `cd`한 뒤 같은 폴더의 `*.py`를 부르고, 파이썬은 **루트의 공용 `venv/`를 `..\venv\Scripts\python.exe`로** 찾는다(없으면 시스템 python) — 이 상대경로(`..\venv`)를 깨지 말 것. `rag_core.py`는 `office_server.py`를, `test_outlook.py`는 `outlook_server.py`를 같은 폴더에서 import하므로 **파이썬 파일을 폴더 밖으로 따로 옮기면 깨진다**(RAG 코드가 mcp_server를 못 떠나는 이유 — 학습에 넣을 문서만 루트 `rag_docs/`에 둔다). bat 파일은 **CRLF 줄바꿈 + 영어 ASCII만** 유지할 것 — LF로 저장하면 cmd가 줄 경계를 잘못 잘라 주석 조각을 명령으로 실행하고, 한글은 콘솔 코드페이지(cp949)와 파일 인코딩(UTF-8)이 어긋나 깨진다.

### `mcp_server/office_server.py` / `outlook_server.py` — COM 기반 MCP 서버

pywin32(COM)로 **이미 로그인·실행 중인** Office/Outlook을 직접 조종하는 MCP 서버. 파일을 파싱하는 게 아니라 앱 자체를 붙잡으므로, 화면에 열려 있는 저장 전 문서나 현재 사용자의 사서함을 그대로 읽는다.

공통 규약:
- **트랜스포트 3종** — `--transport stdio`(기본, Claude 등 로컬 클라이언트가 프로세스를 직접 실행), `http`(n8n 등이 URL로 접속), `sse`(Streamable HTTP를 못 쓰는 구버전 n8n MCP 노드용). 포트는 office 8087 / outlook 8088.
- **반드시 사용자가 로그인한 그 세션에서 실행해야 한다.** COM 특성상 서비스나 다른 세션에서 띄우면 열린 문서/사서함이 보이지 않는다. 이건 우회할 수 없는 제약이다.
- **우아한 저하** — pywin32가 없으면(`COM_AVAILABLE=False`) 서버는 정상적으로 뜨고 모든 도구가 실패 사유를 담은 안내 메시지를 반환한다. import 에러로 죽지 않는다.
- `path=""` → 지금 활성화된 문서, `path` 지정 → 열려 있으면 그 세션, 아니면 백그라운드에서 읽기 전용으로 열었다 닫는다. 암호 걸린 문서는 `password` 인자로 넘긴다 (대화상자 대신 오류 메시지로 물러선다).

`office_server.py`는 Word/PowerPoint **읽기 전용** + Excel 쓰기(3티어)다: 🟢 읽기(read_* 등 18개) / 🟡 메모리 수정(`write_excel_cell`·`write_excel_range` — **사용자 세션에 열려 있는** 통합문서만, 저장 안 함. COM 수정은 Excel Ctrl+Z에 안 쌓여서 이전 값을 응답으로 돌려준다) / 🔴 디스크 기록(`save_workbook` — confirm 게이트). 쓰기 도구는 백그라운드 읽기 전용 인스턴스(`_document`)를 쓰지 않고 `_writable_workbook`으로 열린 문서만 잡는다 — 이 구분을 깨지 말 것.

`outlook_server.py`는 쓰기가 가능해서 **3티어 안전 등급**을 따른다:
- 🟢 읽기 — 목록/검색/상세/첨부 저장/일정·연락처·작업 조회
- 🟡 로컬 생성(비파괴) — 초안·일정·연락처·작업 '만들기'만
- 🔴 외부 발송·파괴 — `send_email`, `send_draft`, `respond_message(send=True)`, `move_message`, `delete_message`, `create_meeting`, `respond_meeting`

**🔴 도구는 `confirm=True` 없이는 실행되지 않는다.** confirm 없이 부르면 "누구에게/무슨 제목/무슨 동작"을 요약한 프리뷰만 돌려준다(`_confirm_preview`). 이 게이트를 새 파괴적 도구에도 반드시 똑같이 적용할 것. 클라이언트 쪽에서 `HumanInTheLoopMiddleware`의 `INTERRUPT_ON`에 같은 이름을 올리면 이중 안전장치가 된다.

그 외 Outlook 고유 규약:
- **항목 참조는 EntryID 문자열**로 한다. 목록/검색이 `entry_id`를 주고 상세/후속 도구가 그걸 받는다.
- **푸시 트리거가 없다** (MCP는 요청-응답). 새 메일은 `poll_new_mail`로 폴링하고, 반환된 `checkpoint`를 다음 호출의 `since`로 넘겨 중복 없이 이어간다. n8n Schedule 트리거와 맞물리는 지점.
- Outlook의 Programmatic Access 경고창이 뜨면 COM 호출이 그 앞에서 멈춘다(회사 GPO에 좌우). 민감 속성은 우회 조회를 시도하고, 실패하면 대화상자를 띄우는 대신 빈 값/안내로 물러선다.

`test_outlook.py` — outlook_server 도구들을 실제 Outlook에 대고 한 번씩 호출하는 수동 스모크 테스트. **되돌리기 어려운 동작은 절대 실행하지 않는다** (어디에서도 `confirm=True`를 넘기지 않고, 🔴 도구는 프리뷰 경로만 확인한다). `--create` 플래그를 줘야 🟡 로컬 생성을 시도하고, 만든 항목은 곧바로 지운편지함으로 정리한다. 이 원칙을 깨는 수정을 하지 말 것.

### `mcp_server/rag_core.py` · `rag_indexer.py` · `rag_server.py` · `rag_llama.py` — 문서 RAG (구성·서빙 분리)

폴더의 문서를 청킹·임베딩해 인덱싱하고 하이브리드 검색(벡터+FTS5 trigram, RRF 융합)을 제공한다. **인덱싱 대상은 Word(.docx/.doc)·PowerPoint(.pptx/.ppt)·PDF(.pdf)** — `rag_core._extract_text`가 확장자로 분기한다(Word/PPT는 office COM `_document`, PDF는 pdf_server의 저하 체인 재사용). 확장자 그룹은 `WORD_PATTERNS`/`PPT_PATTERNS`/`PDF_PATTERNS`(합쳐서 `DOC_PATTERNS`). **구성과 서빙이 파일로 분리돼 있고, 이 구분을 유지할 것**:

- `rag_core.py` — 공용 코어(문서 읽기·청킹·임베딩·저장소). 실행 파일 아님. 설정(DB_PATH 등)은 CLI가 덮어쓰므로 다른 모듈에서는 `core.DB_PATH`처럼 **매번 속성으로** 읽는다(from-import 복사 금지).
- `rag_indexer.py` — **구성 CLI** (🟡 인덱싱 / 🔴 `--clear`는 `--yes` 없이 프리뷰만). `run_rag_indexer.bat`.
- `rag_server.py` — **서빙 MCP** (🟢 search_docs/rag_status **읽기 전용** — 모델이 인덱스를 못 건드린다). `run_rag_server.bat`, stdio 기본, http/sse는 :8090.
- `rag_llama.py` — **llama-server 런처 공용 모듈**(임베딩·리랭커). 모델/바이너리 탐색, `LlamaServer`(start/stop, 우아한 저하), 서빙용 백그라운드 기동+atexit 정리. 인덱서·서버가 함께 쓴다 — llama-server 로직이 흩어지지 않게 여기 한 곳에 모은다.

저장: 청크 본문·키워드 인덱스는 sqlite(rag_index.db), 벡터는 **Qdrant 로컬(파일) 모드**(rag_vectors/ — 서버 프로세스 없음, qdrant-client는 사내 미러 등록됨). qdrant-client가 없으면 sqlite BLOB 벡터로 우아하게 저하한다. ⚠ Qdrant 로컬은 단일 프로세스 잠금 — **rag_indexer는 서빙이 잠금을 쥐고 있으면 시작을 거부한다(exit 2)**. 조용히 sqlite로 저하해 인덱싱하면 서빙과 다른 저장소에 벡터가 쌓여 검색이 어긋나기 때문이다. 인덱싱할 때는 서빙을 잠시 내릴 것. llm_studio에는 데이터 폴더의 `mcp_servers.json`에 rag_server 등록만으로 장착된다(코드 변경 없음, 도구는 `docs__search_docs` 등).

- **Word/PPT 읽기는 office_server의 `_document`를 import해 재사용**한다 — COM이 여는 것이라 사내 DRM 문서도 읽힌다. 같은 제약(사용자 세션, Windows+Office)을 물려받는다. **PDF는 pdf_server의 `_extract`(direct→word_com→reader_print 체인)를 import해 재사용** — pdf_server를 못 불러오면 PDF만 실패로 집계하고 나머지는 계속(우아한 저하). PPT는 슬라이드별로 제목·본문·표·발표자 노트를 뽑고 '슬라이드 N' 머리말을 붙인다(노트는 페이지번호 자리표시자가 딸려오지 않게 **본문(body) 자리표시자만** 읽는다).
- **임베딩은 llama-server `--embeddings`**(기본 `http://127.0.0.1:8001/v1`, EmbeddingGemma 등 GGUF를 CPU `-ngl 0` 상주 권장). **서버가 없으면 키워드 인덱스만 만들고 검색도 키워드 전용으로 우아하게 저하** — 나중에 서버를 켜고 `--reindex`로 돌리면 벡터가 붙는다. 이 저하 경로 덕에 임베딩 모델 반입 전에도 개발·검증이 가능하다.
- **임베딩/리랭커 서버 자동 기동** (`rag_llama.LlamaServer`, 임베딩·리랭커 공용): 서버가 안 떠 있으면 반입 폴더(**임베딩 `rag_embed/`**, **리랭커 `rag_rerank/`**)의 유일한 `.gguf`를 `llama-server`로 띄우고 **우리가 띄운 프로세스만** 종료한다. 모델/실행파일(`LLAMA_SERVER_BIN` > `rag_embed`/`rag_rerank` 하위 재귀 > PATH)을 못 찾으면 우아하게 저하(임베딩 없음=키워드 전용, 리랭커 없음=RRF 순서). 배치는 `-b/-ub`를 `-c`(임베딩 `RAG_EMBED_CTX`·리랭커 `RAG_RERANK_CTX`, 기본 4096)와 같게 준다 — 안 그러면 기본 ubatch 512에 긴 청크가 넘쳐 요청마다 실패한다. CUDA는 exe 폴더+반입 폴더의 `cudart*`를 PATH에 덧대 폴더 병합 없이 로드한다. 반입 폴더는 gitignore라 README만 추적.
  - **인덱서**(`rag_indexer`): 인덱싱 동안만 임베딩 서버를 띄웠다 내린다(블로킹, `with`). `--no-embed-server`/`--embed-model`/`--server-bin`.
  - **서빙**(`rag_server`): 서버가 뜰 때 임베딩(:8001)·리랭커(:8002)를 **백그라운드로** 띄우고(모델 로드가 MCP 핸드셰이크를 막지 않게) 서버가 내려갈 때 함께 종료(atexit). 준비 전 검색은 우아하게 저하했다가 준비되면 자동으로 하이브리드·리랭크가 붙는다. `--no-embed-server`/`--no-rerank-server`/`--embed-model`/`--rerank-model`/`--server-bin`. **검색 시점에도 임베딩 서버가 필요**하다(질의 임베딩) — 그래서 인덱서뿐 아니라 서빙도 자동 기동한다. 이 경로는 Qdrant 단일 프로세스 잠금과 무관(별개 프로세스).
  - **리랭커**는 하이브리드 상위 후보를 질의-문서 관련도로 재정렬(`_rerank`, llama-server `--reranking` = `/v1/rerank`). 특히 한↔영 **교차언어** 상위 정렬을 개선한다. 없으면 RRF 순서 그대로(무해). 다국어 크로스인코더(BGE-reranker-v2-m3 등) 권장.
- 임베딩 모델을 바꿔 벡터 차원이 달라지면 Qdrant 컬렉션을 자동 재생성한다(stderr 경고) — 이후 전체 `--reindex` 필요.
- 키워드 검색은 llm_studio `memory.py`와 같은 패턴(FTS5 trigram + 조사 제거 LIKE 저하)이다 — 한쪽 휴리스틱을 고치면 다른 쪽도 확인할 것.

### `mcp_server/ansys_server.py` — ANSYS MAPDL 열해석 MCP 서버

PyMAPDL(`ansys-mapdl-core`, gRPC)로 MAPDL을 조종해 열해석(정상상태·과도)을 수행한다. COM이 아니므로 **사용자 세션 제약이 없고** 원격 인스턴스 접속도 된다. 대신 **ANSYS 본체 + 라이선스**가 필요하고, 서버가 MAPDL 세션 하나를 전역으로 상주시킨다(도구 호출은 락으로 직렬화).

- 워크플로: `launch_ansys` → 형상(create_block/cylinder) → `set_thermal_element` → `define_thermal_material`(과도면 밀도·비열 필수) → `mesh_model` → 경계조건(apply_temperature/convection/heat_flux/heat_generation — 면 번호는 `list_areas`) → `solve_steady`/`solve_transient` → `result_*`/`capture_plot`.
- 3티어: 🟢 상태·결과 조회 / 🟡 세션·모델 구축·solve·`run_apdl`(단 /CLEAR·/EXIT는 차단) / 🔴 `clear_model`·`shutdown_ansys`(confirm 게이트).
- `capture_plot`은 pyvista 없이 **MAPDL 자체 렌더러(/SHOW,PNG)**를 쓴다 — 폐쇄망 의도적 선택. 기존 파일은 덮어쓰지 않는다.
- ⚠ 개발 PC에 ANSYS가 없어 APDL 시퀀스·*GET 조회는 실기 미검증 — `⚠ 실기 검증 대상` 주석 위치를 실기에서 확인할 것. `ansys-mapdl-core`는 사내 미러 등록 여부 미확인이라 requirements.txt에 주석으로만 있다.

### `mcp_server/pdf_server.py` — DRM PDF 텍스트 추출 MCP 서버

사내 보안프로그램(DRM)이 암호화한 PDF의 **텍스트를 읽어오는** 읽기 전용(🟢) 서버. 핵심 전제: DRM은 **인증된 SW에만 실시간 복호화**를 해주므로(Python `open`·cmd `copy`는 둘 다 암호화 바이트만 읽힘 — 실측 확인됨), DRM을 우회하는 게 아니라 **인증 앱을 통과**해야 한다. office_server가 Word COM으로 DRM Word 문서를 읽는 것과 같은 원리다. (notepad로 읽어 재구성하는 접근은 불가 — 일반 프로세스엔 복호화를 안 해주고 PDF는 바이너리라 텍스트로 읽으면 손상.)

- **추출 백엔드 3종을 순서대로 시도(우아한 저하 체인)**: ① `direct` — 앞부분이 `%PDF`면(DRM 미적용) pypdf로 바로 추출, 암호화면 조용히 다음으로 ② `word_com` — Word를 백그라운드(DispatchEx)로 띄워 PDF를 열고(Word가 자동 변환) `doc.Content.Text` 추출. **DRM PDF의 실질 경로.** office_server와 같은 제약(사용자 세션·Windows+Word) 상속 ③ `reader_print` — (실험적) Acrobat Reader로 'Microsoft Print to PDF' 인쇄 후 pypdf. 조용한 파일 출력이 환경에 의존해 신뢰도 낮음. 어느 것도 실패하면 예외 대신 안내 문자열.
- 도구(전부 🟢): `read_pdf_text(path, pages, max_chars)`, `read_pdf_metadata(path)`, `pdf_status(path)`. `pypdf`/`pywin32`가 없어도 import에서 죽지 않고 도구가 안내로 저하. http/sse는 :8092.
- **word_com은 hang-safe**다: Word의 'PDF를 편집 가능한 문서로 변환' 확인창은 `DisplayAlerts=0`으로 안 꺼지므로(개발 PC 재현), Open을 데몬 스레드에서 돌리고 워치독이 그 대화상자를 자동 확인하며 `WORD_TIMEOUT`(기본 90초) 초과 시 **우리가 띄운 Word PID만**(생성 전후 차집합) taskkill한다 — 사용자 Word는 건드리지 않고, 막혀도 MCP 서버가 얼지 않는다.
- ⚠ 실기 검증 대상: DRM이 **Word.exe에 .pdf 복호화까지 허용하는지**(확장자 스코프 DRM이면 막힐 수 있음), 대화상자 자동 확인이 실기에서 실제로 통하는지. **서버 없이 `python mcp_server\pdf_server.py --probe <PDF경로>`로 각 백엔드를 진단**할 것. 개발 PC엔 실제 DRM이 없어(nProtect만 상주) word_com의 성공 여부는 사내 PC에서만 확정된다. `pypdf`는 `llm_studio`가 이미 쓰던 것을 루트 requirements.txt에 추가했다.

### `llm_studio/serve_llm.py` — 헤드리스 LLM 서빙

로컬 GGUF 모델을 llama.cpp의 `llama-server`로 띄워 OpenAI 호환 API(`/v1/chat/completions`)를 여는 CLI 스크립트. LangChain·n8n·HTML 페이지 등이 `base_url`만 바꿔 붙는 용도다. 표준 라이브러리만 쓰므로 pip 의존성이 없고, 대신 `llama-server` 실행 파일과 `.gguf`를 별도 반입해야 한다. (LLM 서빙 관련 코드를 한곳에 모으려고 앱과 같은 `llm_studio/`에 둔다.)

같은 폴더의 `llm_studio/server/llama_proc.py`와 **의도적으로 공존한다** (후자가 전자의 로직을 앱 내장 클래스로 옮긴 것). 갈라진 지점:

| | `serve_llm.py` | `llama_proc.py` |
|---|---|---|
| 형태 | CLI 스크립트 | 앱 내장 클래스 (`start`/`stop`/`restart`) |
| 바인딩 | `--host`로 `0.0.0.0` 공개 가능 | 항상 `127.0.0.1` (외부 공개는 UI 서버 담당) |
| **`--jinja`** | **없음 → 도구 호출 불가** | 항상 켬 |
| 동시 슬롯 | `--parallel`(`-np`) 지원 | 없음 (슬롯 1, 순차 처리) |

**`--jinja` 차이가 중요하다.** 이 플래그가 chat template 기반 function calling을 켠다. `serve_llm.py`로 띄운 서버에 도구 호출을 붙이면 동작하지 않으므로, 도구가 필요하면 `--extra --jinja`로 넘기거나 llm_studio를 쓸 것.

Gemma 3 계열 권장 샘플링(`GEMMA_SAMPLING`: temp 1.0 / top-p 0.95 / top-k 64 / min-p 0.0)은 양쪽에 중복 정의돼 있다. 직접 서빙에서는 이 값이 응답 품질을 좌우하므로 서버 기본값에 맡기지 않는다 — 한쪽을 바꾸면 다른 쪽도 확인할 것.

### `llm_studio/` — 폐쇄망용 올인원 로컬 LLM 앱

FastAPI 서버 + 브라우저 채팅 UI + llama-server 프로세스 관리를 하나로 묶은 앱. 스트리밍 채팅, 대화 기록, 파일 첨부, MCP 도구 연결, 외부 LLM(OpenAI 호환) 전환을 지원한다. 자세한 건 `llm_studio/README.md` (폐쇄망 반입 체크리스트 포함).

- **저장 위치 원칙**: 대화기록·설정·API 키·첨부는 **전부 서버 쪽 데이터 폴더**(`C:\ProgramData\LocalLLMStudio`, 권한 없으면 `%LOCALAPPDATA%`로 폴백)에 파일로 저장한다. **브라우저 저장소(localStorage/쿠키/IndexedDB)는 일절 쓰지 않는다** — 보안 프로그램이 브라우저 데이터를 지워도 아무것도 잃지 않게 하기 위한 의도적 설계다. 이 원칙을 깨지 말 것.
- **프로젝트(프롬프트·기억 격리)** (`server/projects.py`, claude Projects 스타일): **폴더 = 프로젝트**. `ConversationStore`/`MemoryStore`가 이미 디렉터리 단위라, 프로젝트마다 `projects/<id>/`(project.json + conversations/ + memory.db)를 주고 그 안에 대화·메모리를 담는다. 루트의 `conversations/`·`memory.db`는 **"기본" 공간**(프로젝트 없음)으로 그대로 산다 — 기존 데이터 마이그레이션 0, 프로젝트는 순수 추가. `ProjectManager`가 프로젝트별 스토어를 지연 생성·캐시하고, `stores(pid)`는 pid가 falsy·미존재면 기본 공간으로 폴백한다(chat이 잘못된 pid로 안 죽게). **프롬프트는 대체**(프로젝트 프롬프트가 있으면 그것만, 비면 전역 `system_prompt`로 폴백 — `base_prompt`). **메모리는 완전 격리**(회상·자동저장·비우기가 전부 그 프로젝트 memory.db만 대상). API·요청에 `project_id`를 실어 라우팅한다(`/api/chat`, `/api/conversations*`, `/api/memory*`에 `project_id` 쿼리, `/api/projects` CRUD). UI는 사이드바 프로젝트 선택기 + 설정 모달(이름·프롬프트·이 프로젝트 기억 비우기·삭제). **이 격리 규약을 깨지 말 것** — `_build_system`/`_maybe_autosummarize`는 state.memory/state.store가 아니라 **요청에서 해석한 프로젝트 store/memory**를 받는다. (프로젝트 삭제는 폴더째 rmtree라 `MemoryStore.close()`로 sqlite 잠금을 먼저 푼다.)
- **실행 경로 두 가지 모두 유지한다**:
  - 소스 직접 실행 — `pip install -r llm_studio/requirements.txt && python app.py` (`--mock`으로 모델 없이 UI만 확인 가능). 폐쇄망에 미러가 있으니 이쪽이 기본.
  - exe 배포 — `build_exe.bat`(PyInstaller) → `installer.iss`(Inno Setup) → `Setup.exe` 하나로 앱+llama-server 반입. 빌드는 개인 PC에서 한다.
- **MCP 클라이언트** (`server/mcp_client.py`) — 데이터 폴더의 `mcp_servers.json`을 Claude Desktop과 같은 `mcpServers` 규격으로 읽는다. `url`→streamable_http, `command`→stdio. 서버마다 전용 워커 태스크를 두고 큐로 요청을 전달하는데, 이건 MCP 세션을 **열었던 태스크에서 닫아야 한다는 anyio cancel scope 제약** 때문이다 — 구조를 단순화하려다 이 제약을 깨지 말 것. 연결 실패한 서버는 비활성 표시만 하고 나머지로 계속 동작한다. 도구 이름은 `<서버이름>__<도구이름>`으로 모델에 노출된다.
- **모델 서빙은 앱 시작과 분리돼 있다.** `app.py`는 llama-server를 자동으로 띄우지 않고 **유휴 상태로 뜬다**. 사용자가 UI(헤더 아래 셋업 바 / 설정 → 로컬 LLM 서버)에서 GGUF와 옵션을 골라 `POST /api/server/start`로 서빙을 켜고 `/stop`·`/restart`로 관리한다. `state.mock`은 이제 **오직 `--mock`**일 때만 참이다(개발용 canned 응답) — "모델 없음"과 혼용하지 말 것. 로컬을 골랐는데 서빙 중이 아니면 `/api/chat`이 실제 호출 대신 안내 error 이벤트를 흘린다. `config.autostart_local`(기본 False)을 켜거나 `--llama-url`을 주면 시작 시 자동 서빙하되, **실패해도 목 모드로 떨어지지 않고 유휴로 뜬다**.
- **위험 도구 승인 게이트** (`server/approvals.py`) — 모델이 `confirm=true` 인자로 도구를 부르거나 config `approval_tools`에 오른 도구를 부르면, 실행 전에 SSE `approval_request` 이벤트로 브라우저에 승인/거절 버튼을 띄우고 `POST /api/chat/approve` 응답을 기다린다(시간 초과·거절이면 실행하지 않고 그 사실을 도구 결과로 모델에 알림). 상태는 전부 RAM(asyncio Future) — 저장 위치 원칙과 무관. **MCP 서버 쪽 confirm 게이트와 이중 안전장치**로, 모델이 사용자에게 묻지 않고 스스로 confirm=true를 넣는 사고를 막는다. `approval_enabled`로 켜고 끈다(기본 켬).
- **기억은 두 층위다 — 섞지 말 것**: (1) **장기 기억**(`server/memory.py`, `memory.db` — 대화를 넘나드는 사실, 매 턴 관련 항목을 system에 회상 주입 + `_maybe_autosummarize`가 사실 자동 추출) / (2) **대화 문맥**(단기 — `conv["messages"]`, 매 턴 통째로 전송). 각각에 **비우기/압축**이 붙는다: 장기 기억 전체 비우기는 `DELETE /api/memory`(🔴 `_require_local` 게이트, `MemoryStore.clear()`) — 설정 → 대화·프롬프트 탭의 '전체 비우기' 버튼(**`project_id` 스코프 — 활성 프로젝트/기본 공간의 메모리만 비운다**; 위 프로젝트 절 참고). 대화 문맥의 **파괴적 압축**은 `POST /api/conversations/{id}/compact` — 채팅 헤더의 🗜 버튼. **압축 규약**: 최근 `compact_keep_recent_turns`(기본 4)턴만 원문으로 남기고 그 이전을 `agent.summarize_conversation`의 요약 한 덩어리로 치환해 `conv["summary"]`에 저장, 다음 요청부터 **기존 system 프롬프트에 합쳐** 주입한다(별도 2번째 system 메시지를 만들지 말 것 — Gemma 템플릿이 system 블록을 하나만 기대한다). 자르는 경계(`_split_for_compaction`)는 **반드시 user 메시지에서** — 안 그러면 `assistant.tool_calls`↔`role=tool` 쌍이 끊겨 전송 규격이 깨진다. 요약 생성이 실패하면 **원본을 건드리지 않는다**(우아한 저하). ⚠ `--mock`으로 띄워도 데이터 폴더는 실제 `C:\ProgramData\LocalLLMStudio`를 쓰므로, 비우기·압축을 검증할 땐 실데이터를 지우지 않게 격리된 데이터 폴더에서 할 것.
- llama-server는 `--jinja`로 실행돼 Gemma의 chat template 기반 함수 호출을 쓴다. Gemma는 공식 tool-use 학습이 약한 편이라 도구 호출 정확도가 모델에 따라 갈린다.
- 외부 LLM 프리셋(OpenAI/Anthropic/Gemini)은 폐쇄망에선 쓸 수 없다. 그 자리에 **사내 프록시/게이트웨이 주소를 등록하는 용도**로 남겨둔 것이다.

### `Examples/` — 강의 자료 (추적 안 됨)

LangChain/LangGraph 한국어 코스. 번호순 랩 노트북(01 tool calling → 11 multi-agent)과 그 결과물인 `.py` 모듈들(에이전트 팩토리 `build_agent.py`, Slack/Discord 봇 하니스, MCP 서버들, skills/memory/summarization 확장). 모든 코드·주석·프롬프트가 한국어다.

여기 있는 패턴 몇 가지가 루트 코드의 설계 배경이다:
- `build_agent.py`의 `VLLM_DECODING` ↔ `serve_llm.py`의 `GEMMA_SAMPLING` (자체 서빙은 샘플링을 명시한다는 같은 취지)
- `HumanInTheLoopMiddleware`의 `INTERRUPT_ON` ↔ `outlook_server.py`의 `confirm` 게이트
- `test_tools.py` ↔ `test_outlook.py` (수동 스모크 테스트 형식)

빌드 시스템·패키지 매니페스트·테스트 스위트가 없다. Jupyter로 대화식 실행하거나 스크립트를 직접 돌린다. 자세한 실행법은 `Examples/` 안의 노트북과 `.md` 파일들을 볼 것.

## 실행

```bash
# MCP 서버 (사용자 세션에서, Office/Outlook이 켜진 상태로)
pip install -r requirements.txt
python mcp_server\office_server.py                    # stdio
python mcp_server\office_server.py --transport http   # n8n용, :8087
python mcp_server\outlook_server.py --transport http  # n8n용, :8088 (catia :8089, rag :8090, ansys :8091, pdf :8092)
python mcp_server\test_outlook.py                     # 읽기 전용 스모크 테스트
mcp_server\run_office_server.bat                      # 위 http 실행의 더블클릭용 (서버별, mcp_server 안)
mcp_server\run_rag_indexer.bat ..\rag_docs            # RAG 인덱스 구성 (rag_docs 투입, 서빙은 내리고 실행)

# 로컬 LLM
python llm_studio\serve_llm.py --model C:/models/gemma-12b-it-qat.gguf
cd llm_studio && python app.py             # 유휴로 시작 — UI에서 모델을 골라 서빙 (run_app.bat 동일)
cd llm_studio && python app.py --mock      # 모델 없이 UI 확인 (목 응답)
```

린트/테스트 커맨드가 따로 없다. 변경은 해당 스크립트를 직접 돌려서 검증한다.

## 전반적 규약

- **모든 코드·주석·docstring·프롬프트는 한국어로 쓴다.** (예외: `.bat`은 전부 영어 ASCII — cmd 인코딩 문제로 한글이 깨진다. 위 mcp_server 절 참고.)
- **예외보다 우아한 저하** — 라이브러리 없음 → 안내 메시지 반환하는 스텁, MCP 서버 연결 실패 → 그 도구만 비활성, 확장 모듈 로드 실패 → 그것만 건너뜀. 선택적/외부 설정 때문에 프로세스가 죽는 경로를 만들지 말 것.
- **stdio 트랜스포트에서 stdout은 MCP 프로토콜 채널이다.** 로그는 반드시 stderr로 보낼 것 (`print(..., file=sys.stderr)`).
- **Windows 전제** — COM 서버들은 Windows + Office 없이는 의미가 없다. pywin32는 `sys_platform == "win32"` 마커로 걸려 있다.
- **requirements.txt는 UTF-8 BOM 포함으로 유지할 것** — 한국어 주석이 있는데 구버전 pip는 BOM이 없으면 로케일(cp949)로 읽어 `UnicodeDecodeError`가 난다. 파일을 다시 쓸 때 BOM을 떨어뜨리지 말 것. (파이썬 코드의 파일 I/O는 항상 `encoding="utf-8"` 명시 — 이미 전부 그렇게 돼 있다.)
- TLS/CA 번들: 외부 HTTPS를 호출하는 모듈은 네트워킹 라이브러리 import 전에 `SSL_CERT_FILE`을 `certifi.where()`로 설정한다 (사내망의 비표준 CA 체인 대응). 폐쇄망 코드에는 해당 없음.
