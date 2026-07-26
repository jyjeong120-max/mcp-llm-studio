# LocalLLM Studio

자체 호스팅 로컬 LLM 채팅 앱 — 브라우저 UI + 스트리밍 + MCP 도구 호출 + 외부 API 전환을
하나로 묶은 올인원. 원래 인터넷 없는 폐쇄망 Windows PC용으로 만들었지만 로컬 어디서나 돈다.

## 빠른 시작 (Windows + Python 3.10+)

```bat
pip install -r requirements.txt
python app.py            :: 유휴로 시작 → 브라우저 UI에서 GGUF 모델을 골라 서빙
python app.py --mock     :: 모델 없이 UI만 둘러보기 (목 응답)
```

> **llama-server 실행 파일과 `.gguf` 모델은 이 저장소에 포함돼 있지 않다**(용량 때문).
> llama.cpp의 `llama-server`를 따로 받아 `llama/` 폴더에 두거나 설정에서 경로를 지정하고,
> GGUF 모델은 로컬 어디든 두고 UI에서 고르면 된다. 모델·바이너리 없이도 `--mock`으로 UI는 확인된다.

아래는 원래 폐쇄망 반입을 전제로 쓴 상세 문서다.

---

폐쇄망 Windows PC에서 로컬 LLM(Gemma 12B IT QAT 등 GGUF 모델)을 서빙하고,
브라우저 채팅 UI로 사용하는 올인원 앱.

- **앱 시작과 모델 서빙이 분리됨** — 앱은 아무 모델도 서빙하지 않은 유휴 상태로 뜨고,
  UI에서 GGUF 모델과 옵션을 골라 서빙을 시작/중지한다 (외부 API도 그대로 선택 가능)
- 스트리밍 채팅 (마크다운 렌더링, 응답 중단)
- 대화 기록 저장/불러오기/이름변경/삭제
- 파일 첨부 질의 (텍스트/PDF/DOCX)
- 설정 화면 (시스템 프롬프트, 생성 파라미터, 로컬 서버 설정)
- **MCP 도구 연결** — mcp_servers.json에 등록하면 모델이 도구를 호출
- **외부 LLM 연결** — OpenAI/Claude/Gemini 등 OpenAI 호환 API를 키로 등록하고
  화면 상단 드롭다운으로 로컬 모델과 전환

> **저장 위치 원칙**: 대화기록·설정·API 키·첨부파일은 전부 서버 쪽 데이터 폴더
> (`C:\ProgramData\LocalLLMStudio`)의 파일로 저장된다. 브라우저 저장소
> (localStorage/쿠키/IndexedDB)는 일절 쓰지 않으므로, 보안 프로그램이 브라우저
> 데이터를 지워도 아무 것도 잃지 않는다.

## 구조

```
LocalLLMStudio.exe (통합 런처)
  ├─ 웹서버 실행 (기본 127.0.0.1:8080 — UI + API + 도구 호출 루프)
  ├─ 브라우저 자동 오픈
  └─ (유휴로 시작) UI에서 [서빙 시작] → llama-server 실행 (127.0.0.1:8000, 내부 전용)
        · '앱 시작 시 자동 서빙'을 켜두면 마지막 설정으로 자동 시작

C:\ProgramData\LocalLLMStudio\    ← 인스톨러가 Users 쓰기 권한 부여
  ├─ models\            GGUF 모델을 여기에 넣는다
  ├─ conversations\     대화 기록 (JSON)
  ├─ uploads\           첨부 파일
  ├─ logs\              llama-server 로그
  ├─ config.json        설정
  └─ mcp_servers.json   MCP 서버 등록
```

## 개발 실행 (소스로)

```bash
cd llm_studio
pip install -r requirements.txt
python app.py --mock        # 모델 없이 UI 확인 (목 응답)
python app.py               # 유휴로 시작 — UI에서 모델을 골라 [서빙 시작]
python app.py --host 0.0.0.0  # 같은 망의 다른 PC에서 접속 허용
python app.py --llama-url http://127.0.0.1:8000  # 시작 시 따로 띄운 서버에 붙기
```

앱은 기본적으로 아무 로컬 모델도 서빙하지 않고 뜬다. 헤더 아래 셋업 바 또는
설정(⚙) → "로컬 LLM 서버"에서 GGUF를 고르고 [서빙 시작]을 누른다. 마지막 설정은
`config.json`에서 그대로 불러와 미리 채워진다. '앱 시작 시 자동 서빙'을 켜두면
다음 실행부터 그 설정으로 자동 시작한다(실패해도 유휴로 뜬다).

개발 중에는 ProgramData에 쓰기 권한이 없으면 `%LOCALAPPDATA%\LocalLLMStudio`를 자동 사용한다.

## exe 빌드 → 인스톨러

```bash
pip install pyinstaller
# 1) llama.cpp 릴리스에서 llama-bXXXX-bin-win-cuda-x64.zip을 받아
#    llm_studio\llama\ 폴더에 풀어둔다 (llama-server.exe + DLL들)
build_exe.bat               # → dist\LocalLLMStudio\
# 2) Inno Setup 6 설치 후 installer.iss 컴파일 → Output\LocalLLMStudio-Setup-1.0.0.exe
```

## 폐쇄망 반입 체크리스트

| 반입물 | 비고 |
|---|---|
| `LocalLLMStudio-Setup-x.x.x.exe` | 인스톨러 하나에 앱+llama-server 포함 |
| Gemma 12B IT QAT `.gguf` | 설치 후 `C:\ProgramData\LocalLLMStudio\models\`에 복사 |
| NVIDIA 드라이버 | 서버 PC에 설치돼 있어야 함 (CUDA 툴킷은 불필요, DLL 동봉됨) |

설치 → models에 GGUF 복사 → 바탕화면 아이콘 실행 → 브라우저가 자동으로 열림.

## MCP 서버 등록

설정(⚙) → MCP 항목에서 편집하거나 `mcp_servers.json`을 직접 수정:

```json
{
  "mcpServers": {
    "사내검색": { "url": "http://10.x.x.x:8082/mcp" },
    "파일도구": { "command": "python", "args": ["C:/tools/file_server.py"] },
    "꺼둔서버": { "url": "http://...", "disabled": true }
  }
}
```

- `url` → streamable_http, `command` → stdio 방식으로 연결
- 연결 실패한 서버는 건너뛰고 나머지로 동작 (앱이 죽지 않음)
- 도구 이름은 `서버이름__도구이름`으로 모델에 노출됨

## 외부 LLM 연결 (API 키)

설정(⚙) → "외부 LLM 연결"에서 추가한다. 프리셋: OpenAI / Anthropic(Claude) /
Google Gemini / 직접 입력(사내 게이트웨이 등). OpenAI 호환 `chat/completions`
규격이면 무엇이든 등록 가능하다.

| 프로바이더 | base_url | 모델 예시 |
|---|---|---|
| OpenAI | `https://api.openai.com/v1` | `gpt-4o-mini` |
| Anthropic | `https://api.anthropic.com/v1` | `claude-sonnet-5` |
| Google Gemini | `https://generativelanguage.googleapis.com/v1beta/openai` | `gemini-2.5-flash` |

- 화면 상단 드롭다운으로 로컬 ↔ 외부를 전환한다. 선택값도 서버에 저장되어
  브라우저를 바꿔도 유지된다.
- API 키는 `config.json`(데이터 폴더)에만 저장되고, 브라우저나 상태 API로는
  내려가지 않는다 (`has_key` 여부만 표시).
- 외부 모델 선택 시에는 로컬 모델이 없어도(목 모드) 실제 API가 호출된다.
- MCP 도구는 외부 모델에도 동일하게 노출된다 — GPT/Claude는 도구 호출
  정확도가 높아 MCP 활용에는 오히려 유리하다.
- 폐쇄망에서는 외부 API 대신 사내 프록시/게이트웨이 주소를 등록하는 용도로 쓴다.

## 참고/제약

- llama-server는 `--jinja`로 실행되어 Gemma의 chat template 기반 함수 호출을 쓴다.
  모델에 따라 도구 호출 정확도가 다를 수 있다 (Gemma는 공식 tool-use 학습이 약한 편).
- llama-server는 127.0.0.1에만 바인딩된다. 외부 공개는 UI 서버(`--host 0.0.0.0`)가 담당하므로
  모델 API가 직접 노출되지 않는다.
- 여러 명이 동시에 쓰려면 llama-server의 병렬 슬롯이 1이라 순차 처리된다.
  동시성이 필요하면 `llama_proc.py`에 `-np` 옵션을 추가할 것 (컨텍스트가 슬롯 수로 나뉜다).
