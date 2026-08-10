"""config.py

데이터 폴더 해석과 설정 파일 관리.

데이터 폴더는 인스톨러가 권한을 부여한 C:\\ProgramData\\LocalLLMStudio를 우선 사용하고,
쓰기가 불가능하면(개발 중, 미설치 환경) %LOCALAPPDATA%\\LocalLLMStudio로 내려간다.
어느 쪽이든 하위 구조는 동일하다:

    models/         GGUF 모델 파일을 넣는 곳
    conversations/  대화 기록 (JSON)
    uploads/        첨부 파일 원본 + 추출 텍스트
    logs/           llama-server 로그
    config.json     생성/서버 설정
    mcp_servers.json MCP 서버 연결 설정
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

APP_NAME = "LocalLLMStudio"
SUBDIRS = ("models", "conversations", "uploads", "logs")

DEFAULT_CONFIG = {
    # 생성 파라미터 (요청마다 반영, 재시작 불필요)
    "system_prompt": "당신은 사내망에서 동작하는 로컬 LLM 어시스턴트입니다. 한국어로 정확하고 간결하게 답하세요.",
    "temperature": 1.0,
    "top_p": 0.95,
    "top_k": 64,
    "max_tokens": 4096,
    # 한 턴에서 모델이 도구를 부를 수 있는 최대 라운드 수 (무한 도구 루프 방지).
    # 초과하면 오류 이벤트로 중단한다. 도구를 많이 쓰는 작업이면 높인다.
    "max_tool_rounds": 8,
    # 서버 파라미터 (변경 시 llama-server 재시작 필요)
    "ctx": 32768,
    "kv_quant": False,
    "model_path": "",  # 비우면 models/ 폴더에서 가장 최근 .gguf를 자동 선택
    "model_alias": "gemma-12b-it-qat",
    "llama_port": 8000,
    # GPU 오프로드 레이어 수. 99=전 레이어(VRAM 넉넉한 서버용 기본값). VRAM이 작으면
    # 낮춰 일부만 올리고 나머지는 CPU로(예: 2GB dGPU면 16 안팎). 0이면 CPU 전용.
    "gpu_layers": 99,
    # 비우면 이 앱이 llama-server를 직접 띄운다(managed). 주소가 있으면 그 서버에
    # 붙기만 한다(external) — serve_llm.py로 LLM 서버를 따로 띄웠을 때 쓴다.
    # 예: "http://127.0.0.1:8000". app.py --llama-url 로도 설정된다.
    "llama_external_url": "",
    # 앱은 기본적으로 아무 로컬 모델도 서빙하지 않은 '유휴' 상태로 뜬다 (사용자가 UI에서
    # 모델을 골라 시작한다). 이 값을 True로 켜면 앱 시작 시 위 서버 설정으로 자동 서빙한다.
    # 자동 시작이 실패해도 앱은 유휴 상태로 계속 뜬다 (목 모드로 떨어지지 않음).
    "autostart_local": False,
    # 외부 LLM 연결 (OpenAI 호환 API — API 키로 접속)
    # 각 항목: {"name": "표시이름", "base_url": "https://.../v1", "api_key": "...", "model": "모델ID"}
    "providers": [],
    # 현재 선택된 모델: "local"(서빙 중인 로컬 LLM) 또는 providers의 name
    "active_provider": "local",
    # 장기 메모리 (대화 넘나드는 사실 기억 — memory.db). 전부 생성측이라 재시작 불필요.
    "memory_enabled": True,             # 끄면 회상 주입·자동요약을 건너뛴다 (채팅은 정상)
    "memory_recall_top_k": 5,           # 매 턴 주입할 관련 기억 최대 개수
    "memory_autosummary_enabled": True, # 대화에서 사실을 자동 추출해 저장할지
    "memory_autosummary_turn_interval": 25,   # 이 턴 수마다 자동요약 1회
    "memory_autosummary_char_threshold": 24000,  # 누적 이력이 이 문자수를 넘으면 자동요약
    "compact_keep_recent_turns": 4,     # 대화 압축 시 원문으로 남길 최근 user 턴 수 (그 이전은 요약)
    # 계획-실행(작업 모드) + 실행 코크핏. 요청 mode가 "task"거나, "auto"에서 라우터가
    # 작업으로 분류하면 다단계로 처리한다. "chat"이면 항상 일반 채팅.
    "task_mode_enabled": True,   # 끄면 작업 요청도 일반 채팅으로 처리
    "task_max_steps": 10,        # 총 스텝 실행 상한 (무한 루프 방지)
    "task_max_replans": 2,       # 재계획 예산 (실패 시 남은 계획 재수립 횟수)
    # 의도 라우터: mode="auto"일 때 요청을 작업/채팅으로 자동 분류한다. 끄면 auto는
    # 채팅으로 처리한다(사용자가 🧭를 눌러 강제 작업으로 돌릴 수 있음).
    "task_router_enabled": True,
    # 실패 게이트(반자동): 스텝 실패 시 재시도/건너뛰기/재계획/편집/중단을 묻는다.
    # 무응답이면 그대로 진행. (계획 게이트는 제거됨 — 계획을 세우면 항상 바로 실행한다.
    # 위험 도구는 실행 직전 approval 게이트가 따로 잡는다.)
    "task_failure_gate": True,
    "task_steer_timeout": 600,   # 조종 대기 상한(초). 0 이하 = 무제한. 초과 시 그대로 진행.
    # 조건 분기('[?→M]'): 계획가가 분기 태그를 붙이면 조건 판정으로 이후 단계를 건너뛴다.
    # 끄면 분기 태그를 본문의 일부로 두어 일반 스텝처럼 실행한다.
    "task_branch_enabled": True,
    # 위험 도구 승인 게이트: 모델이 confirm=true 인자(파괴적 동작의 실제 실행)로 도구를
    # 부르거나 approval_tools에 오른 도구를 부르면, 실행 전에 브라우저에 승인/거절
    # 버튼을 띄워 사용자의 결정을 기다린다. MCP 서버 쪽 confirm 게이트와 이중 안전장치 —
    # 모델이 사용자에게 묻지 않고 스스로 confirm=true를 넣는 사고를 막는다.
    "approval_enabled": True,
    "approval_timeout": 600,   # 승인 대기 제한(초). 0 이하 = 무제한 대기. 초과 시 실행하지 않음(거절과 동일)
    "approval_tools": [],      # 항상 승인이 필요한 도구 이름 목록 (send_email 또는 outlook__send_email)
    # 첨부 처리. True면 Office 문서(docx/xlsx/pptx)를 서버가 바이트로 추출하지 않고
    # 저장 경로를 모델에 줘서 office MCP 도구(COM)가 읽게 한다. 사내 DRM처럼 파일이
    # 암호화돼 바이트 파싱은 암호문만 나오고 Word/Excel(COM)로 열어야만 복호화되는
    # 환경용 스위치. office_server가 연결돼 있어야 실제로 읽힌다. (업로드 시점 적용, 재시작 불필요)
    "attachment_com_office": False,
    # 내장(builtin) MCP 도구 서버 중 끌 것들의 이름 목록. 앱에 동봉된 office/outlook/pdf/text
    # 서버를 별도 프로세스 없이 인메모리로 장착하는데(builtin_servers.py), 여기 이름을 넣으면
    # 그 서버만 로드하지 않는다. (설정 → MCP의 내장 도구 토글이 조작)
    # outlook은 외부로 메일을 실제 발송하는 파괴적 도구(send_email 등)를 포함하므로 기본으로
    # 꺼 둔다 — opt-in. 필요하면 토글로 켠다(켜도 confirm·승인 게이트는 그대로 유효). 도구 수를
    # 줄여 약한 로컬 모델의 도구 선택 정확도도 함께 지킨다. 나머지 셋(pdf/text/office 읽기 중심)은 기본 켬.
    "builtin_disabled": ["outlook"],
}

# 이 키들이 바뀌면 llama-server를 재시작해야 반영된다.
RESTART_KEYS = {"ctx", "kv_quant", "model_path", "model_alias", "llama_port",
                "llama_external_url", "gpu_layers"}

DEFAULT_MCP_CONFIG = {"mcpServers": {}}


def app_dir() -> Path:
    """실행 파일(또는 소스 루트)이 있는 폴더. 동봉된 llama-server를 찾는 기준."""
    if getattr(sys, "frozen", False):  # PyInstaller로 묶인 경우
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


def static_dir() -> Path:
    """웹 UI 정적 파일 폴더. PyInstaller 번들이면 _MEIPASS(_internal) 아래에 있다."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", str(app_dir()))) / "static"
    return app_dir() / "static"


def _writable(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return True
    except OSError:
        return False


def resolve_data_dir(override: str | None = None) -> Path:
    """쓰기 가능한 데이터 폴더를 정한다. 인자 > ProgramData > LocalAppData 순."""
    candidates = []
    if override:
        candidates.append(Path(override))
    program_data = os.environ.get("PROGRAMDATA")
    if program_data:
        candidates.append(Path(program_data) / APP_NAME)
    local_appdata = os.environ.get("LOCALAPPDATA")
    if local_appdata:
        candidates.append(Path(local_appdata) / APP_NAME)
    candidates.append(Path.home() / f".{APP_NAME}")

    for cand in candidates:
        if _writable(cand):
            for sub in SUBDIRS:
                (cand / sub).mkdir(exist_ok=True)
            return cand
    raise RuntimeError("쓰기 가능한 데이터 폴더를 찾지 못했습니다.")


def load_config(data_dir: Path) -> dict:
    """config.json을 읽어 기본값 위에 덮어쓴다. 파일이 없으면 기본값으로 만든다."""
    path = data_dir / "config.json"
    config = dict(DEFAULT_CONFIG)
    if path.exists():
        try:
            config.update(json.loads(path.read_text(encoding="utf-8")))
        except (json.JSONDecodeError, OSError) as e:
            print(f"[주의] config.json을 읽지 못해 기본값을 씁니다: {e}")
    else:
        save_config(data_dir, config)
    return config


def save_config(data_dir: Path, config: dict) -> None:
    path = data_dir / "config.json"
    path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def mcp_config_path(data_dir: Path) -> Path:
    path = data_dir / "mcp_servers.json"
    if not path.exists():
        path.write_text(
            json.dumps(DEFAULT_MCP_CONFIG, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    return path
