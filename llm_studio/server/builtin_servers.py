"""builtin_servers.py

`mcp_server/`의 FastMCP 서버들을 **별도 프로세스 없이** 앱 안에서 인메모리로 쓰기 위한
레지스트리. llm_studio는 이미 사용자 로그인 세션에서 실행되므로(=COM 가능), office/outlook/
pdf/text 코드를 stdio·http로 띄워 소켓으로 오갈 필요가 없다. 앱이 그 FastMCP 객체에 직접
붙으면(fastmcp 인메모리 트랜스포트, mcp_client 참고) 설정·기동 없이 바로 도구로 쓴다.

- **로직 중복 0**: 같은 서버 파일이 (1) 여기서 내장 도구로, (2) `python pdf_server.py` 등
  독립 MCP 서버로 이중 사용된다. 도구 스키마·디스패치·confirm 게이트가 전부 재사용된다.
- **우아한 저하**: 어떤 모듈이 import에 실패하면(의존성 없음 등) 그 서버만 건너뛰고 사유를
  status에 남긴다. 나머지 내장 서버는 계속 뜬다.
- rag는 내장에서 제외한다 — 임베딩 llama-server와 Qdrant가 있어야 하고 인덱싱이 전제라
  '설정 없이 바로'라는 내장 취지와 맞지 않는다(외부 MCP로 등록해 쓴다).

⚠ exe 배포(build_exe.bat/PyInstaller)에선 `mcp_server/`의 이 모듈들과 pywin32를
hidden-import로 포함해야 한다 — 소스 실행은 아래 sys.path 삽입으로 바로 동작한다.
"""

from __future__ import annotations

import sys
from pathlib import Path

# 내장으로 제공할 서버들: (내장 이름, mcp_server 모듈명, 사람이 읽을 설명).
# 내장 이름이 곧 도구 접두사가 된다 (예: pdf__read_pdf_text). 설정 파일에서 같은 이름으로
# 외부 서버를 정의하면 그쪽이 우선한다(MCPManager.start 참고).
BUILTIN_MODULES: list[tuple[str, str, str]] = [
    ("pdf", "pdf_server", "DRM PDF 텍스트 읽기 (Word COM 경유)"),
    ("text", "text_server", "DRM 텍스트/코드 파일 읽기 (Word COM 경유)"),
    ("office", "office_server", "실행 중인 Office(Word/Excel/PPT) 읽기·수정"),
    ("outlook", "outlook_server", "Outlook 메일·일정·연락처·작업"),
]


def _mcp_server_dir() -> Path:
    """이 파일(llm_studio/server/builtin_servers.py) 기준으로 repo의 mcp_server/ 폴더."""
    return Path(__file__).resolve().parents[2] / "mcp_server"


def load_builtin_servers(disabled: set[str] | None = None) -> tuple[dict, list[dict]]:
    """내장 서버들을 import해 (이름→FastMCP 인스턴스, 상태 목록)을 돌려준다.

    disabled에 든 이름은 로드하지 않는다(설정 토글). import에 실패한 서버는 건너뛰되
    사유를 상태에 남긴다 — 프로세스가 죽지 않도록 예외를 문자열로 흡수한다.
    """
    disabled = disabled or set()
    d = _mcp_server_dir()
    if d.is_dir() and str(d) not in sys.path:
        # 내장 서버 모듈과 그들 간 상호 import(rag가 office를 부르는 식)를 위해 폴더를 얹는다.
        # append로 '끝에' 붙인다 — insert(0)로 앞에 두면 mcp_server/에 site-packages·stdlib과
        # 이름이 겹치는 모듈이 생길 때 그걸 프로세스 전역에서 가려 버린다. 우리 모듈명은
        # 유일하므로 뒤에 둬도 그대로 찾힌다(우선순위만 안전하게 낮춘다).
        sys.path.append(str(d))

    instances: dict[str, object] = {}
    statuses: list[dict] = []
    for name, module_name, desc in BUILTIN_MODULES:
        if name in disabled:
            statuses.append({"name": name, "desc": desc, "loaded": False, "reason": "설정에서 끔"})
            continue
        try:
            module = __import__(module_name)
            instance = getattr(module, "mcp", None)
            if instance is None:
                raise AttributeError(f"{module_name}에 FastMCP 인스턴스 'mcp'가 없습니다")
            instances[name] = instance
            statuses.append({"name": name, "desc": desc, "loaded": True, "reason": ""})
        except Exception as e:  # noqa: BLE001 — 의존성 부재/모듈 없음 등은 그 서버만 저하
            statuses.append({
                "name": name, "desc": desc, "loaded": False,
                "reason": f"{type(e).__name__}: {e}",
            })
    return instances, statuses
