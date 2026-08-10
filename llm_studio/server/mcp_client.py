"""mcp_client.py

mcp_servers.json에 정의된 MCP 서버들에 연결해 도구를 발견하고 실행한다.

설정 형식 (Claude Desktop과 동일한 mcpServers 규격):
    {
      "mcpServers": {
        "search":  {"url": "http://10.0.0.5:8082/mcp"},
        "files":   {"command": "python", "args": ["file_server.py"], "env": {}}
      }
    }
    - "url"이 있으면 streamable_http, "command"가 있으면 stdio로 연결한다.

구현 노트: MCP 세션은 열었던 태스크에서 닫아야 한다는 제약(anyio cancel scope)이
있어서, 서버마다 전용 워커 태스크를 두고 요청을 큐로 전달하는 구조를 쓴다.
연결 실패는 해당 서버만 비활성으로 표시하고 나머지는 계속 동작한다(우아한 저하).
"""

from __future__ import annotations

import asyncio
import json
from pathlib import Path

CALL_TIMEOUT_SEC = 120
SEPARATOR = "__"  # 도구 이름 규칙: <서버이름>__<도구이름>


class _ServerWorker:
    """MCP 서버 하나와의 연결을 소유하는 워커 태스크."""

    def __init__(self, name: str, spec: dict):
        self.name = name
        self.spec = spec
        self.queue: asyncio.Queue = asyncio.Queue()
        self.tools: list[dict] = []       # MCP 도구 원본 (name, description, inputSchema)
        self.connected = False
        self.error: str | None = None
        self.task: asyncio.Task | None = None
        self._ready = asyncio.Event()

    async def start(self) -> None:
        self.task = asyncio.create_task(self._run(), name=f"mcp-{self.name}")
        await self._ready.wait()

    async def _run(self) -> None:
        try:
            if "url" in self.spec:
                from mcp import ClientSession
                from mcp.client.streamable_http import streamablehttp_client

                async with streamablehttp_client(self.spec["url"]) as (read, write, _):
                    async with ClientSession(read, write) as session:
                        await self._serve(session)
            elif "command" in self.spec:
                from mcp import ClientSession, StdioServerParameters
                from mcp.client.stdio import stdio_client

                params = StdioServerParameters(
                    command=self.spec["command"],
                    args=self.spec.get("args", []),
                    env=self.spec.get("env") or None,
                )
                async with stdio_client(params) as (read, write):
                    async with ClientSession(read, write) as session:
                        await self._serve(session)
            elif "builtin" in self.spec:
                # 내장 서버 — 별도 프로세스 없이 FastMCP 객체에 인메모리로 붙는다.
                # spec["builtin"]은 mcp_server/의 FastMCP 인스턴스(직렬화 대상 아님, in-code).
                from fastmcp import Client

                async with Client(self.spec["builtin"]) as client:
                    await self._serve_fastmcp(client)
            else:
                raise ValueError('"url"·"command"·"builtin" 중 하나가 필요합니다.')
        except asyncio.CancelledError:
            raise
        except Exception as e:  # noqa: BLE001 — 서버 하나의 실패는 전체를 막지 않는다
            self.error = f"{type(e).__name__}: {e}"
            self.connected = False
        finally:
            self._ready.set()

    async def _serve(self, session) -> None:
        await session.initialize()
        listing = await session.list_tools()
        self.tools = [
            {
                "name": t.name,
                "description": t.description or "",
                "inputSchema": t.inputSchema or {"type": "object", "properties": {}},
            }
            for t in listing.tools
        ]
        self.connected = True
        self._ready.set()
        while True:  # 요청 처리 루프 — 워커 태스크 취소로 종료된다
            tool_name, arguments, future = await self.queue.get()
            try:
                result = await session.call_tool(tool_name, arguments)
                future.set_result(_flatten_content(result))
            except Exception as e:  # noqa: BLE001
                if not future.done():
                    future.set_exception(e)

    async def _serve_fastmcp(self, client) -> None:
        """인메모리 FastMCP 클라이언트용 서브 루프. _serve와 같은 일을 하되 fastmcp.Client의
        API(list_tools가 리스트를 바로 반환, call_tool은 raise_on_error로 결과 객체 회수)에
        맞춘다. (fastmcp.Client는 __aenter__에서 initialize를 대신하므로 별도 호출이 없다.)"""
        listing = await client.list_tools()
        self.tools = [
            {
                "name": t.name,
                "description": t.description or "",
                "inputSchema": t.inputSchema or {"type": "object", "properties": {}},
            }
            for t in listing
        ]
        self.connected = True
        self._ready.set()
        while True:  # 요청 처리 루프 — 워커 태스크 취소로 종료된다
            tool_name, arguments, future = await self.queue.get()
            try:
                # raise_on_error=False: 도구가 오류를 내도 예외 대신 결과 객체(is_error 포함)를
                # 받아 _flatten_content가 stdio/http 경로와 똑같이 문자열로 평탄화하게 한다.
                result = await client.call_tool(tool_name, arguments, raise_on_error=False)
                future.set_result(_flatten_content(result))
            except Exception as e:  # noqa: BLE001
                if not future.done():
                    future.set_exception(e)

    async def call(self, tool_name: str, arguments: dict) -> str:
        if not self.connected:
            raise RuntimeError(f"MCP 서버 '{self.name}'가 연결되어 있지 않습니다: {self.error}")
        future: asyncio.Future = asyncio.get_running_loop().create_future()
        await self.queue.put((tool_name, arguments, future))
        return await asyncio.wait_for(future, timeout=CALL_TIMEOUT_SEC)

    async def stop(self) -> None:
        if self.task and not self.task.done():
            self.task.cancel()
            try:
                await self.task
            except (asyncio.CancelledError, Exception):  # noqa: BLE001
                pass


def _flatten_content(result) -> str:
    """CallToolResult의 content 목록을 하나의 문자열로 합친다."""
    parts = []
    for item in result.content or []:
        text = getattr(item, "text", None)
        if text is not None:
            parts.append(text)
        else:
            parts.append(f"[{getattr(item, 'type', '비텍스트')} 콘텐츠]")
    joined = "\n".join(parts) or "(빈 결과)"
    # mcp SDK는 isError, fastmcp.Client는 is_error 로 오류를 표시한다 — 둘 다 본다.
    if getattr(result, "isError", False) or getattr(result, "is_error", False):
        joined = f"[도구가 오류를 반환했습니다]\n{joined}"
    return joined


class MCPManager:
    """설정 파일의 모든 MCP 서버를 관리하고, OpenAI tools 규격으로 노출한다."""

    def __init__(self, config_path: Path, config: dict | None = None):
        self.config_path = config_path
        # 앱 config 딕셔너리 참조 — 내장 서버 on/off(builtin_disabled)를 start/reload 때
        # 실시간으로 읽는다. 없으면 빈 dict(=전부 켬).
        self.config = config if config is not None else {}
        self.workers: dict[str, _ServerWorker] = {}
        self.builtin_status: list[dict] = []   # 내장 서버 로드 결과(미로드 사유 포함) — status()에 노출

    def _read_external_servers(self) -> dict:
        """mcp_servers.json의 mcpServers 매핑을 읽는다.

        파일이 없거나(신규 설치) 깨졌으면 빈 dict — 내장 도구는 계속 뜬다(우아한 저하).
        start()와 set_builtin()이 공유한다(내장/외부 이름 충돌 판정에 같은 소스를 봐야 한다).
        """
        try:
            raw = json.loads(self.config_path.read_text(encoding="utf-8"))
            return raw.get("mcpServers", {})
        except FileNotFoundError:
            return {}  # 설정 파일 없음 — 내장만으로 동작한다(정상)
        except (OSError, json.JSONDecodeError) as e:
            print(f"[주의] mcp_servers.json을 읽지 못했습니다(내장 도구는 계속): {e}")
            return {}

    async def start(self) -> None:
        # 1) 내장 서버(인메모리) — mcp_servers.json이 없어도 설정 없이 뜬다.
        from .builtin_servers import load_builtin_servers

        # 설정 파일 먼저 읽어 둔다(내장과 이름이 겹치면 외부 정의를 우선하려고).
        servers = self._read_external_servers()

        disabled = set(self.config.get("builtin_disabled", []))
        instances, self.builtin_status = load_builtin_servers(disabled=disabled)
        for name, instance in instances.items():
            if name in servers:
                continue  # 사용자가 같은 이름으로 외부 서버를 정의 → 그쪽을 쓴다(내장은 양보)
            worker = _ServerWorker(name, {"builtin": instance})
            self.workers[name] = worker
            await worker.start()
            if worker.connected:
                print(f"[정보] 내장 MCP '{name}' 로드됨 (도구 {len(worker.tools)}개)")
            else:
                print(f"[주의] 내장 MCP '{name}' 로드 실패: {worker.error}")

        # 2) 설정 파일의 외부 서버(stdio/http).
        for name, spec in servers.items():
            if spec.get("disabled"):
                continue
            worker = _ServerWorker(name, spec)
            self.workers[name] = worker
            await worker.start()
            if worker.connected:
                print(f"[정보] MCP '{name}' 연결됨 (도구 {len(worker.tools)}개)")
            else:
                print(f"[주의] MCP '{name}' 연결 실패: {worker.error}")

    async def stop(self) -> None:
        for worker in self.workers.values():
            await worker.stop()
        self.workers = {}

    async def reload(self) -> None:
        await self.stop()
        await self.start()

    async def set_builtin(self, name: str, enabled: bool) -> None:
        """내장 서버 하나만 켜고 끈다 — 다른 서버(내장·외부) 연결은 건드리지 않는다.

        전체 reload()는 외부 http/stdio 세션까지 전부 끊었다 다시 붙이므로(진행 중 호출
        끊김·재핸드셰이크 지연), 내장 토글 같은 국소 변경에는 해당 워커만 add/remove 한다.
        config.builtin_disabled는 호출 측(main.py)이 이미 갱신했다고 가정하고 그 값을 읽어 반영만.
        """
        from .builtin_servers import load_builtin_servers

        disabled = set(self.config.get("builtin_disabled", []))
        # import는 캐시되어 저렴 — 최신 로드 상태(미로드 사유 포함)를 갱신해 status()에 반영.
        instances, self.builtin_status = load_builtin_servers(disabled=disabled)
        external = self._read_external_servers()
        # 외부에서 같은 이름을 정의했으면 내장은 양보 — 그 외부 워커는 절대 건드리지 않는다.
        want = enabled and name in instances and name not in external
        existing = self.workers.get(name)
        have_builtin = existing is not None and "builtin" in existing.spec
        if want and not have_builtin:
            worker = _ServerWorker(name, {"builtin": instances[name]})
            self.workers[name] = worker
            await worker.start()
            if worker.connected:
                print(f"[정보] 내장 MCP '{name}' 로드됨 (도구 {len(worker.tools)}개)")
            else:
                print(f"[주의] 내장 MCP '{name}' 로드 실패: {worker.error}")
        elif not want and existing is not None and "builtin" in existing.spec:
            await existing.stop()
            self.workers.pop(name, None)

    def connected_servers(self) -> list[str]:
        """지금 연결된 MCP 서버 이름 목록 (계획-실행의 서버-스코프 라우팅용)."""
        return [name for name, w in self.workers.items() if w.connected]

    def openai_tools(self, servers=None) -> list[dict]:
        """연결된 서버의 도구를 OpenAI function calling 규격으로 변환한다.

        servers를 주면(서버 이름들의 목록/집합) 그 서버들의 도구만 내보낸다
        (서버-스코프 — 작업 모드가 스텝마다 필요한 서버로 좁힐 때 쓴다).
        None이면 전부. 빈 목록이면 도구 없음.
        """
        allow = set(servers) if servers is not None else None
        specs = []
        for server_name, worker in self.workers.items():
            if not worker.connected:
                continue
            if allow is not None and server_name not in allow:
                continue
            for tool in worker.tools:
                specs.append({
                    "type": "function",
                    "function": {
                        "name": f"{server_name}{SEPARATOR}{tool['name']}",
                        "description": tool["description"][:1024],
                        "parameters": tool["inputSchema"],
                    },
                })
        return specs

    async def call(self, qualified_name: str, arguments: dict) -> str:
        server_name, _, tool_name = qualified_name.partition(SEPARATOR)
        worker = self.workers.get(server_name)
        if worker is None:
            raise RuntimeError(f"알 수 없는 MCP 서버: {server_name}")
        return await worker.call(tool_name, arguments)

    def status(self) -> dict:
        disabled = set(self.config.get("builtin_disabled", []))
        # 내장 서버 설명(UI 표시용) 조회를 위해 로드 상태를 이름으로 색인.
        desc_by_name = {st["name"]: st.get("desc", "") for st in self.builtin_status}
        out = {}
        for name, w in self.workers.items():
            is_builtin = "builtin" in w.spec
            entry = {
                "connected": w.connected,
                "error": w.error,
                "tools": [t["name"] for t in w.tools],
                "builtin": is_builtin,
            }
            if is_builtin:
                entry["enabled"] = name not in disabled  # 워커로 떴으면 켜진 것
                entry["desc"] = desc_by_name.get(name, "")
            out[name] = entry
        # 워커로 뜨지 못한 내장 서버(import 실패/설정에서 끔)도 상태로 노출한다.
        for st in self.builtin_status:
            if not st["loaded"] and st["name"] not in out:
                out[st["name"]] = {
                    "connected": False,
                    "error": st["reason"],
                    "tools": [],
                    "builtin": True,
                    "enabled": st["name"] not in disabled,
                    "desc": st.get("desc", ""),
                }
        return out
