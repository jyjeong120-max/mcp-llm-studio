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
            else:
                raise ValueError('"url" 또는 "command" 중 하나가 필요합니다.')
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
    if getattr(result, "isError", False):
        joined = f"[도구가 오류를 반환했습니다]\n{joined}"
    return joined


class MCPManager:
    """설정 파일의 모든 MCP 서버를 관리하고, OpenAI tools 규격으로 노출한다."""

    def __init__(self, config_path: Path):
        self.config_path = config_path
        self.workers: dict[str, _ServerWorker] = {}

    async def start(self) -> None:
        try:
            raw = json.loads(self.config_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as e:
            print(f"[주의] mcp_servers.json을 읽지 못했습니다: {e}")
            return
        servers = raw.get("mcpServers", {})
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
        return {
            name: {
                "connected": w.connected,
                "error": w.error,
                "tools": [t["name"] for t in w.tools],
            }
            for name, w in self.workers.items()
        }
