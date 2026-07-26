"""main.py

FastAPI 앱: 웹 UI 정적 파일 + REST/SSE API.

엔드포인트:
    GET  /api/status                 서버/모델/MCP 상태
    POST /api/chat                   채팅 (SSE 스트리밍, 도구 호출 이벤트 포함)
    POST /api/chat/approve           위험 도구 승인/거절 응답 (approval_request에 대한 답)
    GET  /api/conversations          대화 목록
    GET  /api/conversations/{id}     대화 내용
    DELETE /api/conversations/{id}   대화 삭제
    POST /api/conversations/{id}/rename  제목 변경
    POST /api/upload                 파일 첨부 (텍스트 추출)
    GET  /api/settings, PUT /api/settings
    GET  /api/mcp                    MCP 상태/도구 목록
    GET  /api/mcp/config, PUT /api/mcp/config (저장 후 재연결)
    POST /api/server/restart         llama-server 재시작 (설정 반영)
"""

from __future__ import annotations

import asyncio
import json
import os
from contextlib import asynccontextmanager

from fastapi import FastAPI, HTTPException, Request, UploadFile
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import agent, planner
from .approvals import ApprovalBroker
from .config import RESTART_KEYS, save_config, static_dir

ATTACH_MAX_CHARS = 30_000  # 첨부 파일당 프롬프트에 넣는 텍스트 한도


class ChatRequest(BaseModel):
    message: str
    conversation_id: str | None = None
    attachments: list[str] = []
    provider: str | None = None  # None이면 config의 active_provider 사용
    task_mode: bool = False      # True면 계획-실행(다단계)으로 처리 ('작업 모드')


def _resolve_provider(state, name: str | None) -> dict:
    """요청된 프로바이더를 접속 정보 dict로 푼다.

    "local" → 서빙 중인 llama-server. 그 외에는 config["providers"]에서
    이름으로 찾는다 (외부 OpenAI 호환 API — 키는 config.json에만 저장됨).
    """
    name = name or state.config.get("active_provider") or "local"
    if name == "local":
        return {
            "name": "local",
            "base_url": f"{state.llama.base_url}/v1",
            "api_key": "local",
            "model": state.llama.alias,
            "send_top_k": True,
            "mock": state.mock,  # 목 모드는 로컬 모델이 없을 때만 해당
        }
    for p in state.config.get("providers", []):
        if p.get("name") == name:
            if not p.get("base_url"):
                raise HTTPException(400, f"프로바이더 '{name}'에 base_url이 없습니다.")
            return {
                "name": name,
                "base_url": p["base_url"].rstrip("/"),
                "api_key": p.get("api_key", ""),
                "model": p.get("model", ""),
                "send_top_k": False,
                "mock": False,  # 외부 API는 로컬 모델이 없어도 실제 호출
            }
    raise HTTPException(404, f"프로바이더 '{name}'를 찾을 수 없습니다. 설정에서 등록하세요.")


class RenameRequest(BaseModel):
    title: str


class ApproveRequest(BaseModel):
    id: str          # approval_request 이벤트로 받은 요청 id
    approved: bool   # True=승인(실행), False=거절(실행 안 함)


class MCPConfigRequest(BaseModel):
    content: str


def create_app(state) -> FastAPI:
    """state: app.py가 조립한 AppState (config, llama, mcp, store, uploads...)."""

    @asynccontextmanager
    async def lifespan(_app):
        await state.mcp.start()
        yield
        await state.mcp.stop()
        state.llama.stop()

    app = FastAPI(title="LocalLLM Studio", lifespan=lifespan)

    # 위험 도구 승인 브로커 — 상태는 전부 RAM (대화 저장과 무관, 재시작하면 초기화).
    approvals = ApprovalBroker()

    # ---------- 상태 ----------

    @app.get("/api/status")
    async def get_status():
        return {
            "mock": state.mock,
            "mock_reason": state.mock_reason,
            "llama": state.llama.status(),
            # 유휴(미서빙) 상태에서도 UI가 '무엇을 시작할지' 판단할 수 있도록 저장된
            # 서버 설정을 함께 내려준다 (external URL이 있으면 붙기 버튼, 없으면 모델 선택).
            "llama_config": {
                "external_url": state.config.get("llama_external_url", ""),
                "model_path": state.config.get("model_path", ""),
                "autostart": state.config.get("autostart_local", False),
            },
            "mcp": state.mcp.status(),
            "data_dir": str(state.data_dir),
            "active_provider": state.config.get("active_provider", "local"),
            # 키는 내려보내지 않는다 — 등록 여부만 표시
            "providers": [
                {"name": p.get("name", ""), "model": p.get("model", ""),
                 "has_key": bool(p.get("api_key"))}
                for p in state.config.get("providers", [])
            ],
        }

    # ---------- 채팅 (SSE) ----------

    @app.post("/api/chat")
    async def chat(req: ChatRequest):
        provider = _resolve_provider(state, req.provider)

        # 로컬을 골랐는데 아직 요청을 받을 수 없으면(유휴이거나 모델 로딩 중) 실제 호출
        # 대신 안내만 흘려보낸다. running이 아니라 ready로 판단한다 — 프로세스는 떴지만
        # 헬스체크 전(로딩 중)인 서버로 요청을 보내면 원시 연결 오류가 나기 때문이다.
        # (목 모드는 --mock 전용 canned 응답이라 이 경로를 타지 않는다.)
        if provider["name"] == "local" and not provider["mock"] and not state.llama.ready:
            loading = state.llama.running  # 프로세스는 떴지만 아직 준비 안 됨 = 로딩 중
            msg = ("로컬 모델을 로딩하는 중입니다(모델 크기에 따라 몇 분 걸릴 수 있습니다). "
                   "잠시 후 다시 시도하세요."
                   if loading else
                   "로컬 LLM이 아직 서빙되지 않았습니다. 상단의 [서빙 시작] 또는 설정에서 "
                   "모델을 골라 시작하거나, 외부 API를 선택하세요.")

            async def not_serving():
                yield _sse({"type": "error", "message": msg})
            return StreamingResponse(
                not_serving(), media_type="text/event-stream",
                headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
            )

        conv = state.store.load(req.conversation_id) if req.conversation_id else None
        if conv is None:
            conv = state.store.create(state.store.title_from(req.message))

        user_content = _with_attachments(state, req.message, req.attachments)
        conv["messages"].append({"role": "user", "content": user_content})

        # 회상(읽기): 원본 메시지로 관련 기억을 찾아 system에 주입한다. 검색 질의는
        # 첨부 텍스트가 아니라 사용자가 친 원문(req.message)을 쓴다.
        memory_enabled = state.config.get("memory_enabled", True)
        system_content = _build_system(state, req.message)
        msgs = [{"role": "system", "content": system_content}, *conv["messages"]]
        mem = state.memory if memory_enabled else None

        # 라우터: 작업 모드 요청이면 계획-실행으로 분기한다. 단순 대화는 기존 경로.
        # 목 모드에선 계획을 세울 수 없으므로 항상 일반 채팅으로 처리한다.
        use_task = (req.task_mode and not provider["mock"]
                    and state.config.get("task_mode_enabled", True))
        if use_task:
            runner = planner.run_task(
                base_url=provider["base_url"], model=provider["model"],
                api_key=provider["api_key"], send_top_k=provider["send_top_k"],
                messages=msgs, settings=state.config, mcp=state.mcp, memory=mem,
                max_steps=int(state.config.get("task_max_steps", 10)),
                max_replans=int(state.config.get("task_max_replans", 2)),
                approver=approvals,
            )
        else:
            runner = agent.run_chat(
                base_url=provider["base_url"], model=provider["model"],
                api_key=provider["api_key"], send_top_k=provider["send_top_k"],
                messages=msgs, settings=state.config, mcp=state.mcp, memory=mem,
                mock=provider["mock"], approver=approvals,
            )

        async def event_stream():
            yield _sse({"type": "meta", "conversation_id": conv["id"], "title": conv["title"]})
            streamed: list[str] = []
            finished = False
            try:
                async for event in runner:
                    if event["type"] == "token":
                        streamed.append(event["text"])
                    if event["type"] == "done":
                        # system 프롬프트를 뺀 나머지를 대화 이력으로 저장
                        conv["messages"] = event["messages"][1:]
                        finished = True
                        yield _sse({"type": "done"})
                    else:
                        yield _sse(event)
            except asyncio.CancelledError:
                raise  # 클라이언트 중단 — finally에서 부분 응답을 저장
            finally:
                if not finished and streamed:
                    conv["messages"].append(
                        {"role": "assistant", "content": "".join(streamed) + "\n\n*(중단됨)*"}
                    )
                state.store.save(conv)
                # 쓰기(자동요약): 정상 완료된 턴에서만, 트리거가 맞으면 사실을 추출해 저장한다.
                # 사용자는 이미 응답을 다 받았으므로 여기서의 지연은 스트림 종료만 늦춘다.
                if finished:
                    await _maybe_autosummarize(state, conv, provider)

        return StreamingResponse(
            event_stream(),
            media_type="text/event-stream",
            headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
        )

    @app.post("/api/chat/approve")
    async def approve_tool(req: ApproveRequest, request: Request):
        """위험 도구 승인 요청(approval_request)에 대한 사용자의 결정을 반영한다."""
        # 0.0.0.0 공개 시 원격 호스트가 승인 게이트를 대신 풀 수 없게 로컬 전용
        # (id도 추측 불가 토큰이지만 이중 방어 — 파일 탐색 API와 같은 원칙).
        _require_local(request)
        if not approvals.resolve(req.id, req.approved):
            raise HTTPException(404, "해당 승인 요청이 없습니다 (이미 처리됐거나 시간 초과).")
        return {"ok": True}

    # ---------- 대화 기록 ----------

    @app.get("/api/conversations")
    async def list_conversations():
        return state.store.list()

    @app.get("/api/conversations/{conv_id}")
    async def get_conversation(conv_id: str):
        conv = state.store.load(conv_id)
        if conv is None:
            raise HTTPException(404, "대화를 찾을 수 없습니다.")
        return conv

    @app.delete("/api/conversations/{conv_id}")
    async def delete_conversation(conv_id: str):
        if not state.store.delete(conv_id):
            raise HTTPException(404, "대화를 찾을 수 없습니다.")
        return {"ok": True}

    @app.post("/api/conversations/{conv_id}/rename")
    async def rename_conversation(conv_id: str, req: RenameRequest):
        if not state.store.rename(conv_id, req.title):
            raise HTTPException(404, "대화를 찾을 수 없습니다.")
        return {"ok": True}

    # ---------- 장기 메모리 (조회·삭제) ----------
    # 사용자가 어시스턴트가 무엇을 기억하는지 보고 지울 수 있어야 한다 (신뢰·프라이버시).
    # 폴리시된 관리 UI는 fast-follow. 여기서는 최소 API만 노출한다.

    @app.get("/api/memory")
    async def list_memory():
        return {"count": state.memory.count(), "items": state.memory.all()}

    @app.delete("/api/memory/{mem_id}")
    async def delete_memory(mem_id: int):
        if not state.memory.delete(mem_id):
            raise HTTPException(404, "해당 기억을 찾을 수 없습니다.")
        return {"ok": True}

    # ---------- 파일 첨부 ----------

    @app.post("/api/upload")
    async def upload(file: UploadFile):
        data = await file.read()
        try:
            return state.uploads.save(
                file.filename or "attachment", data,
                com_office=state.config.get("attachment_com_office", False),
            )
        except ValueError as e:
            raise HTTPException(400, str(e))

    # ---------- 네이티브 파일 선택 대화상자 (원본 제자리 읽기용) ----------
    # 사내 DRM 문서는 사본을 만들면(브라우저 업로드) 복호화가 깨진다. 그래서 앱이
    # 로컬에서 도는 점을 이용해 서버가 OS 기본 '열기' 대화상자를 띄워 사용자가 고른
    # 원본 경로를 그대로 받고, 복사 없이 경로만 등록한다(office COM이 제자리에서 읽음).
    # 대화상자는 서버 프로세스가 있는 그 데스크톱 세션에 뜬다(=사용자 세션, COM과 동일).
    # 로컬(127.0.0.1) 접속에서만 허용한다 (0.0.0.0로 열어도 다른 PC가 못 띄우게).

    @app.post("/api/fs/dialog")
    async def fs_dialog(request: Request):
        _require_local(request)
        loop = asyncio.get_running_loop()
        try:
            path = await loop.run_in_executor(None, _native_open_dialog)
        except Exception as e:  # noqa: BLE001 — tkinter 부재/데스크톱 없음 등
            raise HTTPException(
                500,
                f"파일 대화상자를 열지 못했습니다({e}). 이 PC의 로그인 세션에서 "
                f"소스로 실행 중인지 확인하세요.",
            )
        if not path:
            return {"cancelled": True}  # 사용자가 취소
        try:
            return state.uploads.register_path(
                path, com_office=state.config.get("attachment_com_office", False),
            )
        except ValueError as e:
            raise HTTPException(400, str(e))

    # ---------- 설정 ----------

    @app.get("/api/settings")
    async def get_settings():
        return state.config

    @app.put("/api/settings")
    async def put_settings(new: dict):
        unknown = set(new) - set(state.config)
        if unknown:
            raise HTTPException(400, f"알 수 없는 설정 키: {sorted(unknown)}")
        if "providers" in new:
            if not isinstance(new["providers"], list) or any(
                not isinstance(p, dict) or not str(p.get("name", "")).strip()
                for p in new["providers"]
            ):
                raise HTTPException(400, "providers는 name이 있는 객체의 목록이어야 합니다.")
            names = [p["name"].strip() for p in new["providers"]]
            if "local" in names or len(names) != len(set(names)):
                raise HTTPException(400, "프로바이더 이름은 중복될 수 없고 'local'은 예약어입니다.")
        if "active_provider" in new:
            providers = new.get("providers", state.config.get("providers", []))
            valid = {"local"} | {p.get("name") for p in providers}
            if new["active_provider"] not in valid:
                raise HTTPException(400, f"등록되지 않은 프로바이더: {new['active_provider']}")
        restart_required = any(
            key in RESTART_KEYS and state.config.get(key) != value
            for key, value in new.items()
        )
        state.config.update(new)
        save_config(state.data_dir, state.config)
        return {"ok": True, "restart_required": restart_required}

    # ---------- MCP ----------

    @app.get("/api/mcp")
    async def get_mcp():
        return state.mcp.status()

    @app.get("/api/mcp/config")
    async def get_mcp_config():
        return {"content": state.mcp.config_path.read_text(encoding="utf-8")}

    @app.put("/api/mcp/config")
    async def put_mcp_config(req: MCPConfigRequest):
        try:
            json.loads(req.content)  # 저장 전 JSON 문법 검증
        except json.JSONDecodeError as e:
            raise HTTPException(400, f"JSON 문법 오류: {e}")
        state.mcp.config_path.write_text(req.content, encoding="utf-8")
        await state.mcp.reload()
        return {"ok": True, "status": state.mcp.status()}

    # ---------- 로컬 모델 서빙 (선택·시작·중지·재시작) ----------
    # 서빙은 앱 시작과 분리돼 있다. 사용자가 UI에서 GGUF를 고르고 옵션을 정한 뒤
    # 시작/중지한다. 서버 파라미터(모델 경로·gpu_layers·ctx 등)는 프런트가 먼저
    # PUT /api/settings로 저장하고, 여기서는 저장된 config로 start/restart를 건다.

    @app.get("/api/models")
    async def list_models():
        """models/ 폴더의 .gguf 목록. UI 모델 선택 드롭다운이 쓴다."""
        return {"models": state.llama.list_models(), "current": state.config.get("model_path", "")}

    @app.post("/api/server/start")
    async def start_server():
        if state.mock:
            raise HTTPException(400, "목(--mock) 모드에서는 서버를 시작할 수 없습니다.")
        if state.llama.running:
            raise HTTPException(400, "이미 서빙 중입니다. 설정을 바꿨다면 [재시작]을 쓰세요.")
        loop = asyncio.get_running_loop()
        try:
            await loop.run_in_executor(None, state.llama.start, state.config)
        except Exception as e:  # noqa: BLE001 — 시작 실패는 500으로 알리되 앱은 계속 뜬다
            raise HTTPException(500, f"서빙 시작 실패: {e}")
        return {"ok": True, "llama": state.llama.status()}

    @app.post("/api/server/stop")
    async def stop_server():
        # 외부 서버는 우리 것이 아니므로 stop()이 연결만 끊는다(프로세스 유지).
        loop = asyncio.get_running_loop()
        await loop.run_in_executor(None, state.llama.stop)
        return {"ok": True, "llama": state.llama.status()}

    @app.post("/api/server/restart")
    async def restart_server():
        if state.mock:
            raise HTTPException(400, "목 모드에서는 재시작할 수 없습니다.")
        loop = asyncio.get_running_loop()
        try:
            await loop.run_in_executor(None, state.llama.restart, state.config)
        except Exception as e:  # noqa: BLE001
            raise HTTPException(500, f"재시작 실패: {e}")
        return {"ok": True, "llama": state.llama.status()}

    # ---------- 호스팅(앱 서버) 종료 ----------

    @app.post("/api/shutdown")
    async def shutdown():
        """이 앱(웹 UI + 자체 llama-server)을 통째로 내린다.

        uvicorn Server.should_exit를 세우면 우아한 종료가 돌면서 lifespan의
        shutdown이 실행돼 mcp.stop()·llama.stop()이 호출된다. external 모드라면
        따로 띄운 LLM 서버는 건드리지 않고 연결만 끊는다(그 서버는 계속 떠 있다).
        응답을 먼저 돌려보내야 페이지가 '종료됨'을 표시할 수 있으므로, 종료 신호는
        약간 늦춰 건다.
        """
        server = getattr(state, "server", None)

        async def _trigger():
            await asyncio.sleep(0.3)  # 이 요청의 응답이 먼저 나가도록 잠깐 양보
            if server is not None:
                server.should_exit = True
            else:
                # uvicorn 핸들이 없는 예외적 실행 경로(테스트 등) — 직접 종료.
                state.llama.stop()
                os._exit(0)

        asyncio.create_task(_trigger())
        return {"ok": True, "external": state.llama.status().get("external", False)}

    # 정적 파일은 마지막에 마운트한다 (/api/*가 우선 매칭되도록)
    app.mount("/", StaticFiles(directory=str(static_dir()), html=True), name="static")
    return app


def _build_system(state, query: str) -> str:
    """system 프롬프트에 관련 장기 기억을 [기억] 블록으로 덧붙인다.

    회상(읽기)은 모델의 도구 호출에 맡기지 않고 하네스가 매 턴 결정적으로 주입한다
    (약한 로컬 모델도 확실히 기억을 참고하도록). 메모리가 꺼져 있거나 검색이
    실패하거나 관련 기억이 없으면 원래 프롬프트를 그대로 쓴다 — 우아하게 저하한다.
    """
    base = state.config["system_prompt"]
    if not state.config.get("memory_enabled", True):
        return base
    try:
        top_k = int(state.config.get("memory_recall_top_k", 5))
        hits = state.memory.search(query, top_k)
    except Exception as e:  # noqa: BLE001 — 검색 실패가 채팅을 막지 않게 한다
        print(f"[주의] 메모리 검색 실패, 기억 없이 진행합니다: {e}")
        return base
    if not hits:
        return base
    block = "\n".join(f"- {h['content']}" for h in hits)
    return (
        f"{base}\n\n"
        "[기억] 아래는 이전 대화에서 알게 된 사용자·프로젝트에 관한 사실이다. "
        "관련될 때 참고하되, 확실하지 않으면 사용자에게 확인하라.\n"
        f"{block}"
    )


async def _maybe_autosummarize(state, conv: dict, provider: dict) -> None:
    """트리거가 맞으면 대화에서 사실을 추출해 장기 메모리에 저장한다 (쓰기·자동 경로).

    트리거는 '지난 요약 이후' 성장분 기준이다 — 턴이 interval만큼 늘었거나, 누적 이력
    문자수가 threshold만큼 늘었을 때. 그래서 임계값을 넘긴 뒤에도 매 턴 재실행되지 않는다.
    어떤 실패든 조용히 넘어가 대화 흐름을 깨지 않는다 (우아한 저하).
    """
    cfg = state.config
    if not cfg.get("memory_enabled", True) or not cfg.get("memory_autosummary_enabled", True):
        return
    if provider.get("mock"):
        return  # 목 모드는 실제 사실을 만들지 않는다

    msgs = conv.get("messages", [])
    turns = sum(1 for m in msgs if m.get("role") == "user")
    chars = sum(len(str(m.get("content") or "")) for m in msgs)
    last_turn = int(conv.get("memory_summarized_turn", 0))
    last_chars = int(conv.get("memory_summarized_chars", 0))
    interval = int(cfg.get("memory_autosummary_turn_interval", 25))
    threshold = int(cfg.get("memory_autosummary_char_threshold", 24000))

    by_turn = interval > 0 and (turns - last_turn) >= interval
    by_chars = threshold > 0 and (chars - last_chars) >= threshold
    if not (by_turn or by_chars):
        return

    try:
        facts = await agent.extract_memories(
            base_url=provider["base_url"],
            model=provider["model"],
            api_key=provider["api_key"],
            send_top_k=provider["send_top_k"],
            messages=msgs,
        )
        for fact in facts:
            state.memory.add(fact, kind="fact", source_conversation_id=conv["id"])
        if facts:
            print(f"[정보] 자동요약: 사실 {len(facts)}건 저장 (대화 {conv['id']})")
    except Exception as e:  # noqa: BLE001
        print(f"[주의] 자동요약 실패, 건너뜁니다: {e}")
    finally:
        # 성공/실패와 무관하게 이번 시점을 기록해 다음 트리거까지 재실행을 막는다.
        conv["memory_summarized_turn"] = turns
        conv["memory_summarized_chars"] = chars
        state.store.save(conv)


def _require_local(request: Request) -> None:
    """민감 API(파일 탐색, 위험 도구 승인)는 로컬 접속에서만 허용한다
    (0.0.0.0 공개 시 원격의 디스크 노출·승인 게이트 우회를 막는다)."""
    host = request.client.host if request.client else ""
    if host not in ("127.0.0.1", "::1", "localhost"):
        raise HTTPException(403, "이 기능은 이 PC(127.0.0.1)에서만 사용할 수 있습니다.")


def _native_open_dialog() -> str:
    """OS 기본 '열기' 대화상자를 띄워 선택한 절대경로를 돌려준다(취소 시 빈 문자열).

    별도 스레드(run_in_executor)에서 호출된다. tkinter는 해당 스레드 안에서
    루트를 만들고 파괴하므로 asyncio 메인 루프와 충돌하지 않는다. tkinter가 없거나
    데스크톱 세션이 아니면 예외가 나고, 호출측에서 안내 메시지로 물러선다.
    """
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)  # 브라우저 뒤로 숨지 않게 최상단으로
    try:
        path = filedialog.askopenfilename(
            title="문서 선택 (원본 제자리 읽기 — 복사 안 함)",
            filetypes=[
                ("문서", "*.docx *.doc *.xlsx *.xls *.xlsm *.pptx *.ppt"),
                ("모든 파일", "*.*"),
            ],
        )
    finally:
        root.destroy()
    return path or ""


# 확장자 → 그 문서를 읽는 MCP 도구를 찾을 때 쓸 키워드(도구 이름에 이게 들어감).
_EXT_TOOL_KEYWORDS = {
    ".xlsx": ("excel",), ".xls": ("excel",), ".xlsm": ("excel",), ".xlsb": ("excel",),
    ".pptx": ("powerpoint", "ppt"), ".ppt": ("powerpoint", "ppt"), ".pptm": ("powerpoint", "ppt"),
    ".docx": ("word",), ".doc": ("word",), ".docm": ("word",),
}
# 카테고리별로 '첫 호출'에 좋은 개요/본문 도구를 우선한다(여럿이면 이걸 먼저 집는다).
_PREFERRED_TOOLS = (
    "describe_excel", "read_excel_range",
    "read_word_document", "read_word_outline",
    "read_powerpoint_outline", "read_powerpoint_slides",
)


def _pick_read_tool(state, ext: str) -> str:
    """확장자에 맞는, 실제로 연결된 MCP 도구의 정식 이름(<서버>__<도구>)을 고른다. 없으면 ''.

    약한 모델(예: Gemma)은 '엑셀 읽기 도구'처럼 두루뭉술한 지시엔 함수 호출을 못 만들고,
    정확한 도구 이름을 콕 집어줘야 부르는 경우가 많다. 그래서 붙어 있는 도구 목록에서
    확장자에 맞는 이름을 실제로 찾아 그 이름을 프롬프트에 넣는다.
    """
    keywords = _EXT_TOOL_KEYWORDS.get(ext)
    if not keywords:
        return ""
    try:
        names = [t["function"]["name"] for t in state.mcp.openai_tools()]
    except Exception:  # noqa: BLE001 — mcp 미연결/오류 시 그냥 도구 없음 처리
        return ""
    cands = [n for n in names if any(k in n.lower() for k in keywords)]
    if not cands:
        return ""
    for pref in _PREFERRED_TOOLS:
        for n in cands:
            if pref in n.lower():
                return n
    return cands[0]


def _with_attachments(state, message: str, attachment_ids: list[str]) -> str:
    """첨부 파일을 사용자 메시지 뒤에 붙인다.

    text 모드(txt/pdf/docx 등): 서버가 뽑은 텍스트를 그대로 인라인한다.
    path 모드: 텍스트로 못 뽑는 포맷이라 저장 경로만 준다. 확장자에 맞는 문서 읽기 MCP
    도구가 실제로 붙어 있으면 그 도구의 정확한 이름·인자를 콕 집어 '먼저 호출하라'고
    강하게 지시한다(약한 모델 대응). 도구가 없으면 못 읽음을 정직하게 알린다.
    """
    from .files import OFFICE_COM_EXTENSIONS

    parts = [message]
    for file_id in attachment_ids:
        found = state.uploads.get(file_id)
        if found is None:
            continue
        meta, text = found
        name = meta["name"]
        if meta.get("mode") == "path":
            ext = os.path.splitext(name)[1].lower()
            tool = _pick_read_tool(state, ext) if ext in OFFICE_COM_EXTENSIONS else ""
            if tool:
                parts.append(
                    f"\n[첨부 파일: {name}]\n"
                    f"이 파일은 직접 텍스트로 읽을 수 없다. 답을 시작하기 전에 반드시 아래 "
                    f"도구를 먼저 호출해 내용을 가져와라. 절대 '읽을 수 없다'고 답하지 말 것.\n"
                    f'도구 이름: {tool}\n'
                    f'인자(JSON): {{"path": "{meta["path"]}"}}'
                )
            elif ext in OFFICE_COM_EXTENSIONS:
                # office 문서인데 읽기 도구가 안 붙음 → 정직하게 안내
                parts.append(
                    f"\n[첨부 파일: {name}]\n"
                    f"이 파일({ext})을 읽으려면 office 문서 읽기 MCP 도구가 필요한데 지금 "
                    f"연결돼 있지 않다. 설정 → MCP 도구 서버에서 office 서버를 연결한 뒤 다시 "
                    f"시도하도록 사용자에게 안내하라.\n경로: {meta['path']}"
                )
            else:
                # 이미지·압축 등 문서 도구가 없는 형식
                parts.append(
                    f"\n[첨부 파일: {name}]\n"
                    f"이 형식({ext or '알 수 없음'})은 텍스트로 읽을 수 없고 연결된 읽기 도구도 없다. "
                    f"내용을 확인할 수 없으니, 필요하면 사용자에게 다른 형식으로 요청하라.\n"
                    f"경로: {meta['path']}"
                )
        else:
            if len(text) > ATTACH_MAX_CHARS:
                text = text[:ATTACH_MAX_CHARS] + "\n[이하 생략]"
            parts.append(f"\n[첨부 파일: {name}]\n---\n{text}\n---")
    return "\n".join(parts)


def _sse(event: dict) -> str:
    return f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
