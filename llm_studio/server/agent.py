"""agent.py

llama-server(OpenAI 호환 API)에 스트리밍 요청을 보내고, 모델이 도구 호출을
요구하면 MCP 도구를 실행해 결과를 넣어 다시 호출하는 루프를 돈다.

run_chat()은 이벤트 dict를 순서대로 내보내는 async generator다:
    {"type": "reasoning",   "text": "..."}                  모델의 생각(추론) 조각
    {"type": "token",       "text": "..."}                  응답 토큰 조각
    {"type": "tool_call",   "name": "...", "arguments": "...", "executed": bool}
                                                            도구 호출 (executed=False면
                                                            승인 거절/시간초과로 실행 안 됨)
    {"type": "tool_result", "name": "...", "result": "...", "executed": bool}
                                                            도구 실행 결과 (위와 동일)
    {"type": "approval_request", "id": "...", "name": "...", "arguments": "..."}
                                                            위험 도구 — 사용자 승인 대기
    {"type": "approval_result",  "id": "...", "approved": bool, "timeout": bool}
                                                            승인/거절/시간초과 결정
    {"type": "done",        "messages": [...]}              완료 (누적 메시지 전체)
    {"type": "error",       "message": "..."}               오류

reasoning: llama-server가 추론형 모델의 <think> 블록을 reasoning_content로 분리해
주면(예: DeepSeek-R1/Qwen3 계열) 그 조각을 흘린다. Gemma처럼 추론을 따로 내보내지
않는 모델에서는 이 이벤트가 나오지 않는다(그냥 답변 토큰만).
"""

from __future__ import annotations

import asyncio
import json
from typing import AsyncIterator

from openai import AsyncOpenAI, NOT_GIVEN

from .config import DEFAULT_CONFIG

MAX_TOOL_ROUNDS = 8       # 무한 도구 루프 방지 (settings의 max_tool_rounds가 없을 때 기본값)
TOOL_RESULT_MAX = 20000   # 도구 결과가 컨텍스트를 다 먹지 않게 자르는 한도 (문자)

# 작업 모드(planner)가 스텝 텍스트로 감싸지 않고 그대로 통과시켜야 하는 전역 UI 이벤트.
# 새 전역 이벤트 타입(질문·진행바 등)을 추가할 때 여기에만 올리면 채팅/작업 모드
# 양쪽에서 동작한다 — planner 쪽 if 분기에 흩어 놓지 말 것.
PASSTHROUGH_EVENTS = {"approval_request", "approval_result"}

# 승인 거절/시간초과 시 모델에 돌려주는 도구 결과. 저장된 대화를 다시 열 때 UI가
# 이 접두사로 '실행 안 됨'을 판별하므로(app.js isDeniedToolResult) 문구를 바꾸면
# 접두사를 양쪽에서 함께 바꿔야 한다.
DENIED_RESULT = (
    "[사용자가 이 도구 실행을 거부했습니다. 실행되지 않았습니다. "
    "강행하지 말고, 대안을 제시하거나 이유를 물어보세요.]"
)
TIMEOUT_RESULT = (
    "[사용자 승인 대기가 시간 초과되어 실행하지 않았습니다. "
    "필요하면 사용자에게 다시 요청할지 물어보세요.]"
)

# 내장 도구: 장기 메모리에 사실을 저장한다 (MCP가 아니라 앱 내부 처리).
# 읽기(회상)는 하네스가 자동 주입하므로 recall 도구는 두지 않는다 — 쓰기만 노출한다.
REMEMBER_TOOL_SPEC = {
    "type": "function",
    "function": {
        "name": "remember",
        "description": (
            "앞으로의 대화에서도 계속 유효한 사용자·프로젝트에 관한 사실을 장기 메모리에 "
            "저장한다. 사용자의 선호, 배경, 진행 중인 작업처럼 지속적으로 참일 정보에만 쓴다. "
            "일회성 잡담이나 이번 대화에만 의미 있는 내용은 저장하지 않는다."
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "content": {"type": "string", "description": "기억할 사실 한 문장 (한국어)"},
                "kind": {
                    "type": "string",
                    "description": "분류",
                    "enum": ["user", "preference", "project", "fact"],
                },
            },
            "required": ["content"],
        },
    },
}


async def run_chat(
    *,
    base_url: str,
    model: str,
    messages: list[dict],
    settings: dict,
    api_key: str = "local",
    send_top_k: bool = True,
    mcp=None,
    memory=None,
    mock: bool = False,
    tool_servers=None,
    approver=None,
) -> AsyncIterator[dict]:
    """대화 한 턴을 실행한다. messages는 system 포함 전체 이력.

    base_url은 /v1까지 포함한 OpenAI 호환 엔드포인트 전체 주소.
    send_top_k: llama-server는 top_k를 받지만 OpenAI 등 외부 API는
    알 수 없는 파라미터로 거절하므로 로컬일 때만 보낸다.
    memory: 주어지면 내장 remember 도구를 노출해 모델이 사실을 저장할 수 있게 한다.
    tool_servers: MCP 도구를 이 서버들로만 좁힌다(서버-스코프). None이면 전체,
    빈 목록이면 MCP 도구 없음. 작업 모드가 스텝마다 필요한 서버만 넘길 때 쓴다.
    approver: ApprovalBroker. 주어지면 위험 도구 호출(confirm=true 인자 또는
    settings['approval_tools']에 오른 이름) 전에 approval_request 이벤트를 흘리고
    사용자의 승인/거절을 기다린다. 거절·시간초과면 실행하지 않고 그 사실을 도구
    결과로 모델에 알린다.
    """
    if mock:
        async for event in _mock_stream(messages):
            yield event
        return

    client = AsyncOpenAI(base_url=base_url, api_key=api_key or "none", timeout=600)
    msgs = [dict(m) for m in messages]
    tool_specs = list(mcp.openai_tools(servers=tool_servers)) if mcp else []
    if memory is not None:
        tool_specs.append(REMEMBER_TOOL_SPEC)  # MCP 도구 옆에 내장 쓰기 도구를 더한다

    # 도구 라운드 상한: 설정에서 조정할 수 있다 (최소 1로 보정 — 0이면 응답 자체가 불가).
    try:
        max_rounds = max(1, int(settings.get("max_tool_rounds", MAX_TOOL_ROUNDS)))
    except (TypeError, ValueError):
        max_rounds = MAX_TOOL_ROUNDS

    try:
        for round_no in range(max_rounds):
            stream = await client.chat.completions.create(
                model=model,
                messages=msgs,
                stream=True,
                tools=tool_specs or NOT_GIVEN,
                temperature=settings.get("temperature", 1.0),
                top_p=settings.get("top_p", 0.95),
                max_tokens=settings.get("max_tokens", 4096),
                extra_body={"top_k": settings.get("top_k", 64)} if send_top_k else None,
            )

            content_parts: list[str] = []
            tool_calls: dict[int, dict] = {}  # index -> {id, name, args}
            async for chunk in stream:
                if not chunk.choices:
                    continue
                delta = chunk.choices[0].delta
                # 추론형 모델의 생각 조각(비표준 필드 reasoning_content). pydantic은
                # 알 수 없는 필드를 model_extra에 담으므로 양쪽을 다 확인한다.
                rc = getattr(delta, "reasoning_content", None)
                if rc is None:
                    rc = (getattr(delta, "model_extra", None) or {}).get("reasoning_content")
                if rc:
                    yield {"type": "reasoning", "text": rc}
                if delta.content:
                    content_parts.append(delta.content)
                    yield {"type": "token", "text": delta.content}
                for tc in delta.tool_calls or []:
                    slot = tool_calls.setdefault(tc.index, {"id": "", "name": "", "args": ""})
                    if tc.id:
                        slot["id"] = tc.id
                    if tc.function:
                        slot["name"] += tc.function.name or ""
                        slot["args"] += tc.function.arguments or ""

            content = "".join(content_parts)
            if not tool_calls:
                msgs.append({"role": "assistant", "content": content})
                yield {"type": "done", "messages": msgs}
                return

            # 모델이 도구 호출을 요구했다: assistant 메시지 기록 후 각 도구 실행
            calls = [tool_calls[i] for i in sorted(tool_calls)]
            for i, call in enumerate(calls):
                call["id"] = call["id"] or f"call_{round_no}_{i}"
            msgs.append({
                "role": "assistant",
                "content": content or None,
                "tool_calls": [
                    {
                        "id": c["id"],
                        "type": "function",
                        "function": {"name": c["name"], "arguments": c["args"] or "{}"},
                    }
                    for c in calls
                ],
            })
            for call in calls:
                # 위험 도구면 실행 전에 사용자 승인을 받는다 (브라우저에 버튼 표시).
                approved, timed_out = True, False
                if approver is not None and _needs_approval(call["name"], call["args"], settings):
                    # 시간 제한은 create() 전에 확정한다 — create() 뒤에서 예외가 나면
                    # 대기 항목이 누수된다. 숫자가 아니면 기본값으로 물러서고,
                    # 0 이하 = 무제한 대기 판정은 브로커 wait() 한 곳이 담당한다.
                    try:
                        timeout = float(settings.get(
                            "approval_timeout", DEFAULT_CONFIG["approval_timeout"]))
                    except (TypeError, ValueError):
                        timeout = float(DEFAULT_CONFIG["approval_timeout"])
                    req_id = approver.create()
                    try:
                        yield {"type": "approval_request", "id": req_id,
                               "name": call["name"], "arguments": call["args"]}
                        approved, timed_out = await approver.wait(req_id, timeout)
                    finally:
                        # yield 지점에서 스트림이 중단(GeneratorExit)돼 wait()에 못
                        # 들어가도 대기 항목이 남지 않게 정리한다 (discard는 멱등).
                        approver.discard(req_id)
                    yield {"type": "approval_result", "id": req_id,
                           "approved": approved, "timeout": timed_out}
                yield {"type": "tool_call", "name": call["name"],
                       "arguments": call["args"], "executed": approved}
                if not approved:
                    result = TIMEOUT_RESULT if timed_out else DENIED_RESULT
                else:
                    result = await _execute_tool(
                        mcp, call["name"], call["args"], memory, approved=True)
                yield {"type": "tool_result", "name": call["name"],
                       "result": result[:4000], "executed": approved}
                msgs.append({"role": "tool", "tool_call_id": call["id"], "content": result})

        yield {
            "type": "error",
            "message": (
                f"도구 호출이 {max_rounds}회를 초과해 중단했습니다. "
                "설정 → 생성 파라미터의 '도구 호출 상한'에서 늘릴 수 있습니다."
            ),
        }
    except asyncio.CancelledError:
        raise  # 사용자가 중단 버튼을 누른 경우: 상위에서 부분 응답을 저장한다
    except Exception as e:  # noqa: BLE001 — 모든 오류를 UI에 이벤트로 전달
        yield {"type": "error", "message": f"{type(e).__name__}: {e}"}


# MCP 서버 쪽(pydantic lax bool)이 False로 해석하는 문자열들. 약한 모델이 confirm에
# "false"/"0" 같은 문자열 불리언을 내면 서버는 프리뷰만 돌려줄 호출인데 클라이언트
# 게이트만 참으로 읽어 불필요한 승인 카드가 뜨는 어긋남을 막는다.
_FALSY_STRINGS = {"", "0", "false", "no", "off", "none", "null"}


def _confirm_truthy(value) -> bool:
    """confirm 인자를 MCP 서버(pydantic lax bool)와 같은 방식으로 해석한다."""
    if isinstance(value, str):
        return value.strip().lower() not in _FALSY_STRINGS
    return bool(value)


def _needs_approval(name: str, raw_args: str, settings: dict) -> bool:
    """이 도구 호출이 사용자 승인을 요구하는지 판정한다.

    1) 인자에 confirm이 참으로 들어 있으면 — MCP 서버들의 🔴 confirm 게이트 규약상
       '파괴적 동작의 실제 실행'이므로 항상 승인 대상 (이중 안전장치).
    2) settings['approval_tools'] 목록에 오른 이름이면 — 서버 접두사가 붙은
       전체 이름(outlook__send_email)과 짧은 이름(send_email) 둘 다 인정한다.
    """
    if not settings.get("approval_enabled", True):
        return False
    try:
        args = json.loads(raw_args or "{}")
    except json.JSONDecodeError:
        args = {}
    if isinstance(args, dict) and _confirm_truthy(args.get("confirm")):
        return True
    listed = settings.get("approval_tools") or []
    short = name.split("__", 1)[-1]
    return name in listed or short in listed


async def _execute_tool(mcp, name: str, raw_args: str, memory=None,
                        *, approved: bool = False) -> str:
    try:
        args = json.loads(raw_args or "{}")
    except json.JSONDecodeError as e:
        return f"[도구 실행 실패: 인자 JSON 파싱 오류 — {e}]"
    # fail-closed: confirm이 참인 호출은 run_chat의 승인 게이트를 거쳤다는 표식
    # (approved=True) 없이는 실행하지 않는다. 미래에 이 함수를 직접 부르는 경로가
    # 생겨도 파괴적 동작이 승인 없이 조용히 실행되는 일을 구조적으로 막는다.
    if not approved and isinstance(args, dict) and _confirm_truthy(args.get("confirm")):
        return ("[도구 실행 거부: confirm이 참인 호출은 승인 게이트를 거쳐야 합니다 "
                "(호출부가 approved=True를 명시하지 않음)]")
    if name == "remember":  # 내장 메모리 쓰기 도구 (MCP가 아니라 앱 내부 처리)
        return _do_remember(memory, args)
    if mcp is None:
        return "[도구 실행 실패: MCP가 연결되어 있지 않습니다]"
    try:
        result = await mcp.call(name, args)
    except Exception as e:  # noqa: BLE001
        return f"[도구 실행 실패: {type(e).__name__}: {e}]"
    return result[:TOOL_RESULT_MAX]


def _do_remember(memory, args: dict) -> str:
    """내장 remember 도구 실행: 장기 메모리에 사실 하나를 저장한다."""
    if memory is None:
        return "[기억 실패: 메모리가 비활성화되어 있습니다]"
    content = (args.get("content") or "").strip()
    if not content:
        return "[기억 실패: 저장할 내용(content)이 비어 있습니다]"
    kind = args.get("kind") or "fact"
    try:
        mem_id = memory.add(content, kind=kind)
    except Exception as e:  # noqa: BLE001 — 저장 실패가 대화를 끊지 않게 한다
        return f"[기억 실패: {type(e).__name__}: {e}]"
    if mem_id is None:
        return "[기억 실패: 빈 내용]"
    return f"기억했습니다: {content}"


SUMMARY_MAX_MESSAGES = 30    # 자동요약에 넣을 최근 메시지 수
SUMMARY_MSG_CHARS = 800      # 메시지 하나당 자르는 한도

_SUMMARY_SYSTEM = (
    "너는 대화에서 앞으로도 계속 유효할 사용자·프로젝트에 관한 사실만 뽑아내는 추출기다. "
    "사용자의 선호·배경·환경·진행 중인 작업처럼 다음 대화에서도 참일 정보만 남긴다. "
    "일회성 질문, 이번 대화에만 의미 있는 내용, 모델의 일반 지식은 제외한다. "
    "각 사실은 독립적으로 이해되는 한 문장으로 쓴다. "
    "남길 사실이 없으면 정확히 '없음'이라고만 답한다. "
    "있으면 한 줄에 하나씩, 각 줄을 '- '로 시작한다. 다른 말은 하지 않는다."
)


def _render_transcript(messages: list[dict]) -> str:
    """자동요약용으로 최근 user/assistant 발화를 간단한 대화록 텍스트로 만든다."""
    lines = []
    for m in messages[-SUMMARY_MAX_MESSAGES:]:
        role = m.get("role")
        if role not in ("user", "assistant"):
            continue  # system/tool 메시지는 사실 추출에 방해되므로 뺀다
        content = str(m.get("content") or "").strip()
        if not content:
            continue
        who = "사용자" if role == "user" else "어시스턴트"
        lines.append(f"{who}: {content[:SUMMARY_MSG_CHARS]}")
    return "\n".join(lines)


async def extract_memories(
    *, base_url: str, model: str, messages: list[dict],
    api_key: str = "local", send_top_k: bool = True,
) -> list[str]:
    """대화에서 장기 기억으로 남길 사실들을 뽑아 문자열 리스트로 반환한다.

    별도의 비스트리밍 LLM 호출 1회. 어떤 실패든(네트워크·형식) 빈 리스트로 물러서
    호출부가 대화 흐름을 잃지 않게 한다.
    """
    transcript = _render_transcript(messages)
    if not transcript:
        return []
    client = AsyncOpenAI(base_url=base_url, api_key=api_key or "none", timeout=120)
    try:
        resp = await client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": _SUMMARY_SYSTEM},
                {"role": "user", "content": f"다음 대화에서 사실을 추출하라:\n\n{transcript}"},
            ],
            stream=False,
            temperature=0.3,  # 추출은 결정적일수록 좋다
            max_tokens=800,
            extra_body={"top_k": 20} if send_top_k else None,
        )
        content = (resp.choices[0].message.content or "").strip()
    except Exception:  # noqa: BLE001 — 요약 실패는 조용히 넘어간다
        return []
    return _parse_facts(content)


def _parse_facts(text: str) -> list[str]:
    """모델 출력에서 '- '로 시작하는 사실 줄만 뽑는다. '없음'이면 빈 리스트."""
    if text.strip() in ("없음", "- 없음"):
        return []
    facts = []
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("- "):
            fact = line[2:].strip()
            if fact and fact != "없음":
                facts.append(fact)
    return facts


async def _mock_stream(messages: list[dict]) -> AsyncIterator[dict]:
    """llama-server 없이 UI를 개발/시험하기 위한 목(mock) 응답."""
    last_user = next((m["content"] for m in reversed(messages) if m["role"] == "user"), "")
    text = (
        "**목(mock) 모드 응답입니다.** 실제 모델이 연결되어 있지 않습니다 (--mock).\n\n"
        f"입력하신 내용: {str(last_user)[:200]}\n\n"
        "실제 모델을 쓰려면 `--mock` 없이 실행한 뒤:\n"
        "1. `models/` 폴더에 GGUF 파일을 넣고 (또는 이미 있다면)\n"
        "2. 상단의 [서빙 시작] 또는 설정 → 로컬 LLM 서버에서 모델을 골라 서빙을 시작하세요.\n"
        "3. 외부 API를 쓰려면 설정에서 등록 후 헤더 드롭다운에서 선택하세요.\n\n"
        "```python\n# 코드 블록 렌더링 테스트\nprint('hello')\n```\n"
    )
    for word in text.split(" "):
        yield {"type": "token", "text": word + " "}
        await asyncio.sleep(0.02)
    msgs = [dict(m) for m in messages] + [{"role": "assistant", "content": text}]
    yield {"type": "done", "messages": msgs}
