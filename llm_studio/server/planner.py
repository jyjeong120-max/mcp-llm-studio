"""planner.py — 계획-실행(plan-and-execute) 에이전트.

약한 로컬 모델(Gemma 12B 등)은 긴 도구 루프에서 "지금 뭘 하는 중인지"를 쉽게 잃는다.
그래서 요청을 먼저 **계획**으로 외재화하고, 하네스가 그 계획을 붙든 채 스텝마다
좁은 컨텍스트만 주며 하나씩 **실행**한 뒤 결과를 **종합**한다. 판단을 모델의 즉흥적
추론이 아니라 하네스의 제어 흐름이 쥐는 구조다.

run_task()는 run_chat()과 같은 방식으로 이벤트 dict를 내보내는 async generator다:
    {"type": "plan",       "steps": [...], "replan": n?}   계획(또는 재계획) 수립
    {"type": "step_start", "index": i, "text": "..."}      한 스텝 실행 시작
    {"type": "step_token", "index": i, "text": "..."}      스텝 진행 중 출력 조각(도구 포함)
    {"type": "step_done",  "index": i, "ok": bool, "result": "..."}
    {"type": "token",      "text": "..."}                  최종 종합 답변 조각
    {"type": "done",       "messages": [...]}              완료 (대화에 저장할 메시지)
    {"type": "error",      "message": "..."}

각 스텝 실행은 기존 agent.run_chat()을 그대로 재사용한다(도구 루프 포함). 스텝 안에서
나오는 토큰·도구 이벤트는 전부 step_token으로 감싸 해당 스텝 블록에만 흐르게 하고,
최종 답변(종합)만 일반 token으로 흘려 대화의 답변 영역에 렌더링되게 한다.

단기(작업) 상태는 RAM에만 둔다 — 이 요청이 사는 동안만 유효하며 디스크에 남기지 않는다.
재시작·크래시 시 진행 중 작업은 사라진다(대화형 전제). 무한 루프 방지를 위해
max_steps(총 실행 상한)와 max_replans(재계획 예산)를 둔다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import AsyncIterator

from openai import AsyncOpenAI

from . import agent

DEFAULT_MAX_STEPS = 10
DEFAULT_MAX_REPLANS = 2

_PLAN_SYSTEM = (
    "너는 사용자의 요청을 실행 가능한 단계로 쪼개는 계획가다. "
    "각 단계는 한 줄로, 구체적이고 순서대로 수행 가능해야 한다. 3~7단계로 간결하게 세운다. "
    "형식: 각 줄을 '1. ', '2. '처럼 번호로 시작한다. 설명이나 다른 말 없이 단계 목록만 출력한다."
)
_STEP_SYSTEM = (
    "너는 더 큰 작업의 한 단계를 수행하는 실행자다. 전체 목표와 계획, 지금까지의 결과가 주어진다. "
    "'지금 수행할 단계'만 처리하고 그 결과를 간결히 보고한다. 도구가 필요하면 사용한다. "
    "그 단계를 완료할 수 없으면 마지막 줄에 정확히 '[실패] 사유' 형식으로 적는다."
)
_REPLAN_SYSTEM = (
    "너는 계획을 수정하는 계획가다. 목표와 원래 계획, 지금까지의 결과, 그리고 방금 실패한 "
    "단계와 사유가 주어진다. 남은 목표를 이루기 위한 '앞으로의 단계'만 새로 세운다. "
    "'사용자 거절'로 표시된 동작은 사용자가 하지 않기로 결정한 것이므로 새 계획에 다시 "
    "넣지 않는다. "
    "형식: 각 줄을 '1. '처럼 번호로 시작한다. 다른 말 없이 단계 목록만 출력한다."
)
_SYNTH_SYSTEM = (
    "너는 작업의 단계별 결과를 종합해 사용자에게 최종 답변을 제시한다. "
    "단계 결과를 바탕으로 목표에 대한 완결된 한국어 답변을 작성한다. "
    "'사용자 거절'로 표시된 단계는 수행되지 않았음을 분명히 밝히고, 수행된 것처럼 "
    "서술하지 않는다. "
    "내부 단계나 계획 자체에 대한 언급은 최소화하고, 사용자가 원한 결과를 직접 제시한다."
)

_STEP_RE = re.compile(r"^\s*\d+[.)]\s*(.+)$")


@dataclass
class TaskState:
    """단기(작업) 메모리 — RAM에만 존재한다. 계획과 진행 상황을 담는다."""
    goal: str
    plan: list[str]
    # 각 단계의 도구 스코프. plan과 같은 길이로 유지한다(재계획 시 함께 스플라이스).
    # 원소: 서버 이름 리스트(그 서버 툴만) / [] (도구 없음) / None (태그 없음 → 전체 툴).
    servers: list = field(default_factory=list)
    results: list[str] = field(default_factory=list)  # 완료된 스텝의 결과 텍스트
    executed: int = 0     # 총 스텝 실행 횟수(재시도 포함) — max_steps 상한 대상
    replans: int = 0      # 사용한 재계획 횟수 — max_replans 예산 대상


def _parse_steps(text: str) -> list[str]:
    """번호 매긴 목록 텍스트에서 각 단계 문장을 뽑는다."""
    steps = []
    for line in (text or "").splitlines():
        m = _STEP_RE.match(line)
        if m:
            step = m.group(1).strip()
            if step:
                steps.append(step)
    return steps


def _last_user(messages: list[dict]) -> str:
    return next((str(m.get("content") or "") for m in reversed(messages)
                 if m.get("role") == "user"), "")


_HISTORY_MAX_MSGS = 10    # 계획/종합에 넣을 직전 대화 발화 수
_HISTORY_MSG_CHARS = 600  # 발화 하나당 자르는 한도
# 스텝 프롬프트는 목표+계획+누적결과로 이미 크다. 여기에 전체 이력까지 매 스텝 붙이면
# 좁은 ctx(예: 8192)에서 앞부분이 잘려 정작 '수행할 단계' 지시가 날아갈 수 있어,
# 스텝에는 훨씬 짧은 이력만 넣는다(맥락 유지 + ctx 초과 방지의 절충).
_HISTORY_STEP_MAX_MSGS = 4
_HISTORY_STEP_MSG_CHARS = 200


def _history_context(messages: list[dict], *, max_msgs: int = _HISTORY_MAX_MSGS,
                     msg_chars: int = _HISTORY_MSG_CHARS) -> str:
    """직전 대화를 간단한 대화록으로 만든다(현재 목표 메시지는 제외). 없으면 ''.

    작업 모드는 스텝별로 좁은 컨텍스트만 주지만, 그렇다고 이전 대화를 통째로 버리면
    다단계 요청이 앞선 맥락(내 이름, 방금 얘기한 파일 등)을 잊는다. 그래서 계획·스텝·
    종합 프롬프트에 이 대화록을 함께 넣어 다중 턴 맥락을 유지한다. max_msgs/msg_chars로
    분량을 줄일 수 있다(스텝 프롬프트는 축약본을 쓴다).
    """
    convo = [m for m in messages if m.get("role") in ("user", "assistant")]
    if convo and convo[-1].get("role") == "user":
        convo = convo[:-1]  # 마지막 user는 이번 목표이므로 뺀다
    lines = []
    for m in convo[-max_msgs:]:
        content = str(m.get("content") or "").strip()
        if not content:
            continue
        who = "사용자" if m.get("role") == "user" else "어시스턴트"
        lines.append(f"{who}: {content[:msg_chars]}")
    return "\n".join(lines)


async def _complete(client: AsyncOpenAI, model: str, system: str, user: str,
                    *, send_top_k: bool, max_tokens: int = 700) -> str:
    """비스트리밍 단발 completion. 계획/재계획처럼 형식만 필요한 호출에 쓴다."""
    resp = await client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": system},
                  {"role": "user", "content": user}],
        stream=False,
        temperature=0.3,  # 계획은 결정적일수록 좋다
        max_tokens=max_tokens,
        extra_body={"top_k": 20} if send_top_k else None,
    )
    return (resp.choices[0].message.content or "").strip()


_TAG_RE = re.compile(r"^\[([^\]]*)\]\s*(.*)$")


def _server_menu(mcp):
    """계획가에게 줄 '도구 서버 메뉴'와 유효한 서버 이름 목록을 만든다.

    반환: (프롬프트에 붙일 메뉴 문자열, 연결된 서버 이름 리스트). 서버가 없으면
    ("", []) — 이때는 계획에 서버 태그를 요구하지 않는다(기존과 동일하게 동작).

    서버-스코프의 핵심: 계획가가 각 단계에 필요한 서버를 [이름]으로 표시하게 하고,
    실행 때 그 서버 도구만 노출해 약한 모델이 좁은 도구셋에서 정확히 고르게 한다.
    """
    if mcp is None:
        return "", []
    try:
        servers = mcp.connected_servers()
    except Exception:  # noqa: BLE001
        servers = []
    if not servers:
        return "", []
    lines = []
    for s in servers:
        try:
            names = [t["function"]["name"].split("__", 1)[-1]
                     for t in mcp.openai_tools(servers=[s])]
        except Exception:  # noqa: BLE001
            names = []
        summary = ", ".join(names[:10]) + ("…" if len(names) > 10 else "")
        lines.append(f"  - {s}: {summary}")
    menu = (
        "\n\n사용 가능한 도구 서버(각 단계에 필요한 서버를 번호 뒤 [이름]으로 표시하고, "
        "도구가 필요 없으면 [-]):\n" + "\n".join(lines) +
        "\n예: '1. [office] 엑셀 표를 읽는다'"
    )
    return menu, servers


def _split_tag(step_text: str, valid: list[str]):
    """단계 앞머리의 [서버] 태그를 떼어 (본문, 스코프)로 나눈다.

    스코프: 유효 서버 리스트 / [] (도구 불필요) / None (태그 없음·무효 → 전체 툴).
    태그 형식이 어긋나도 전체 툴로 우아하게 저하한다(약한 모델의 형식 실수 대비).
    """
    m = _TAG_RE.match(step_text.strip())
    if not m:
        return step_text.strip(), None
    raw = m.group(1).strip()
    body = m.group(2).strip() or step_text.strip()
    if raw in ("-", "", "없음", "none", "None"):
        return body, []
    picks = [p.strip() for p in re.split(r"[,\s/]+", raw) if p.strip()]
    valid_picks = [p for p in picks if p in valid]
    return body, (valid_picks or None)


def _scope_suffix(scope) -> str:
    """스텝 표시에 붙일 도구 스코프 라벨(투명성용). None이면 표시 없음."""
    if scope is None:
        return ""
    if not scope:
        return "  ⟨도구 없음⟩"
    return f"  ⟨{', '.join(scope)}⟩"


async def _run_step(step_idx: int, prompt_messages: list[dict], *, base_url, model,
                    settings, api_key, send_top_k, mcp, memory,
                    tool_servers=None, approver=None) -> AsyncIterator[dict]:
    """한 스텝을 agent.run_chat으로 실행하며 진행 이벤트를 흘린다.

    run_chat의 토큰·도구 이벤트를 전부 step_token으로 감싼다(해당 스텝 블록에만 표시).
    마지막에 {"__result__": text, "__ok__": bool}를 담은 이벤트를 하나 내보내
    호출부가 스텝 결과/성공 여부를 받게 한다.

    tool_servers: 이 스텝에 노출할 MCP 서버 스코프(서버-스코프 라우팅). None이면
    전체, 빈 목록이면 도구 없음. run_chat에 그대로 넘긴다.
    """
    final_text = ""
    ok = True
    denied = False
    async for ev in agent.run_chat(
        base_url=base_url, model=model, messages=prompt_messages, settings=settings,
        api_key=api_key, send_top_k=send_top_k, mcp=mcp, memory=memory, mock=False,
        tool_servers=tool_servers, approver=approver,
    ):
        etype = ev.get("type")
        if etype in agent.PASSTHROUGH_EVENTS:
            # 전역 UI 이벤트는 스텝 텍스트로 감싸지 않고 그대로 — 버튼이 떠야 한다.
            # 거절/시간초과는 텍스트가 아니라 구조적 신호로 상위(run_task)에 전달한다
            # — 모델의 스텝 요약 문구('[실패]' 유무)에 의존하면 거절이 성공으로 둔갑한다.
            if etype == "approval_result" and not ev.get("approved"):
                denied = True
            yield ev
        elif etype in ("token", "reasoning"):
            yield {"type": "step_token", "index": step_idx, "text": ev["text"]}
        elif etype == "tool_call":
            mark = "" if ev.get("executed", True) else " (거절됨 — 실행 안 함)"
            yield {"type": "step_token", "index": step_idx,
                   "text": f"\n🔧 도구 호출: {ev['name']}{mark} {ev.get('arguments', '')}\n"}
        elif etype == "tool_result":
            yield {"type": "step_token", "index": step_idx,
                   "text": f"↳ 결과: {str(ev.get('result', ''))[:400]}\n"}
        elif etype == "done":
            msgs = ev.get("messages", [])
            final_text = next((str(m.get("content") or "") for m in reversed(msgs)
                               if m.get("role") == "assistant"), "")
        elif etype == "error":
            ok = False
            yield {"type": "step_token", "index": step_idx,
                   "text": f"\n[오류] {ev.get('message', '')}\n"}
    if "[실패]" in final_text:
        ok = False
    yield {"__result__": final_text.strip(), "__ok__": ok, "__denied__": denied}


async def run_task(
    *,
    base_url: str,
    model: str,
    messages: list[dict],
    settings: dict,
    api_key: str = "local",
    send_top_k: bool = True,
    mcp=None,
    memory=None,
    max_steps: int = DEFAULT_MAX_STEPS,
    max_replans: int = DEFAULT_MAX_REPLANS,
    approver=None,
) -> AsyncIterator[dict]:
    """계획-실행으로 한 턴을 처리한다. messages는 system 포함 전체 이력.

    계획을 못 세우면(파싱 0개/오류) 일반 run_chat으로 우아하게 저하한다.
    """
    goal = _last_user(messages)
    history = _history_context(messages)
    hist_block = f"이전 대화:\n{history}\n\n" if history else ""
    # 스텝 프롬프트용 축약 이력 (ctx 초과 방지 — 위 _HISTORY_STEP_* 참고)
    history_step = _history_context(
        messages, max_msgs=_HISTORY_STEP_MAX_MSGS, msg_chars=_HISTORY_STEP_MSG_CHARS)
    hist_block_step = f"이전 대화(요약):\n{history_step}\n\n" if history_step else ""
    client = AsyncOpenAI(base_url=base_url, api_key=api_key or "none", timeout=600)
    # 서버-스코프: 계획가에게 도구 서버 메뉴를 보여주고 단계마다 [서버] 태그를 받는다.
    # 서버가 없으면 메뉴는 빈 문자열 → 태그를 요구하지 않아 기존 동작과 같아진다.
    server_menu, valid_servers = _server_menu(mcp)

    # ---------- 계획 ----------
    try:
        plan_text = await _complete(
            client, model, _PLAN_SYSTEM,
            f"{hist_block}요청: {goal}{server_menu}", send_top_k=send_top_k)
        raw_steps = _parse_steps(plan_text)[:max_steps]
    except Exception as e:  # noqa: BLE001
        raw_steps = []
        yield {"type": "step_token", "index": -1, "text": f"[계획 실패: {e}]"}

    # 각 단계에서 [서버] 태그를 떼어 본문과 도구 스코프로 나눈다.
    plan, servers = [], []
    for rs in raw_steps:
        body, scope = _split_tag(rs, valid_servers)
        plan.append(body)
        servers.append(scope)

    if not plan:
        # 계획을 못 세움 → 일반 채팅으로 저하 (도구 포함). 이벤트를 그대로 흘린다.
        async for ev in agent.run_chat(
            base_url=base_url, model=model, messages=messages, settings=settings,
            api_key=api_key, send_top_k=send_top_k, mcp=mcp, memory=memory, mock=False,
            approver=approver,
        ):
            yield ev
        return

    state = TaskState(goal=goal, plan=plan, servers=servers)
    yield {"type": "plan", "steps": list(state.plan)}

    # ---------- 실행 루프 ----------
    i = 0
    while i < len(state.plan) and state.executed < max_steps:
        step = state.plan[i]
        scope = state.servers[i] if i < len(state.servers) else None
        state.executed += 1
        yield {"type": "step_start", "index": i, "text": step + _scope_suffix(scope)}

        prior = "\n".join(f"- {r}" for r in state.results) or "(없음)"
        plan_str = "\n".join(f"{n + 1}. {s}" for n, s in enumerate(state.plan))
        step_user = (
            f"{hist_block_step}전체 목표: {state.goal}\n\n계획:\n{plan_str}\n\n"
            f"지금까지의 결과:\n{prior}\n\n"
            f"지금 수행할 단계 ({i + 1}/{len(state.plan)}): {step}"
        )
        result_text, ok, denied = "", True, False
        async for ev in _run_step(
            i, [{"role": "system", "content": _STEP_SYSTEM},
                {"role": "user", "content": step_user}],
            base_url=base_url, model=model, settings=settings, api_key=api_key,
            send_top_k=send_top_k, mcp=mcp, memory=memory, tool_servers=scope,
            approver=approver,
        ):
            if "__result__" in ev:
                result_text, ok = ev["__result__"], ev["__ok__"]
                denied = ev.get("__denied__", False)
            else:
                yield ev

        yield {"type": "step_done", "index": i, "ok": ok and not denied,
               "result": result_text[:600]}

        if denied:
            # 사용자가 이 단계의 위험 도구를 거절 — '실패'가 아니라 '하지 않기로 한 것'
            # 이다. 재계획으로 같은 동작을 다시 만들지 않고, 거절 사실을 결과에 남겨
            # 종합이 수행된 것처럼 서술하지 못하게 한다.
            state.results.append(
                f"[단계 {i + 1} — 사용자 거절] 사용자가 위험 도구 실행을 거절해 "
                f"이 단계는 수행되지 않았다. 모델 보고: {result_text}")
            i += 1
            continue

        if ok:
            state.results.append(f"[단계 {i + 1}] {result_text}")
            i += 1
            continue

        # 실패 → 예산이 있으면 남은 계획을 다시 세우고 같은 자리에서 재시도한다.
        if state.replans < max_replans and state.executed < max_steps:
            state.replans += 1
            try:
                replan_text = await _complete(
                    client, model, _REPLAN_SYSTEM,
                    f"목표: {state.goal}\n\n원래 계획:\n{plan_str}\n\n"
                    f"지금까지의 결과:\n{prior}\n\n"
                    f"실패한 단계: {step}\n실패 결과: {result_text}{server_menu}",
                    send_top_k=send_top_k)
                new_raw = _parse_steps(replan_text)
            except Exception:  # noqa: BLE001
                new_raw = []
            if new_raw:
                # 재계획 단계도 [서버] 태그를 떼어 plan/servers를 나란히 스플라이스한다.
                keep = max_steps - i
                new_bodies, new_scopes = [], []
                for rs in new_raw[:keep]:
                    b, sc = _split_tag(rs, valid_servers)
                    new_bodies.append(b)
                    new_scopes.append(sc)
                state.plan = state.plan[:i] + new_bodies
                state.servers = state.servers[:i] + new_scopes
                yield {"type": "plan", "steps": list(state.plan), "replan": state.replans}
                continue  # i 그대로 → 새 단계 재시도
        # 재계획 못 하거나 예산 소진 → 실패를 기록하고 다음 단계로 넘어간다.
        state.results.append(f"[단계 {i + 1} 실패] {result_text}")
        i += 1

    if state.executed >= max_steps and i < len(state.plan):
        yield {"type": "step_token", "index": -1,
               "text": f"\n[알림] 스텝 상한({max_steps})에 도달해 남은 단계를 건너뛰고 종합합니다.\n"}

    # ---------- 종합 ----------
    results_str = "\n".join(state.results) or "(수행된 단계 없음)"
    synth_user = (
        f"{hist_block}목표: {state.goal}\n\n단계별 결과:\n{results_str}\n\n"
        f"위 결과와 이전 대화를 함께 고려해 최종 답변을 작성하라."
    )
    final_parts: list[str] = []
    async for ev in agent.run_chat(
        base_url=base_url, model=model,
        messages=[{"role": "system", "content": _SYNTH_SYSTEM},
                  {"role": "user", "content": synth_user}],
        settings=settings, api_key=api_key, send_top_k=send_top_k,
        mcp=None, memory=None, mock=False,  # 종합은 도구 없이 답변만
    ):
        if ev.get("type") == "token":
            final_parts.append(ev["text"])
            yield ev
        elif ev.get("type") == "error":
            yield ev

    final_text = "".join(final_parts).strip() or "(종합 결과가 비어 있습니다.)"
    yield {"type": "done",
           "messages": list(messages) + [{"role": "assistant", "content": final_text}]}
