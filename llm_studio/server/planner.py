"""planner.py — 계획-실행(plan-and-execute) 에이전트.

약한 로컬 모델(Gemma 12B 등)은 긴 도구 루프에서 "지금 뭘 하는 중인지"를 쉽게 잃는다.
그래서 요청을 먼저 **계획**으로 외재화하고, 하네스가 그 계획을 붙든 채 스텝마다
좁은 컨텍스트만 주며 하나씩 **실행**한 뒤 결과를 **종합**한다. 판단을 모델의 즉흥적
추론이 아니라 하네스의 제어 흐름이 쥐는 구조다.

run_task()는 run_chat()과 같은 방식으로 이벤트 dict를 내보내는 async generator다:
    {"type": "plan",       "steps": [...], "replan": n?, "edited": bool?}   계획 수립
    {"type": "step_start", "index": i, "text": "..."}      한 스텝 실행 시작
    {"type": "step_token", "index": i, "text": "..."}      스텝 진행 중 출력 조각(도구 포함)
    {"type": "step_done",  "index": i, "ok": bool, "result": "..."}
    {"type": "branch_start","index": i, "cond": "...", "mode": "llm"|"rule"}  분기 판정 시작
    {"type": "branch",     "index": i, "cond": "...", "result": bool, "mode": ..., "target": M, "skipped": [...]}
    {"type": "steer_request", "phase": "plan"|"step_failed", ...}  조종 게이트
    {"type": "token",      "text": "..."}                  최종 종합 답변 조각
    {"type": "done",       "messages": [...]}              완료 (대화에 저장할 메시지)
    {"type": "error",      "message": "..."}

각 스텝 실행은 기존 agent.run_chat()을 그대로 재사용한다(도구 루프 포함). 스텝 안에서
나오는 토큰·도구 이벤트는 전부 step_token으로 감싸 해당 스텝 블록에만 흐르게 하고,
최종 답변(종합)만 일반 token으로 흘려 대화의 답변 영역에 렌더링되게 한다.

단기(작업) 상태는 RAM에만 둔다 — 이 요청이 사는 동안만 유효하며 디스크에 남기지 않는다.
재시작·크래시 시 진행 중 작업은 사라진다(대화형 전제). 무한 루프 방지를 위해
max_steps(총 실행 상한)와 max_replans(재계획 예산)를 둔다.

## 실행 코크핏 확장 (Layer 2·3)

- **아티팩트 캡처** — 하네스가 각 스텝의 도구 원출력+최종텍스트를 `TaskState.artifacts`에
  스텝 번호로 담는다. 모델의 요약 문구가 아니라 하네스가 결정적으로 붙잡으므로 raw
  데이터가 소실되지 않는다.
- **의존 태그 `[←N]`** — 계획가가 각 단계에 붙인 `[←2,3]`으로 그 단계가 어느 앞선 단계의
  결과를 입력으로 쓰는지 표시한다. 실행 때 지정한 스텝의 아티팩트 원본만 주입하고,
  역인덱스로 "이 결과가 어느 뒷 단계에서 쓰이는지(소비처)"도 계약(system)으로 알려
  양방향 인지를 만든다. 태그가 없으면 "앞 결과 전부 요약 주입"으로 우아하게 저하한다.
- **조건 분기 `[?→M]`** — 조건이 거짓이면 단계 M까지 건너뛴다(전방 전용, 루프 없음).
  판정은 규칙(하네스 결정적) 또는 LLM으로 하며(`[?규칙→M]`이면 규칙 우선, 판정 불가 시
  LLM으로 저하), 대상이 유효하지 않으면 분기를 무시하고 그대로 진행한다.

원칙: **시스템 프롬프트 = 계약, 유저 메시지 = 데이터.** 의존 지도(짧은 계약)는 system에,
아티팩트 원본(무거운 데이터)은 user 메시지에 둔다 — 약한 모델은 system 지시는 잘
따르지만 데이터를 system에 부으면 앞부분 지시를 흘려버리기 때문이다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import AsyncIterator

from openai import AsyncOpenAI

from . import agent

DEFAULT_MAX_STEPS = 10
DEFAULT_MAX_REPLANS = 2
DEFAULT_STEER_TIMEOUT = 600  # 조종 게이트 대기 상한(초). 0 이하 = 무제한. 초과 시 그대로 진행.
_ARTIFACT_CHARS = 2000       # 아티팩트로 저장할 도구 원출력 하나당 상한(RAM·ctx 보호)
_DEP_INJECT_CHARS = 1500     # 의존 주입 시 아티팩트 한 조각을 자르는 한도

# 의도 라우터: 이번 요청을 '단순 대화(run_chat)'로 처리할지 '다단계 작업(run_task)'으로
# 처리할지 두 갈래로 분류한다. 약한 모델도 맞히도록 2지선다·한 단어 출력으로 조인다.
# 오분류해도 계획 확정 게이트에서 사람이 회복할 수 있어 정확도에 목매지 않는다.
_ROUTE_SYSTEM = (
    "너는 사용자의 요청을 두 갈래로 분류하는 라우터다. "
    "여러 단계·여러 도구·순차적 수행이 필요한 복합 작업이면 'TASK', "
    "한 번의 답변으로 되는 단순 질문·잡담·설명·단일 조회면 'CHAT'으로 분류한다. "
    "다른 말 없이 정확히 'TASK' 또는 'CHAT' 한 단어만 출력한다."
)

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
_BRANCH_SYSTEM = (
    "너는 조건이 참인지 거짓인지 판정한다. 지금까지의 작업 결과가 주어진다. "
    "주어진 조건이 그 결과에 비추어 성립하면 '예', 성립하지 않으면 '아니오'만 "
    "한 단어로 답한다. 다른 말은 하지 않는다."
)

# 계획가에게 의존 태그(Layer 2)와 분기 태그(Layer 3)를 가르치는 지시. 서버 메뉴 뒤에
# 붙여 계획 요청 프롬프트에 넣는다. 약한 모델이 태그를 안 붙여도 우아하게 저하하므로
# (의존 없음 → 앞 결과 전부, 분기 없음 → 일반 스텝) 정확도에 목매지 않는다.
_DEP_INSTRUCT = (
    "\n\n어떤 단계가 앞선 특정 단계의 결과를 입력으로 쓰면 번호 뒤에 [←N] 또는 "
    "[←N,M]으로 표시한다(N은 참조할 단계 번호). 예: '3. [←1,2] 두 결과를 비교한다'. "
    "서버 태그와 함께 쓰면 '[office ←1] ...'처럼 한 대괄호에 담아도 된다. "
    "표시가 없으면 앞의 모든 결과가 주어진다."
)
_BRANCH_INSTRUCT = (
    "\n\n어떤 단계가 조건에 따라 이후 단계를 건너뛰는 분기라면 번호 뒤에 [?→M]으로 "
    "표시하고(M은 조건이 거짓일 때 건너뛰고 이어갈 단계 번호), 그 단계 본문에는 "
    "판정할 조건을 적는다. 예: '3. [?→6] 검색 결과가 존재한다'. 결정적 규칙으로 "
    "판정하려면 '[?규칙→M]'으로 표시한다(불가하면 자동으로 LLM 판정으로 저하). "
    "M은 반드시 그 단계보다 뒤여야 한다."
)

_STEP_RE = re.compile(r"^\s*\d+[.)]\s*(.+)$")


@dataclass
class Step:
    """계획의 한 단계. 본문과 함께 도구 스코프·의존·분기 메타를 담는다.

    - scope: 서버 스코프. None=전체 툴 / []=도구 없음 / [이름…]=그 서버 도구만.
    - deps:  의존 스텝(0-based). None=태그 없음(앞 결과 전부) / []=의존 없음 명시 /
             [i…]=그 스텝들의 아티팩트만 원본 주입.
    - branch: 분기 메타 dict{"mode":"llm"|"rule", "target": 0-based} 또는 None(일반 스텝).
             본문(body)이 판정할 조건이다.
    """
    body: str
    scope: "list[str] | None" = None
    deps: "list[int] | None" = None
    branch: "dict | None" = None


@dataclass
class TaskState:
    """단기(작업) 메모리 — RAM에만 존재한다. 계획과 진행 상황을 담는다."""
    goal: str
    plan: list  # list[Step]
    results: list[str] = field(default_factory=list)  # 완료된 스텝의 결과 요약(종합·기본 컨텍스트용)
    # 스텝 번호(0-based) → {"step": 본문, "text": 최종텍스트, "tools": [원출력…]}.
    # 하네스가 결정적으로 캡처한 raw 아티팩트. 의존 태그가 지정한 스텝을 여기서 원본으로 끌어온다.
    artifacts: dict = field(default_factory=dict)
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


async def classify_intent(
    *, base_url: str, model: str, messages: list[dict],
    api_key: str = "local", send_top_k: bool = True,
) -> str:
    """이번 요청을 'task'(다단계 작업) 또는 'chat'(단순 대화)으로 분류한다.

    코크핏의 0번 노드(의도 판단). 단발·저온·비스트리밍 1회 호출이며, 프롬프트에는
    전체 이력이 아니라 마지막 사용자 발화만 넣어 프리필을 작게 유지한다(약한 로컬
    모델에서도 빠르게). 빈 입력·형식 오류·네트워크 오류 등 어떤 실패든 'chat'으로
    물러선다 — 값비싼 계획-실행을 잘못 태우는 것보다 손해가 작다(오분류는 계획
    게이트에서 사람이 회복 가능).
    """
    goal = _last_user(messages)
    if not goal.strip():
        return "chat"
    client = AsyncOpenAI(base_url=base_url, api_key=api_key or "none", timeout=120)
    try:
        out = await _complete(client, model, _ROUTE_SYSTEM, f"요청: {goal}",
                              send_top_k=send_top_k, max_tokens=8)
    except Exception:  # noqa: BLE001 — 분류 실패는 조용히 chat으로 저하
        return "chat"
    return "task" if "TASK" in out.upper() else "chat"


async def _steer(steerer, request: dict, timeout: float):
    """steer_request를 흘리고 사용자의 결정을 기다리는 async generator.

    요청 이벤트를 yield한 뒤, 마지막에 {"__steer__": <결정 dict 또는 None>}를
    yield한다(_run_step의 __result__ 패턴과 같은 방식 — 호출부가 이벤트는 그대로
    흘리고 결정만 뽑아낸다). None이면 시간 초과/무응답이므로 호출부는 기본 동작으로
    진행한다. yield 지점에서 스트림이 끊겨도 finally가 대기 항목을 정리한다.
    """
    req_id = steerer.create()
    request = dict(request)
    request["id"] = req_id
    request["timeout"] = timeout
    try:
        yield request
        decision = await steerer.wait(req_id, timeout)
    finally:
        steerer.discard(req_id)
    yield {"__steer__": decision}


_LEAD_TAG_RE = re.compile(r"^\[([^\]]*)\]\s*")
_ARROW_L_RE = re.compile(r"←|<-|<=")               # 의존(왼쪽 화살표)
_BRANCH_RE = re.compile(r"^\?\s*(.*?)\s*(?:→|->|=>)\s*(\d+)\s*$")  # 분기 '?[모드]→M'


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


def _parse_dep_indices(raw: str) -> list[int]:
    """'1,2 / 3' 같은 문자열에서 스텝 번호를 뽑아 0-based 인덱스로 만든다."""
    out = []
    for tok in re.split(r"[,\s/]+", raw.strip()):
        tok = tok.strip().lstrip("#").strip()
        if tok.isdigit():
            n = int(tok)
            if n >= 1 and (n - 1) not in out:
                out.append(n - 1)
    return out


def _parse_branch(raw: str):
    """'?[모드]→M' 대괄호 내용을 분기 메타로 파싱한다. 형식이 아니면 None.

    반환: {"mode": "llm"|"rule", "target": 0-based} — target은 조건이 거짓일 때
    건너뛰고 이어갈 단계(0-based). 대상 유효성(전방·범위)은 실행 시점에 검사한다.
    """
    m = _BRANCH_RE.match(raw.strip())
    if not m:
        return None
    modeword = (m.group(1) or "").strip().lower()
    target = int(m.group(2)) - 1
    mode = "rule" if ("규칙" in modeword or "rule" in modeword) else "llm"
    return {"mode": mode, "target": target}


def _split_tags(step_text: str, valid: list[str], *, allow_branch: bool = True) -> Step:
    """단계 앞머리의 [서버]·[←의존]·[?분기] 태그들을 떼어 Step으로 만든다.

    앞머리에 붙은 여러 대괄호를 순서대로 벗겨 각각을 서버/의존/분기로 분류한다.
    한 대괄호에 서버와 의존을 '←'로 함께 담아도 된다('[office ←1]'). 알 수 없는
    대괄호를 만나면 거기서 멈춰 본문의 일부로 남긴다(약한 모델의 형식 실수에
    우아하게 저하 — 태그가 없으면 전체 툴·앞 결과 전부·일반 스텝으로 동작).
    """
    text = step_text.strip()
    scope = None
    deps = None
    branch = None
    while True:
        m = _LEAD_TAG_RE.match(text)
        if not m:
            break
        raw = m.group(1).strip()
        rest = text[m.end():]
        if raw.startswith("?"):
            if not allow_branch:
                break  # 분기 비활성 → 본문으로 남긴다(일반 스텝으로 실행)
            b = _parse_branch(raw)
            if b is None:
                break  # 분기 형식이 아님 → 본문으로 남긴다
            branch = b
            text = rest
            continue
        # 서버/의존 태그 (한 대괄호에 '←'로 섞일 수 있음)
        parts = _ARROW_L_RE.split(raw, maxsplit=1)
        srv_raw = parts[0].strip()
        if len(parts) > 1:
            deps = _parse_dep_indices(parts[1])
        if srv_raw in ("", "-", "없음", "none", "None"):
            if len(parts) == 1:
                scope = []  # [-] → 도구 없음. (순수 [←N]은 서버 미지정 → None 유지)
        else:
            picks = [p.strip() for p in re.split(r"[,\s/]+", srv_raw) if p.strip()]
            valid_picks = [p for p in picks if p in valid]
            if valid_picks:
                scope = valid_picks
            # 무효 서버명뿐이면 scope는 None(전체 툴)으로 둔다
        text = rest
    body = text.strip() or step_text.strip()
    return Step(body=body, scope=scope, deps=deps, branch=branch)


def _step_display(s: Step) -> str:
    """Step을 태그 포함 한 줄 텍스트로 되살린다(계획 이벤트·편집 라운드트립용).

    사람이 계획 게이트에서 편집할 때 태그가 보이고, 편집 결과를 다시 _split_tags로
    파싱하면 같은 구조로 복원된다. 대괄호를 나눠 적어도 파서가 다시 합친다.
    """
    tags = []
    if s.branch is not None:
        mode = "규칙" if s.branch.get("mode") == "rule" else ""
        tags.append(f"?{mode}→{s.branch['target'] + 1}")
    if s.scope is not None:
        tags.append("-" if not s.scope else " ".join(s.scope))
    if s.deps:
        tags.append("←" + ",".join(str(d + 1) for d in s.deps))
    prefix = "".join(f"[{t}] " for t in tags)
    return prefix + s.body


def _step_suffix(s: Step) -> str:
    """스텝 표시에 붙일 도구 스코프·의존 라벨(투명성용). 없으면 빈 문자열."""
    bits = []
    if s.scope is not None:
        bits.append("도구 없음" if not s.scope else ", ".join(s.scope))
    if s.deps:
        bits.append("←" + ",".join(str(d + 1) for d in s.deps))
    return f"  ⟨{' · '.join(bits)}⟩" if bits else ""


def _plan_meta(plan: list) -> list[dict]:
    """계획 이벤트에 실을 스텝별 메타(의존·소비처·분기). 프론트 시각화용."""
    meta = []
    for i, s in enumerate(plan):
        consumers = [j + 1 for j, t in enumerate(plan)
                     if t.deps and i in t.deps]
        m = {}
        if s.deps:
            m["deps"] = [d + 1 for d in s.deps]
        if consumers:
            m["consumers"] = consumers
        if s.scope is not None:
            m["scope"] = list(s.scope)
        if s.branch is not None:
            m["branch"] = {"mode": s.branch["mode"], "target": s.branch["target"] + 1}
        meta.append(m)
    return meta


def _plan_event(plan: list, **extra) -> dict:
    """plan 이벤트를 만든다 — steps(태그 포함 표시)와 meta(시각화용)를 함께 싣는다."""
    ev = {"type": "plan", "steps": [_step_display(s) for s in plan],
          "meta": _plan_meta(plan)}
    ev.update(extra)
    return ev


def _dep_context(state: TaskState, deps) -> str:
    """의존 태그에 따라 스텝 프롬프트의 '지금까지의 결과' 블록을 만든다.

    - deps is None(태그 없음): 기존 동작 — 완료된 모든 스텝의 결과 요약을 넣는다.
    - deps == []: 의존 없음 명시 — 아무 결과도 넣지 않는다(독립 스텝).
    - deps == [i…]: 지정한 스텝의 아티팩트 원본(최종텍스트+도구 원출력)만 넣는다.
      => '값 싼 전체 요약'이 아니라 '필요한 원본'만 주입해 raw 데이터를 살린다.
    """
    if deps is None:
        return "\n".join(f"- {r}" for r in state.results) or "(없음)"
    if not deps:
        return "(이 단계는 앞선 결과에 의존하지 않습니다)"
    blocks = []
    for d in deps:
        art = state.artifacts.get(d)
        if not art:
            continue
        head = f"[단계 {d + 1} 결과] {art.get('step', '')}".strip()
        body = (art.get("text") or "").strip()
        piece = f"{head}\n{body[:_DEP_INJECT_CHARS]}" if body else head
        tools = art.get("tools") or []
        for k, t in enumerate(tools):
            piece += f"\n  · 도구 원출력{k + 1}: {str(t)[:_DEP_INJECT_CHARS]}"
        blocks.append(piece)
    return "\n\n".join(blocks) or "(지정한 의존 단계의 결과가 아직 없습니다)"


def _dep_map_line(deps, consumers) -> str:
    """스텝 system 프롬프트에 넣을 짧은 의존 지도(계약). 없으면 빈 문자열."""
    parts = []
    if deps:
        parts.append("이 단계의 입력 = 단계 " + ", ".join(str(d + 1) for d in deps) + "의 결과")
    if consumers:
        parts.append("이 단계의 결과는 단계 " + ", ".join(str(c + 1) for c in consumers) + "에서 쓰인다")
    return " / ".join(parts)


def _prior_text(state: TaskState) -> str:
    """분기 판정에 줄 '지금까지의 결과' 원본 텍스트(요약+아티팩트+도구출력)."""
    parts = list(state.results)
    for i in sorted(state.artifacts):
        art = state.artifacts[i]
        if art.get("text"):
            parts.append(art["text"])
        parts.extend(str(t) for t in (art.get("tools") or []))
    return "\n".join(parts)


def _eval_rule(cond: str, state: TaskState):
    """조건을 하네스가 결정적으로 판정한다. 판정 불가하면 None(→ LLM으로 저하).

    지원하는 규칙(약한 모델이 만들 만한 것 위주, 최소):
      - 포함 검사: 따옴표로 감싼 문자열 + '포함/있'  → 앞 결과에 그 문자열이 있는가
      - 실패 여부: '실패'                              → 앞 결과에 실패 표시가 있는가
      - 성공/완료: '성공' 또는 '완료'                  → 앞 결과에 실패 표시가 없는가
    그 외는 None을 돌려주어 LLM 판정으로 넘긴다.
    """
    c = cond.strip()
    text = _prior_text(state)
    m = re.search(r'["\'“”]([^"\'“”]+)["\'“”]', c)
    if m and ("포함" in c or "있" in c or "contain" in c.lower()):
        return m.group(1) in text
    if "실패" in c:
        return "[실패]" in text or "실패]" in text
    if "성공" in c or "완료" in c:
        return "[실패]" not in text
    return None


async def _eval_llm(client: AsyncOpenAI, model: str, cond: str, state: TaskState,
                    *, send_top_k: bool) -> bool:
    """조건을 LLM에 물어 참/거짓으로 판정한다. 실패·애매하면 True(건너뛰지 않음).

    거짓 판정은 뒷 단계를 건너뛰는(작업을 덜 하는) 방향이라, 판정이 불확실할 때는
    '참(그대로 진행)'으로 물러서는 편이 안전하다 — 실수로 필요한 단계를 날리지 않는다.
    """
    ctx = _prior_text(state)[:3000] or "(아직 결과 없음)"
    try:
        out = await _complete(
            client, model, _BRANCH_SYSTEM,
            f"지금까지의 결과:\n{ctx}\n\n조건: {cond}\n\n이 조건은 참인가?",
            send_top_k=send_top_k, max_tokens=8)
    except Exception:  # noqa: BLE001
        return True
    o = out.strip().lower()
    if any(w in o for w in ("아니", "no", "false", "거짓")):
        return False
    if any(w in o for w in ("예", "yes", "true", "참", "네")):
        return True
    return True  # 애매 → 진행


async def _eval_branch(client, model, branch: dict, cond: str, state: TaskState,
                       *, send_top_k: bool):
    """분기 조건을 판정한다. 반환: (참/거짓, 실제 사용한 판정법 'rule'|'llm')."""
    if branch.get("mode") == "rule":
        val = _eval_rule(cond, state)
        if val is not None:
            return bool(val), "rule"
        # 규칙으로 판정 불가 → LLM으로 우아하게 저하
    return await _eval_llm(client, model, cond, state, send_top_k=send_top_k), "llm"


async def _run_step(step_idx: int, prompt_messages: list[dict], *, base_url, model,
                    settings, api_key, send_top_k, mcp, memory,
                    tool_servers=None, approver=None) -> AsyncIterator[dict]:
    """한 스텝을 agent.run_chat으로 실행하며 진행 이벤트를 흘린다.

    run_chat의 토큰·도구 이벤트를 전부 step_token으로 감싼다(해당 스텝 블록에만 표시).
    도구 원출력은 __tools__에 모아(아티팩트 캡처) 마지막 __result__ 이벤트로 함께
    돌려준다 — 모델의 요약이 아니라 하네스가 결정적으로 붙잡은 raw 데이터다.
    마지막에 {"__result__": text, "__ok__": bool, "__denied__": bool, "__tools__": [...]}를
    담은 이벤트를 하나 내보내 호출부가 스텝 결과/성공 여부/아티팩트를 받게 한다.

    tool_servers: 이 스텝에 노출할 MCP 서버 스코프(서버-스코프 라우팅). None이면
    전체, 빈 목록이면 도구 없음. run_chat에 그대로 넘긴다.
    """
    final_text = ""
    ok = True
    denied = False
    tool_outputs: list[str] = []
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
            raw = str(ev.get("result", ""))
            # 아티팩트로 원출력을 붙잡는다(상한으로 잘라 RAM·ctx 보호). 표시는 더 짧게.
            tool_outputs.append(raw[:_ARTIFACT_CHARS])
            yield {"type": "step_token", "index": step_idx,
                   "text": f"↳ 결과: {raw[:400]}\n"}
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
    yield {"__result__": final_text.strip(), "__ok__": ok, "__denied__": denied,
           "__tools__": tool_outputs}


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
    steerer=None,
    failure_gate: bool = False,
    branch_enabled: bool = True,
    steer_timeout: float = DEFAULT_STEER_TIMEOUT,
) -> AsyncIterator[dict]:
    """계획-실행으로 한 턴을 처리한다. messages는 system 포함 전체 이력.

    계획을 못 세우면(파싱 0개/오류) 일반 run_chat으로 우아하게 저하한다.

    steerer/failure_gate: '실행 코크핏'의 조종 게이트. steerer(SteeringBroker)가 주어지고
    failure_gate가 켜지면 스텝 실패 시 재시도/건너뛰기/재계획/편집/중단을 묻는다.
    steerer가 없거나 게이트가 꺼져 있으면 기존과 똑같이 자동으로 흐른다(우아한 저하).
    (계획 확정 게이트는 제거됨 — 계획을 세우면 항상 바로 실행한다.)

    branch_enabled: 조건 분기('[?→M]') 태그를 해석할지 여부. 끄면 분기 태그를 본문의
    일부로 두어 일반 스텝처럼 실행한다.
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
    # 계획가에게 의존/분기 태그를 가르친다(분기는 켜져 있을 때만).
    plan_extra = server_menu + _DEP_INSTRUCT + (_BRANCH_INSTRUCT if branch_enabled else "")

    # ---------- 계획 ----------
    try:
        plan_text = await _complete(
            client, model, _PLAN_SYSTEM,
            f"{hist_block}요청: {goal}{plan_extra}", send_top_k=send_top_k)
        raw_steps = _parse_steps(plan_text)[:max_steps]
    except Exception as e:  # noqa: BLE001
        raw_steps = []
        yield {"type": "step_token", "index": -1, "text": f"[계획 실패: {e}]"}

    # 각 단계에서 [서버]/[←의존]/[?분기] 태그를 떼어 Step으로 만든다.
    plan = [_split_tags(rs, valid_servers, allow_branch=branch_enabled)
            for rs in raw_steps]

    if not plan:
        # 계획을 못 세움 → 일반 채팅으로 저하 (도구 포함). 이벤트를 그대로 흘린다.
        async for ev in agent.run_chat(
            base_url=base_url, model=model, messages=messages, settings=settings,
            api_key=api_key, send_top_k=send_top_k, mcp=mcp, memory=memory, mock=False,
            approver=approver,
        ):
            yield ev
        return

    state = TaskState(goal=goal, plan=plan)
    yield _plan_event(state.plan)

    # ---------- 실행 루프 ----------
    i = 0
    while i < len(state.plan) and state.executed < max_steps:
        step = state.plan[i]

        # ── 조건 분기 노드 ──
        # 실행이 아니라 판정이다. 조건이 거짓이면 대상 단계까지 건너뛴다(전방 전용).
        if branch_enabled and step.branch is not None:
            state.executed += 1
            cond = step.body
            br = step.branch
            target = br.get("target", -1)
            yield {"type": "branch_start", "index": i, "cond": cond, "mode": br["mode"]}
            valid_target = i < target <= len(state.plan)
            if not valid_target:
                # 대상이 무효(뒤가 아니거나 범위 밖) → 분기를 무시하고 그대로 진행.
                # 루프·역행을 만들지 않기 위한 안전장치.
                yield {"type": "branch", "index": i, "cond": cond, "result": True,
                       "mode": br["mode"], "skipped": [],
                       "note": "분기 대상이 유효하지 않아 그대로 진행"}
                state.results.append(f"[단계 {i + 1} 분기] 대상이 유효하지 않아 계속 진행")
                state.artifacts[i] = {"step": cond, "text": "(분기: 계속 진행)", "tools": []}
                i += 1
                continue
            result_bool, how = await _eval_branch(
                client, model, br, cond, state, send_top_k=send_top_k)
            if result_bool:
                yield {"type": "branch", "index": i, "cond": cond, "result": True,
                       "mode": how, "target": target + 1, "skipped": []}
                state.results.append(f"[단계 {i + 1} 분기] 조건 '{cond}' → 참, 계속 진행")
                state.artifacts[i] = {"step": cond, "text": "(분기: 참 → 계속)", "tools": []}
                i += 1
            else:
                skipped = list(range(i + 1, target))
                yield {"type": "branch", "index": i, "cond": cond, "result": False,
                       "mode": how, "target": target + 1,
                       "skipped": [s + 1 for s in skipped]}
                span = (f"단계 {i + 2}~{target}" if skipped else "건너뛸 단계 없음")
                state.results.append(
                    f"[단계 {i + 1} 분기] 조건 '{cond}' → 거짓, {span} 건너뜀")
                state.artifacts[i] = {"step": cond, "text": "(분기: 거짓 → 건너뜀)", "tools": []}
                for s in skipped:
                    state.artifacts[s] = {"step": state.plan[s].body,
                                          "text": "(분기로 건너뜀)", "tools": []}
                i = target
            continue

        # ── 일반 실행 스텝 ──
        scope = step.scope
        state.executed += 1
        yield {"type": "step_start", "index": i, "text": step.body + _step_suffix(step)}

        # 의존 태그에 따라 컨텍스트를 좁힌다(원본 주입) 또는 앞 결과 전부(요약).
        prior = _dep_context(state, step.deps)
        # 역인덱스: 이 스텝의 결과를 뒤에서 소비하는 단계(→N). 양방향 인지용(계약).
        consumers = [j + 1 for j, s in enumerate(state.plan)
                     if s.deps and i in s.deps]
        dep_line = _dep_map_line(step.deps, consumers)
        step_system = _STEP_SYSTEM + (f"\n\n[의존 관계] {dep_line}" if dep_line else "")
        plan_str = "\n".join(f"{n + 1}. {_step_display(s)}"
                             for n, s in enumerate(state.plan))
        step_user = (
            f"{hist_block_step}전체 목표: {state.goal}\n\n계획:\n{plan_str}\n\n"
            f"지금까지의 결과:\n{prior}\n\n"
            f"지금 수행할 단계 ({i + 1}/{len(state.plan)}): {step.body}"
        )
        result_text, ok, denied, tools = "", True, False, []
        async for ev in _run_step(
            i, [{"role": "system", "content": step_system},
                {"role": "user", "content": step_user}],
            base_url=base_url, model=model, settings=settings, api_key=api_key,
            send_top_k=send_top_k, mcp=mcp, memory=memory, tool_servers=scope,
            approver=approver,
        ):
            if "__result__" in ev:
                result_text, ok = ev["__result__"], ev["__ok__"]
                denied = ev.get("__denied__", False)
                tools = ev.get("__tools__", [])
            else:
                yield ev

        yield {"type": "step_done", "index": i, "ok": ok and not denied,
               "result": result_text[:600]}

        if denied:
            # 사용자가 이 단계의 위험 도구를 거절 — '실패'가 아니라 '하지 않기로 한 것'
            # 이다. 재계획으로 같은 동작을 다시 만들지 않고, 거절 사실을 결과에 남겨
            # 종합이 수행된 것처럼 서술하지 못하게 한다.
            note = (f"[단계 {i + 1} — 사용자 거절] 사용자가 위험 도구 실행을 거절해 "
                    f"이 단계는 수행되지 않았다. 모델 보고: {result_text}")
            state.results.append(note)
            state.artifacts[i] = {"step": step.body, "text": note, "tools": tools}
            i += 1
            continue

        if ok:
            state.results.append(f"[단계 {i + 1}] {result_text}")
            # 아티팩트: 하네스가 결정적으로 붙잡은 최종텍스트+도구 원출력. 의존 태그가
            # 이 스텝을 지목하면 이 원본이 뒷 단계 user 메시지로 그대로 흘러간다.
            state.artifacts[i] = {"step": step.body, "text": result_text, "tools": tools}
            i += 1
            continue

        # ---------- 실패 게이트 (반자동) ----------
        # 스텝이 실패하면 멈춰 사람에게 재시도/건너뛰기/재계획/편집/중단을 묻는다.
        # 게이트가 꺼져 있거나 무응답(None)이면 아래 자동 재계획으로 흐른다(기존 동작).
        if failure_gate and steerer is not None:
            fdecision = None
            async for ev in _steer(
                    steerer, {"type": "steer_request", "phase": "step_failed",
                              "index": i, "step": step.body,
                              "result": result_text[:600]}, steer_timeout):
                if "__steer__" in ev:
                    fdecision = ev["__steer__"]
                else:
                    yield ev
            if fdecision:
                action = fdecision.get("action")
                if action == "abort":
                    note = (f"[단계 {i + 1} — 중단] 사용자가 이 지점에서 작업을 중단했다. "
                            f"모델 보고: {result_text}")
                    state.results.append(note)
                    state.artifacts[i] = {"step": step.body, "text": note, "tools": tools}
                    break  # 실행 루프 종료 → 지금까지의 결과로 종합
                if action == "skip":
                    note = (f"[단계 {i + 1} 건너뜀] 사용자가 이 단계를 건너뛰었다. "
                            f"모델 보고: {result_text}")
                    state.results.append(note)
                    state.artifacts[i] = {"step": step.body, "text": note, "tools": tools}
                    i += 1
                    continue
                if action == "edit":
                    new_text = (fdecision.get("step") or "").strip()
                    if new_text:
                        s = _split_tags(new_text, valid_servers, allow_branch=branch_enabled)
                        if s.body:
                            state.plan[i] = s
                        yield _plan_event(state.plan, edited=True)
                    continue  # i 그대로 → 편집한 단계 재시도
                if action == "retry":
                    continue  # i 그대로 → 같은 단계 재시도
                # action == "replan" → 아래 자동 재계획으로 진행

        # 실패 → 예산이 있으면 남은 계획을 다시 세우고 같은 자리에서 재시도한다.
        if state.replans < max_replans and state.executed < max_steps:
            state.replans += 1
            plan_str = "\n".join(f"{n + 1}. {_step_display(s)}"
                                 for n, s in enumerate(state.plan))
            prior_all = "\n".join(f"- {r}" for r in state.results) or "(없음)"
            try:
                replan_text = await _complete(
                    client, model, _REPLAN_SYSTEM,
                    f"목표: {state.goal}\n\n원래 계획:\n{plan_str}\n\n"
                    f"지금까지의 결과:\n{prior_all}\n\n"
                    f"실패한 단계: {step.body}\n실패 결과: {result_text}{plan_extra}",
                    send_top_k=send_top_k)
                new_raw = _parse_steps(replan_text)
            except Exception:  # noqa: BLE001
                new_raw = []
            if new_raw:
                # 재계획 단계도 태그를 떼어 Step으로 만들고 tail을 스플라이스한다.
                keep = max_steps - i
                new_steps = [_split_tags(rs, valid_servers, allow_branch=branch_enabled)
                             for rs in new_raw[:keep]]
                state.plan = state.plan[:i] + new_steps
                yield _plan_event(state.plan, replan=state.replans)
                continue  # i 그대로 → 새 단계 재시도
        # 재계획 못 하거나 예산 소진 → 실패를 기록하고 다음 단계로 넘어간다.
        note = f"[단계 {i + 1} 실패] {result_text}"
        state.results.append(note)
        state.artifacts[i] = {"step": step.body, "text": note, "tools": tools}
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
