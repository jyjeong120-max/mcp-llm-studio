"""steering.py — 작업 조종(steer) 브로커.

계획-실행(planner) '실행 코크핏'에서 사람이 흐름에 끼어드는 통로다. 위험 도구
승인(approvals.py)이 "예/아니오"만 받는다면, 조종은 **구조화된 결정**을 받는다:
계획을 확정할지 편집할지, 실패한 스텝을 재시도할지 건너뛸지 등.

approvals.ApprovalBroker와 같은 Future-대기 패턴을 일반화한 것이다:
    1. planner가 개입 지점(계획 확정·스텝 실패)에서 create()로 요청 id를 만들고
       SSE로 steer_request 이벤트를 흘린다
    2. wait()로 사용자의 결정을 기다린다 (브라우저가 조종 카드를 띄움)
    3. 브라우저가 POST /api/chat/steer 로 resolve()를 부르면 wait가 풀린다

승인과의 차이 두 가지:
- resolve가 bool이 아니라 **결정 dict**({"action": ..., "steps": [...], ...})를 싣는다.
- 시간 초과/무응답은 '거절'이 아니라 **None**을 돌려준다 — 호출부(planner)가 이를
  "그대로 진행"으로 우아하게 저하한다(무한 대기 금지, 관찰만 하다 놔둬도 흐름이 이어짐).

상태는 전부 RAM(asyncio Future)이다 — 대화 저장과 무관하고, 스트림이 끊기면
(사용자 중단) 대기도 함께 취소된다.
"""

from __future__ import annotations

import asyncio
import secrets


class SteeringBroker:
    """조종 대기 중인 요청들을 id → Future로 들고 있는 단순 브로커."""

    def __init__(self):
        self._pending: dict[str, asyncio.Future] = {}

    def create(self) -> str:
        """새 조종 요청을 등록하고 요청 id를 돌려준다.

        id는 추측 불가능한 토큰이어야 한다 — 순차 번호를 쓰면 /api/chat/steer를
        원격에서 열거해 남의 작업 흐름을 대신 조작할 수 있다(승인 브로커와 같은 원칙).
        """
        req_id = f"steer_{secrets.token_urlsafe(16)}"
        self._pending[req_id] = asyncio.get_running_loop().create_future()
        return req_id

    def discard(self, req_id: str) -> None:
        """대기 항목을 제거한다(멱등). create() 후 wait()에 도달하지 못하는 경로
        (steer_request yield 지점에서 스트림 중단 등)에서 호출부의 finally가 불러
        Future가 _pending에 영구히 남는 누수를 막는다."""
        self._pending.pop(req_id, None)

    async def wait(self, req_id: str, timeout: float):
        """사용자 결정을 기다린다. 반환: 결정 dict, 또는 None(시간 초과·미지의 id).

        None은 '거절'이 아니라 '결정 없음'이다 — 호출부가 기본 동작(그대로 진행)으로
        물러선다. timeout 0 이하 = 무제한 대기. 어떤 경로로 끝나든 요청을 제거한다.
        """
        fut = self._pending.get(req_id)
        if fut is None:
            return None
        try:
            return await asyncio.wait_for(fut, timeout=timeout if timeout > 0 else None)
        except asyncio.TimeoutError:
            return None
        finally:
            self._pending.pop(req_id, None)

    def resolve(self, req_id: str, decision: dict) -> bool:
        """브라우저의 결정(dict)을 반영한다. 반환: 유효한 대기 요청이었는지."""
        fut = self._pending.get(req_id)
        if fut is None or fut.done():
            return False
        fut.set_result(dict(decision))
        return True
