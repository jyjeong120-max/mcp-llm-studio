"""approvals.py

위험 도구 호출의 사용자 승인 브로커 — Claude Code의 권한 프롬프트와 같은 방식.

에이전트 루프가 승인이 필요한 도구 호출을 만나면:
    1. create()로 요청 id를 만들고 SSE로 approval_request 이벤트를 흘린다
    2. wait()로 사용자의 결정을 기다린다 (브라우저가 승인/거절 버튼을 띄움)
    3. 브라우저가 POST /api/chat/approve 로 resolve()를 부르면 wait가 풀린다

상태는 전부 RAM(asyncio Future)이다 — 대화 저장과 무관하고, 스트림이 끊기면
(사용자 중단) 대기도 함께 취소된다. 시간 초과는 '거절'로 처리해 실행하지 않는다.
"""

from __future__ import annotations

import asyncio
import secrets


class ApprovalBroker:
    """승인 대기 중인 요청들을 id → Future로 들고 있는 단순 브로커."""

    def __init__(self):
        self._pending: dict[str, asyncio.Future] = {}

    def create(self) -> str:
        """새 승인 요청을 등록하고 요청 id를 돌려준다.

        id는 추측 불가능한 토큰이어야 한다 — 순차 번호(appr_1, appr_2…)를 쓰면
        /api/chat/approve를 원격에서 열거해 대기 중인 위험 도구를 대신 승인할 수 있다.
        """
        req_id = f"appr_{secrets.token_urlsafe(16)}"
        self._pending[req_id] = asyncio.get_running_loop().create_future()
        return req_id

    def discard(self, req_id: str) -> None:
        """대기 항목을 제거한다(멱등). create() 후 wait()에 도달하지 못하는 경로
        (approval_request yield 지점에서 스트림 중단 등)에서 호출부의 finally가
        불러 Future가 _pending에 영구히 남는 누수를 막는다."""
        self._pending.pop(req_id, None)

    async def wait(self, req_id: str, timeout: float) -> tuple[bool, bool]:
        """사용자 결정을 기다린다. 반환: (승인 여부, 시간 초과 여부).

        시간 초과·미지의 id는 (False, ...) — 실행하지 않는 쪽으로 물러선다.
        어떤 경로로 끝나든 요청을 목록에서 제거한다.
        """
        fut = self._pending.get(req_id)
        if fut is None:
            return False, False
        try:
            approved = await asyncio.wait_for(fut, timeout=timeout if timeout > 0 else None)
            return bool(approved), False
        except asyncio.TimeoutError:
            return False, True
        finally:
            self._pending.pop(req_id, None)

    def resolve(self, req_id: str, approved: bool) -> bool:
        """브라우저의 결정을 반영한다. 반환: 유효한 대기 요청이었는지."""
        fut = self._pending.get(req_id)
        if fut is None or fut.done():
            return False
        fut.set_result(bool(approved))
        return True
