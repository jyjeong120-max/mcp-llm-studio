"""test_outlook.py — outlook_server.py 스모크 테스트

outlook_server의 도구들을 실제 Outlook에 대고 한 번씩 호출해 pass/fail을 찍습니다.
AITF의 test_tools.py와 같은 취지의 수동 점검 스크립트입니다.

안전 원칙 (중요)
    이 스크립트는 되돌리기 어려운 동작을 절대 실행하지 않습니다.
    - 🟢 읽기: 기본으로 전부 실행(부작용 없음).
    - 🟡 로컬 생성(초안/일정/연락처/작업): --create 플래그를 줄 때만 실행하고,
      만든 항목은 곧바로 지운편지함으로 삭제해 정리합니다.
    - 🔴 발송/이동/삭제/회의응답: 프리뷰 경로(confirm=False)만 호출합니다.
      즉 "승인 필요" 안내가 제대로 나오는지만 확인하고, 실제로는 아무것도 보내거나
      지우지 않습니다. 이 스크립트는 confirm=True를 어디에서도 넘기지 않습니다.

사용법
    cd Examples 가 아니라 outlook_server.py 가 있는 폴더(AITF)에서 실행하세요.
        python test_outlook.py                # 🟢 읽기 + 🔴 프리뷰만
        python test_outlook.py --create       # 🟡 생성(초안/일정/연락처/작업)까지 (생성 후 삭제)
        python test_outlook.py --folder inbox --limit 5
"""

from __future__ import annotations

import argparse
import re
import sys

# 한국어 Windows 콘솔은 기본 cp949라 이모지/엠대시에서 죽는다. utf-8로 재설정.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import outlook_server as ots


# ─────────────────────────── 결과 집계/출력 헬퍼 ───────────────────────────

PASS, WARN, FAIL, SKIP = "PASS", "WARN", "FAIL", "SKIP"
_ICON = {PASS: "✅", WARN: "⚠️ ", FAIL: "❌", SKIP: "⏭️ "}
_results: list[tuple[str, str, str]] = []

# Outlook 미연결/정책 제한을 나타내는 신호(치명적 실패가 아니라 환경 문제 → WARN).
_WARN_SIGNS = (
    "Outlook에 접근하지 못했습니다",
    "Outlook을 시작하거나 연결하지 못했습니다",
    "pywin32를 불러올 수 없습니다",
    "정책상 조회 불가",
    "정책 제한",
)


def _classify(result: str, expect_substr: str | None) -> tuple[str, str]:
    """도구가 돌려준 문자열을 보고 PASS/WARN/FAIL로 분류한다."""
    if not isinstance(result, str):
        return FAIL, f"문자열이 아님: {type(result).__name__}"
    snippet = result.replace("\n", " ⏎ ")
    if len(snippet) > 160:
        snippet = snippet[:160] + "…"
    if expect_substr is not None:
        if expect_substr in result:
            return PASS, snippet
        return FAIL, f"기대 문자열 '{expect_substr}' 없음 → {snippet}"
    if any(sign in result for sign in _WARN_SIGNS):
        return WARN, snippet
    if result.startswith("작업에 실패했습니다") or "호출이 실패했습니다" in result:
        return FAIL, snippet
    return PASS, snippet


def check(label: str, fn, *, expect: str | None = None):
    """도구 하나를 호출하고 결과를 집계·출력한다. 반환 문자열을 그대로 돌려준다."""
    try:
        result = fn()
    except Exception as e:  # 도구는 예외를 삼키게 돼 있으니 여기 오면 진짜 버그다.
        _results.append((label, FAIL, f"예외 {type(e).__name__}: {e}"))
        print(f"{_ICON[FAIL]} {label}: 예외 {type(e).__name__}: {e}")
        return ""
    status, note = _classify(result, expect)
    _results.append((label, status, note))
    print(f"{_ICON[status]} {label}: {note}")
    return result


def skip(label: str, why: str):
    _results.append((label, SKIP, why))
    print(f"{_ICON[SKIP]} {label}: {why}")


def _first_id(text: str) -> str:
    """도구 출력의 'id: <EntryID>' 중 첫 번째를 뽑는다. 없으면 빈 문자열."""
    m = re.search(r"id:\s*([0-9A-Fa-f]{16,})", text or "")
    return m.group(1) if m else ""


# ─────────────────────────────── 테스트 본문 ───────────────────────────────


def main() -> int:
    ap = argparse.ArgumentParser(description="outlook_server 스모크 테스트")
    ap.add_argument("--folder", default="inbox", help="읽기 테스트에 쓸 폴더(기본 inbox)")
    ap.add_argument("--limit", type=int, default=5, help="목록 조회 개수(기본 5)")
    ap.add_argument(
        "--create",
        action="store_true",
        help="🟡 로컬 생성(초안/일정/연락처/작업)까지 테스트(생성 후 삭제).",
    )
    args = ap.parse_args()

    print("=" * 68)
    print("outlook_server 스모크 테스트")
    print(f"  COM 사용 가능: {ots.COM_AVAILABLE}"
          + ("" if ots.COM_AVAILABLE else f"  ({ots.COM_IMPORT_ERROR})"))
    print(f"  대상 폴더: {args.folder}   |   생성 테스트(--create): {args.create}")
    print("  (이 스크립트는 confirm=True를 절대 넘기지 않습니다 — 발송/삭제 없음)")
    print("=" * 68)

    if not ots.COM_AVAILABLE:
        print("\npywin32(COM)를 불러올 수 없어 실제 호출을 건너뜁니다.")
        print("Windows + Outlook + pywin32 환경에서 다시 실행하세요.")
        return 1

    # ── A. 연결·개요 (🟢) ──
    print("\n[A] 연결·개요")
    check("get_status", ots.get_status)
    check("list_accounts", ots.list_accounts)
    check("list_folders", lambda: ots.list_folders(include_counts=True))

    # ── B. 메일 읽기 (🟢) ──
    print("\n[B] 메일 읽기")
    listing = check("list_messages", lambda: ots.list_messages(folder=args.folder, limit=args.limit))
    check("search_messages", lambda: ots.search_messages(query="a", folders=args.folder, limit=3))
    base = check("poll_new_mail(기준선)", lambda: ots.poll_new_mail(since=""), expect="checkpoint:")
    # 방금 받은 checkpoint로 재호출 → since 왕복이 도는지 확인.
    m = re.search(r"checkpoint:\s*([0-9\- :]+)", base or "")
    if m:
        cp = m.group(1).strip()
        check("poll_new_mail(since 왕복)", lambda: ots.poll_new_mail(since=cp), expect="checkpoint:")
    else:
        skip("poll_new_mail(since 왕복)", "기준선 checkpoint를 파싱하지 못함")

    eid = _first_id(listing)
    if eid:
        check("read_message", lambda: ots.read_message(entry_id=eid, max_chars=500))
        check("list_attachments", lambda: ots.list_attachments(entry_id=eid))
        check("get_conversation", lambda: ots.get_conversation(entry_id=eid, limit=5))
    else:
        for t in ("read_message", "list_attachments", "get_conversation"):
            skip(t, f"'{args.folder}'에서 메일 id를 찾지 못함")
    # 화면 선택 의존 도구는 결과가 환경에 좌우되므로 존재 여부만 가볍게.
    check("get_open_or_selected", ots.get_open_or_selected)

    # ── F/G/H. 조회 (🟢) ──
    print("\n[F/G/H] 캘린더·연락처·작업 조회")
    check("list_appointments", lambda: ots.list_appointments(limit=args.limit))
    check("list_contacts", lambda: ots.list_contacts(limit=args.limit))
    check("list_tasks", lambda: ots.list_tasks(limit=args.limit))

    # ── E. 🔴 프리뷰 경로만 (confirm=False → "승인 필요"만 확인, 실제 동작 없음) ──
    print("\n[E] 🔴 발송/이동/삭제/회의 — 프리뷰만(confirm 없음, 실제 실행 안 함)")
    check(
        "send_email(프리뷰)",
        lambda: ots.send_email(to="nobody@example.com", subject="[테스트] 발송 안 됨", body="프리뷰 확인용"),
        expect="승인 필요",
    )
    check(
        "respond_message(send=True 프리뷰)",
        lambda: ots.respond_message(entry_id=eid or "x", mode="reply", body="테스트", send=True),
        expect="승인 필요" if eid else None,
    ) if eid else skip("respond_message(send=True 프리뷰)", "원본 메일 id 없음")
    if eid:
        check("send_draft(프리뷰)", lambda: ots.send_draft(entry_id=eid), expect="승인 필요")
        check("move_message(프리뷰)", lambda: ots.move_message(entry_ids=eid, folder="inbox"), expect="승인 필요")
        check("delete_message(프리뷰)", lambda: ots.delete_message(entry_ids=eid), expect="승인 필요")
    else:
        for t in ("send_draft(프리뷰)", "move_message(프리뷰)", "delete_message(프리뷰)"):
            skip(t, "대상 메일 id 없음")
    check(
        "create_meeting(프리뷰)",
        lambda: ots.create_meeting(
            subject="[테스트] 회의 안 만들어짐", start="2030-01-01 10:00",
            attendees="nobody@example.com",
        ),
        expect="승인 필요",
    )
    check(
        "respond_meeting(프리뷰)",
        lambda: ots.respond_meeting(entry_id=eid or "x", response="accept"),
        expect="승인 필요" if eid else None,
    ) if eid else skip("respond_meeting(프리뷰)", "회의 요청 id 없음")

    # ── 🟡 로컬 생성 (--create 일 때만; 생성 후 삭제로 정리) ──
    print("\n[D/F/G/H] 🟡 로컬 생성" + ("" if args.create else " (건너뜀 — --create 로 활성화)"))
    if args.create:
        created_ids: list[str] = []

        r = check("create_draft", lambda: ots.create_draft(
            to="nobody@example.com", subject="[테스트] 초안", body="스모크 테스트 초안", display=False))
        did = _first_id(r)
        if did:
            created_ids.append(did)

        r = check("create_appointment", lambda: ots.create_appointment(
            subject="[테스트] 일정", start="2030-01-01 09:00", end="2030-01-01 10:00",
            reminder_minutes=0))
        created_ids.append(_first_id(r))

        r = check("create_contact", lambda: ots.create_contact(
            full_name="[테스트] 홍길동", email="nobody@example.com", company="테스트회사"))
        created_ids.append(_first_id(r))

        r = check("create_task", lambda: ots.create_task(
            subject="[테스트] 작업", due="2030-01-01"))
        tid = _first_id(r)
        created_ids.append(tid)
        if tid:
            check("complete_task", lambda: ots.complete_task(entry_id=tid))

        # 정리: 만든 항목을 지운편지함으로 삭제(테스트 흔적 제거).
        cleaned = 0
        for cid in [c for c in created_ids if c]:
            try:
                ots._require_com()
                import pythoncom
                pythoncom.CoInitialize()
                try:
                    ots._ns().GetItemFromID(cid).Delete()
                    cleaned += 1
                finally:
                    pythoncom.CoUninitialize()
            except Exception:
                pass
        print(f"   ↳ 정리: 생성 항목 {cleaned}개 삭제(지운편지함).")
    else:
        for t in ("create_draft", "create_appointment", "create_contact", "create_task", "complete_task"):
            skip(t, "--create 미지정")

    # ── 요약 ──
    print("\n" + "=" * 68)
    counts = {s: sum(1 for _, st, _ in _results if st == s) for s in (PASS, WARN, FAIL, SKIP)}
    print(f"요약: ✅ {counts[PASS]}   ⚠️ {counts[WARN]}   ❌ {counts[FAIL]}   ⏭️ {counts[SKIP]}"
          f"   (총 {len(_results)})")
    if counts[FAIL]:
        print("\n실패 항목:")
        for label, st, note in _results:
            if st == FAIL:
                print(f"  ❌ {label}: {note}")
    print("=" * 68)
    return 1 if counts[FAIL] else 0


if __name__ == "__main__":
    sys.exit(main())
