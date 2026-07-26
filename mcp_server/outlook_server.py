"""outlook_server.py

Outlook의 메일 / 캘린더 / 연락처 / 작업을 다루는 MCP 서버입니다.
pywin32(COM)로 로그인된 Outlook을 직접 구동하므로, 지금 사용자의 사서함
(받은편지함/보낸편지함/사용자 폴더/일정 등)을 그대로 읽고 쓸 수 있습니다.

STDIO 트랜스포트라, 클라이언트가 필요한 시점에 이 스크립트를 직접 실행합니다.

Windows + Outlook(로그인된 프로필) 설치가 전제입니다. 둘 중 하나라도 없으면
서버는 그대로 실행되고, 각 도구가 실패 사유를 담은 안내 메시지를 반환합니다.

안전 등급 (3티어)
    🟢 읽기(부작용 없음): 목록/검색/상세/첨부 저장/일정·연락처·작업 조회
    🟡 로컬 생성(비파괴): 메일 '초안'만 만들기, 답장/전달 '초안'만 만들기, 일정/연락처/작업 생성
    🔴 외부 발송·파괴(되돌리기 어려움): send_email, send_draft, respond_message(send=True),
       move_message, delete_message, create_meeting, respond_meeting
    🔴 도구는 반드시 confirm=True를 받아야 실행됩니다. confirm 없이 호출하면
    "누구에게/무슨 제목/무슨 동작"을 요약한 프리뷰만 돌려주고 실제 동작은 하지 않습니다.
    클라이언트(AITF 봇 등)는 이 도구 이름들을 HumanInTheLoopMiddleware의 INTERRUPT_ON에
    올려 사람 승인을 한 겹 더 걸 수 있습니다(서버 confirm 게이트와 이중 안전장치).

트리거(새 메일 감지)
    이 서버는 스스로 신호를 밀어주지 않습니다(MCP는 요청-응답). 대신 poll_new_mail로
    "마지막 확인 시각 이후 도착한 메일"을 폴링하세요. n8n의 Schedule 트리거나
    에이전트 루프에서 주기적으로 호출하면 됩니다. 반환된 checkpoint를 다음 호출의
    since로 넘기면 중복 없이 이어집니다.

항목 참조 규칙
    각 항목은 Outlook의 안정적인 문자열 핸들인 EntryID로 가리킵니다. 목록/검색 도구가
    entry_id를 돌려주고, 상세/후속 도구는 그 entry_id를 받습니다.

보안 경고창(Programmatic Access)
    Outlook은 발신자 주소/수신자 주소 조회나 발송 같은 민감 동작에서 경고창을 띄우거나
    차단할 수 있습니다(회사 정책/GPO에 좌우). 경고창이 뜨면 COM 호출이 그 앞에서 멈추므로,
    민감 속성은 우회 조회를 시도하고 실패하면 대화상자 대신 빈 값/안내로 물러섭니다.
"""

from __future__ import annotations

import os
import re
import threading
from contextlib import contextmanager
from datetime import datetime, timedelta

from fastmcp import FastMCP

try:
    import pythoncom
    import win32com.client

    COM_AVAILABLE = True
    COM_IMPORT_ERROR = ""
except ImportError as e:  # Windows가 아니거나 pywin32 미설치
    COM_AVAILABLE = False
    COM_IMPORT_ERROR = str(e)

mcp = FastMCP(
    name="outlook",
    instructions=(
        "Outlook의 메일/캘린더/연락처/작업을 다루는 MCP 서버입니다. "
        "메일을 지칭할 땐 목록/검색이 돌려준 entry_id를 사용하세요. "
        "어떤 폴더가 있는지 모르면 list_folders를, 새로 온 메일을 확인하려면 "
        "poll_new_mail을 호출하세요. 실제 발송/삭제/이동/회의응답 같은 되돌리기 어려운 "
        "동작은 반드시 confirm=True로 호출해야 실행되며, confirm 없이 부르면 프리뷰만 "
        "돌려줍니다. 발송 전에는 프리뷰로 받는사람/제목을 사용자에게 확인시키세요."
    ),
)

# ── 출력 상한(컨텍스트 보호). 도구 인자로 조정 가능 ──
MAX_ITEMS = 50
MAX_CHARS = 20000
MAX_BODY = 8000

# ── Outlook 열거형 상수 (win32com 동적 디스패치는 상수를 노출하지 않아 직접 정의) ──
# OlDefaultFolders
OL_FOLDER = {
    "deleted": 3, "outbox": 4, "sent": 5, "inbox": 6,
    "calendar": 9, "contacts": 10, "journal": 11, "notes": 12,
    "tasks": 13, "drafts": 16, "junk": 23,
}
# OlItemType (CreateItem)
OL_MAIL_ITEM = 0
OL_APPOINTMENT_ITEM = 1
OL_CONTACT_ITEM = 2
OL_TASK_ITEM = 3
# OlObjectClass
OL_CLASS_MAIL = 43
OL_CLASS_APPOINTMENT = 26
OL_CLASS_CONTACT = 40
OL_CLASS_TASK = 48
OL_CLASS_MEETING_REQUEST = 53
# OlImportance
OL_IMPORTANCE = {"low": 0, "normal": 1, "high": 2}
# OlBusyStatus
OL_BUSY = {"free": 0, "tentative": 1, "busy": 2, "oof": 3, "workingelsewhere": 4}
# OlTaskStatus
OL_TASK_COMPLETE = 2
# OlSaveAsType (export)
OL_SAVE_AS = {"txt": 0, "rtf": 1, "msg": 3, "html": 5, "mhtml": 6, "ical": 8}
# OlMeetingStatus / OlMeetingResponse (create_meeting, respond_meeting)
OL_MEETING = 1  # olMeeting — AppointmentItem을 회의로 바꿔 초대장을 발송
OL_MEETING_RESPONSE = {
    "accept": 3, "accepted": 3, "수락": 3,
    "tentative": 2, "미정": 2,
    "decline": 4, "declined": 4, "거절": 4,
}

# PropertyAccessor 태그 — 보안 경고창을 피해 SMTP 주소를 우회 조회할 때 쓴다.
PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"

# poll_new_mail이 클라이언트에게 돌려주는 checkpoint의 형식.
CHECKPOINT_FMT = "%Y-%m-%d %H:%M:%S"


class OutlookError(Exception):
    """도구가 사용자에게 그대로 돌려줄 안내 메시지를 담은 예외."""


_ctx = threading.local()


def _require_com():
    if not COM_AVAILABLE:
        raise OutlookError(
            "Outlook에 접근하지 못했습니다. pywin32를 불러올 수 없습니다"
            f"({COM_IMPORT_ERROR}). Windows에서 pywin32가 설치된 환경으로 실행하세요."
        )


def outlook_tool(fn):
    """COM 초기화와 예외 처리를 감싸는 도구 데코레이터.

    FastMCP는 동기 도구를 워커 스레드에서 실행한다. COM은 스레드마다
    CoInitialize가 필요하므로 매 호출마다 초기화하고 해제한다.
    """
    from functools import wraps

    @wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            _require_com()
        except OutlookError as e:
            return str(e)

        pythoncom.CoInitialize()
        try:
            return fn(*args, **kwargs)
        except OutlookError as e:
            return str(e)
        except pythoncom.com_error as e:
            return f"Outlook 호출이 실패했습니다: {_com_message(e)}"
        except Exception as e:
            return f"작업에 실패했습니다: {type(e).__name__}: {e}"
        finally:
            pythoncom.CoUninitialize()

    return wrapper


def _com_message(e) -> str:
    try:
        info = e.excepinfo
        if info and len(info) > 2 and info[2]:
            return str(info[2]).strip()
    except Exception:
        pass
    return str(e)


# ─────────────────────────────── 연결 헬퍼 ───────────────────────────────


def _app():
    """실행 중인(없으면 새로 기동되는) 단일 Outlook 인스턴스에 붙는다."""
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except pythoncom.com_error as e:
        raise OutlookError(
            "Outlook을 시작하거나 연결하지 못했습니다: "
            f"{_com_message(e)}. Outlook이 설치·로그인되어 있는지 확인하세요."
        )


def _ns():
    """MAPI 네임스페이스. 이미 로그온돼 있으면 Logon은 조용히 넘어간다."""
    ns = _app().GetNamespace("MAPI")
    try:
        ns.Logon("", "", False, False)
    except pythoncom.com_error:
        pass  # Outlook이 이미 프로필로 로그온된 상태면 정상
    return ns


def _split(s: str) -> list[str]:
    """세미콜론/쉼표/줄바꿈으로 구분된 문자열을 리스트로. 빈 항목은 버린다."""
    if not s:
        return []
    return [p.strip() for p in re.split(r"[;,\n]+", s) if p.strip()]


def _resolve_folder(spec: str):
    """폴더를 별칭 또는 경로로 찾는다.

    - 별칭: inbox/sent/drafts/deleted/outbox/junk/calendar/contacts/tasks 등
    - 경로: '보관\\프로젝트' 또는 '내 계정\\받은 편지함\\하위' (역슬래시 구분)
    - 생략/빈값: 받은편지함
    """
    ns = _ns()
    spec = (spec or "").strip()
    if not spec or spec.lower() in OL_FOLDER:
        return ns.GetDefaultFolder(OL_FOLDER.get(spec.lower(), OL_FOLDER["inbox"]))

    parts = [p for p in re.split(r"[\\/]+", spec) if p]
    # 첫 조각이 스토어(계정) 이름이면 그 아래부터, 아니면 기본 스토어에서 탐색.
    try:
        folder = None
        roots = {f.Name: f for f in ns.Folders}
        if parts and parts[0] in roots:
            folder = roots[parts[0]]
            parts = parts[1:]
        else:
            folder = ns.GetDefaultFolder(OL_FOLDER["inbox"]).Parent  # 기본 스토어 루트
        for name in parts:
            folder = folder.Folders[name]
        return folder
    except pythoncom.com_error:
        raise OutlookError(
            f"'{spec}' 폴더를 찾지 못했습니다. 별칭(inbox/sent/drafts/calendar 등) "
            "또는 '계정\\폴더\\하위' 경로로 지정하세요. list_folders로 확인할 수 있습니다."
        )


def _get_item(entry_id: str):
    """EntryID로 항목을 가져온다."""
    entry_id = (entry_id or "").strip()
    if not entry_id:
        raise OutlookError("entry_id가 비어 있습니다.")
    try:
        return _ns().GetItemFromID(entry_id)
    except pythoncom.com_error:
        raise OutlookError(
            f"entry_id '{entry_id[:16]}…'에 해당하는 항목을 찾지 못했습니다. "
            "목록/검색 도구가 돌려준 최신 entry_id인지 확인하세요."
        )


# ─────────────────────────────── 포맷 헬퍼 ───────────────────────────────


def _fmt_dt(v) -> str:
    if not v:
        return ""
    try:
        return v.strftime("%Y-%m-%d %H:%M")
    except Exception:
        try:
            return v.Format("%Y-%m-%d %H:%M")
        except Exception:
            return str(v)


def _parse_dt(s: str, end_of_day: bool = False):
    """'YYYY-MM-DD' 또는 'YYYY-MM-DD HH:MM'을 datetime으로. 빈값이면 None."""
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(s, fmt)
            if fmt == "%Y-%m-%d" and end_of_day:
                dt = dt.replace(hour=23, minute=59, second=59)
            return dt
        except ValueError:
            continue
    raise OutlookError(f"날짜 형식이 잘못되었습니다: '{s}' (예: 2026-07-17 또는 2026-07-17 09:30)")


def _ol_time(dt: datetime) -> str:
    """Restrict 필터에 넣을 Outlook 로컬 시간 문자열."""
    return dt.strftime("%m/%d/%Y %I:%M %p")


def _clean(text) -> str:
    if not text:
        return ""
    return (
        str(text)
        .replace("\x07", " ")
        .replace("\x0b", "\n")
        .replace("\r", "\n")
        .strip()
    )


def _truncate(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + f"\n\n…(전체 {len(text):,}자 중 {limit:,}자만 표시)"


def _preview(action: str, details: list[str], tool_hint: str = "") -> str:
    """🔴 도구가 confirm 없이 호출됐을 때 돌려줄 '실행 전 확인' 프리뷰.

    실제 동작은 하지 않고, 무엇을 할지 요약만 보여준다. 클라이언트/사용자가
    내용을 확인한 뒤 같은 도구를 confirm=True로 다시 부르면 그때 실행된다.
    """
    lines = [f"⚠️ 승인 필요 — 아직 실행하지 않았습니다: {action}", ""]
    lines.extend(f"  {d}" for d in details)
    lines.append("")
    lines.append(
        "이대로 진행하려면 같은 도구를 confirm=true 로 다시 호출하세요. "
        + (tool_hint or "")
    )
    return "\n".join(lines).rstrip()


def _sender_email(item) -> str:
    """발신자 SMTP 주소를 보안 경고창을 피해 얻는다. 못 얻으면 빈 문자열."""
    for getter in (
        lambda: item.PropertyAccessor.GetProperty(PR_SENDER_SMTP),
        lambda: item.Sender.GetExchangeUser().PrimarySmtpAddress,
        lambda: item.SenderEmailAddress,
    ):
        try:
            v = getter()
            if v:
                return str(v)
        except Exception:
            continue
    return ""


def _recipients(item, kinds=(1, 2, 3)) -> list[str]:
    """수신자 이름 목록. kinds: 1=To, 2=CC, 3=BCC (OlMailRecipientType)."""
    out = []
    try:
        for r in item.Recipients:
            try:
                if r.Type in kinds:
                    out.append(r.Name)
            except Exception:
                out.append(getattr(r, "Name", "?"))
    except Exception:
        pass
    return out


def _has_attach(item) -> bool:
    try:
        return item.Attachments.Count > 0
    except Exception:
        return False


def _is_mail(item) -> bool:
    try:
        return item.Class in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST)
    except Exception:
        return False


def _mail_summary(item) -> str:
    """메일 한 통을 여러 줄 요약으로. 후속 조회용 EntryID 포함."""
    subject = _clean(getattr(item, "Subject", "")) or "(제목 없음)"
    sender = _clean(getattr(item, "SenderName", "")) or "(발신자 없음)"
    smtp = _sender_email(item)
    sender_str = f"{sender} <{smtp}>" if smtp else sender
    received = _fmt_dt(getattr(item, "ReceivedTime", None))
    marks = []
    try:
        if item.UnRead:
            marks.append("안읽음")
    except Exception:
        pass
    if _has_attach(item):
        try:
            marks.append(f"📎{item.Attachments.Count}")
        except Exception:
            marks.append("📎")
    cats = _clean(getattr(item, "Categories", ""))
    if cats:
        marks.append(cats)
    mark = f"  [{', '.join(marks)}]" if marks else ""
    eid = getattr(item, "EntryID", "")
    return f"• {subject}{mark}\n    {sender_str}  |  {received}\n    id: {eid}"


def _restricted(folder, since=None, until=None, unread_only=False, newest_first=True):
    """폴더 항목을 시간/읽음으로 서버측 Restrict 후 정렬해 돌려준다."""
    items = folder.Items
    clauses = []
    if since:
        clauses.append(f"[ReceivedTime] >= '{_ol_time(since)}'")
    if until:
        clauses.append(f"[ReceivedTime] <= '{_ol_time(until)}'")
    if unread_only:
        clauses.append("[UnRead] = True")
    try:
        items.Sort("[ReceivedTime]", newest_first)
    except Exception:
        pass
    if clauses:
        try:
            items = items.Restrict(" AND ".join(clauses))
        except pythoncom.com_error:
            pass  # Restrict 실패 시 파이썬 측 필터로 폴백
    return items


# ═══════════════════════════════ A. 연결·개요 ═══════════════════════════════


@mcp.tool()
@outlook_tool
def get_status() -> str:
    """Outlook 연결 상태와 현재 사용자/기본 사서함을 조회합니다.

    다른 도구가 실패할 때 먼저 호출해 Outlook이 로그인되어 있는지 확인하는 용도입니다.

    Returns:
        Outlook 버전, 현재 사용자, 기본 사서함, 받은편지함 전체/안읽음 수.
    """
    ns = _ns()
    out = ["Outlook 연결됨"]
    try:
        out.append(f"버전: {_app().Version}")
    except Exception:
        pass
    try:
        out.append(f"현재 사용자: {ns.CurrentUser.Name}")
    except Exception:
        pass
    try:
        inbox = ns.GetDefaultFolder(OL_FOLDER["inbox"])
        out.append(f"기본 사서함: {inbox.Parent.Name}")
        out.append(f"받은편지함: 전체 {inbox.Items.Count}통, 안읽음 {inbox.UnReadItemCount}통")
    except Exception:
        pass
    return "\n".join(out)


@mcp.tool()
@outlook_tool
def list_folders(account: str = "", include_counts: bool = True) -> str:
    """사서함의 폴더 트리를 조회합니다. 각 폴더의 전체/안읽음 수도 함께 봅니다.

    사용자 폴더의 정확한 이름/경로를 몰라 list_messages의 folder에 무엇을 넣을지
    막힐 때 사용합니다.

    Args:
        account: 특정 계정(스토어)만 볼 때 이름 지정. 생략하면 모든 스토어.
        include_counts: 각 폴더의 항목 수/안읽음 수 표시 여부.

    Returns:
        들여쓰기된 폴더 트리와 (옵션) 항목 수.
    """
    ns = _ns()
    lines = []

    def walk(folder, depth):
        indent = "  " * depth
        suffix = ""
        if include_counts:
            try:
                suffix = f"  ({folder.Items.Count}통, 안읽음 {folder.UnReadItemCount})"
            except Exception:
                suffix = ""
        lines.append(f"{indent}- {folder.Name}{suffix}")
        try:
            for sub in folder.Folders:
                walk(sub, depth + 1)
        except Exception:
            pass

    for store in ns.Folders:
        if account and store.Name != account:
            continue
        lines.append(f"[{store.Name}]")
        try:
            for sub in store.Folders:
                walk(sub, 1)
        except Exception:
            pass
    if not lines:
        return "폴더를 찾지 못했습니다." + (f" (계정 '{account}' 없음)" if account else "")
    return "\n".join(lines)


@mcp.tool()
@outlook_tool
def list_accounts() -> str:
    """구성된 메일 계정(SMTP 주소)과 스토어 목록을 조회합니다.

    Returns:
        계정별 표시 이름·SMTP 주소, 그리고 열려 있는 스토어(사서함) 이름.
    """
    ns = _ns()
    out = ["[계정]"]
    try:
        for acct in _app().Session.Accounts:
            try:
                out.append(f"  - {acct.DisplayName}  <{acct.SmtpAddress}>")
            except Exception:
                out.append(f"  - {getattr(acct, 'DisplayName', '?')}")
    except Exception:
        out.append("  (계정 목록을 가져오지 못했습니다)")
    out.append("\n[스토어(사서함)]")
    try:
        for store in ns.Folders:
            out.append(f"  - {store.Name}")
    except Exception:
        pass
    return "\n".join(out)


# ═══════════════════════════════ B. 메일 읽기 ═══════════════════════════════


@mcp.tool()
@outlook_tool
def list_messages(
    folder: str = "inbox",
    limit: int = 25,
    unread_only: bool = False,
    since: str = "",
    until: str = "",
    from_sender: str = "",
    subject_contains: str = "",
    has_attachments: bool = False,
    category: str = "",
) -> str:
    """폴더의 메일 목록을 조건으로 걸러 요약합니다. (메일 조회의 기본 도구)

    필터를 파라미터로 조합하므로 "안 읽은 메일", "○○가 보낸 메일", "지난주 메일",
    "첨부 있는 메일" 등을 이 하나로 처리합니다. 각 메일의 entry_id를 함께 주므로,
    상세 내용은 read_message에 그 id를 넘겨 읽습니다.

    Args:
        folder: 폴더 별칭(inbox/sent/drafts/deleted/junk 등) 또는 '계정\\폴더' 경로.
        limit: 최대 표시 개수(기본 25).
        unread_only: 안 읽은 메일만.
        since: 이 시각 이후 수신분만. 'YYYY-MM-DD' 또는 'YYYY-MM-DD HH:MM'.
        until: 이 시각 이전 수신분만.
        from_sender: 발신자 이름/주소에 이 문자열이 포함된 메일만(부분 일치).
        subject_contains: 제목에 이 문자열이 포함된 메일만(부분 일치).
        has_attachments: 첨부가 있는 메일만.
        category: 이 분류(Category)가 지정된 메일만.

    Returns:
        조건에 맞는 메일 요약 목록(발신자/제목/수신시각/표시/entry_id).
    """
    fld = _resolve_folder(folder)
    since_dt = _parse_dt(since)
    until_dt = _parse_dt(until, end_of_day=True)
    items = _restricted(fld, since_dt, until_dt, unread_only)

    limit = max(1, min(limit, MAX_ITEMS))
    fs = from_sender.lower()
    sc = subject_contains.lower()
    cat = category.lower()

    hits, scanned = [], 0
    for item in items:
        scanned += 1
        if scanned > 5000:  # 안전장치: 지나치게 큰 폴더 순회 방지
            break
        if not _is_mail(item):
            continue
        if fs:
            who = f"{getattr(item, 'SenderName', '')} {_sender_email(item)}".lower()
            if fs not in who:
                continue
        if sc and sc not in str(getattr(item, "Subject", "")).lower():
            continue
        if has_attachments and not _has_attach(item):
            continue
        if cat and cat not in str(getattr(item, "Categories", "")).lower():
            continue
        hits.append(item)
        if len(hits) >= limit:
            break

    head = f"폴더: {fld.Name}  |  {len(hits)}통"
    if not hits:
        return head + "\n\n조건에 맞는 메일이 없습니다."
    body = "\n".join(_mail_summary(m) for m in hits)
    return f"{head}\n\n{body}"


@mcp.tool()
@outlook_tool
def poll_new_mail(
    since: str = "",
    folder: str = "inbox",
    unread_only: bool = True,
    limit: int = 50,
) -> str:
    """지정 시각(since) 이후 도착한 새 메일을 조회합니다. (폴링 트리거용)

    n8n의 Schedule 트리거나 에이전트 루프에서 주기적으로 호출하는 용도입니다.
    반환 끝의 checkpoint 값을 다음 호출의 since로 그대로 넘기면, 중복 없이
    이어서 새 메일만 받습니다. since를 비우면 '지금부터' 기준선만 잡고 끝냅니다
    (즉 최초 1회는 빈 결과 + 현재 checkpoint를 돌려줍니다).

    Args:
        since: 이 시각 이후 도착분만. 직전 호출이 준 checkpoint를 넣으세요.
            비우면 결과 없이 현재 시각 checkpoint만 반환.
        folder: 감시할 폴더(기본 inbox).
        unread_only: 안 읽은 메일만(기본 True).
        limit: 한 번에 가져올 최대 개수.

    Returns:
        새 메일 요약(오래된 것부터)과 마지막 줄의 `checkpoint: <시각>`.
    """
    now = datetime.now()
    if not since:
        return (
            "기준선을 잡았습니다. 이후 호출부터 새 메일을 감지합니다.\n"
            f"checkpoint: {now.strftime(CHECKPOINT_FMT)}"
        )

    fld = _resolve_folder(folder)
    since_dt = _parse_dt(since)
    items = _restricted(fld, since_dt, None, unread_only, newest_first=False)

    limit = max(1, min(limit, MAX_ITEMS))
    hits = []
    latest = since_dt
    for item in items:
        if not _is_mail(item):
            continue
        rt = getattr(item, "ReceivedTime", None)
        # Restrict 경계값(>=)이 같은 메일을 다시 잡지 않도록 since와 동일 시각은 건너뛴다.
        try:
            rt_naive = datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)
            if since_dt and rt_naive <= since_dt:
                continue
            if latest is None or rt_naive > latest:
                latest = rt_naive
        except Exception:
            pass
        hits.append(item)
        if len(hits) >= limit:
            break

    checkpoint = (latest or now).strftime(CHECKPOINT_FMT)
    if not hits:
        return f"새 메일 없음.\ncheckpoint: {checkpoint}"
    body = "\n".join(_mail_summary(m) for m in hits)
    return f"새 메일 {len(hits)}통 (오래된 순):\n\n{body}\n\ncheckpoint: {checkpoint}"


@mcp.tool()
@outlook_tool
def search_messages(
    query: str,
    folders: str = "inbox,sent",
    since: str = "",
    limit: int = 25,
) -> str:
    """여러 폴더에 걸쳐 메일을 검색합니다.

    발신자/제목/본문 미리보기에 검색어가 포함된 메일을 찾습니다(부분 일치, 대소문자 무시).
    특정 폴더 하나만 정밀 필터링할 땐 list_messages가 더 낫습니다.

    Args:
        query: 찾을 문자열.
        folders: 검색할 폴더들(쉼표 구분 별칭/경로). 기본 'inbox,sent'.
        since: 이 시각 이후 수신분만 검색(큰 폴더에서 속도 확보).
        limit: 최대 표시 개수.

    Returns:
        폴더별로 일치한 메일 요약 목록.
    """
    if not query.strip():
        return "검색어가 비어 있습니다."
    needle = query.lower()
    since_dt = _parse_dt(since)
    limit = max(1, min(limit, MAX_ITEMS))

    out, total = [], 0
    for spec in _split(folders) or ["inbox"]:
        try:
            fld = _resolve_folder(spec)
        except OutlookError as e:
            out.append(f"[{spec}] {e}")
            continue
        items = _restricted(fld, since_dt, None, False)
        found = []
        scanned = 0
        for item in items:
            scanned += 1
            if scanned > 5000 or total >= limit:
                break
            if not _is_mail(item):
                continue
            hay = " ".join([
                str(getattr(item, "Subject", "")),
                str(getattr(item, "SenderName", "")),
                _sender_email(item),
                str(getattr(item, "Body", ""))[:400],
            ]).lower()
            if needle in hay:
                found.append(item)
                total += 1
                if total >= limit:
                    break
        if found:
            out.append(f"[{fld.Name}] {len(found)}통")
            out.extend("  " + _mail_summary(m).replace("\n", "\n  ") for m in found)
    if not out:
        return f"'{query}'와 일치하는 메일이 없습니다."
    return f"검색어: '{query}'\n\n" + "\n".join(out)


@mcp.tool()
@outlook_tool
def read_message(entry_id: str, include_html: bool = False, max_chars: int = MAX_BODY) -> str:
    """메일 한 통의 전체 내용을 읽습니다.

    Args:
        entry_id: 목록/검색이 돌려준 메일 id.
        include_html: 본문을 HTML로 받을지(기본은 일반 텍스트).
        max_chars: 본문 최대 글자 수(기본 8000).

    Returns:
        발신자/수신자/참조/제목/시각/중요도/분류/첨부 목록과 본문.
    """
    item = _get_item(entry_id)
    out = []
    out.append(f"제목: {_clean(getattr(item, 'Subject', '')) or '(제목 없음)'}")
    smtp = _sender_email(item)
    out.append(f"발신: {_clean(getattr(item, 'SenderName', ''))}" + (f" <{smtp}>" if smtp else ""))
    to = _recipients(item, (1,))
    cc = _recipients(item, (2,))
    if to:
        out.append(f"받는사람: {', '.join(to)}")
    if cc:
        out.append(f"참조: {', '.join(cc)}")
    out.append(f"수신: {_fmt_dt(getattr(item, 'ReceivedTime', None))}")
    sent = _fmt_dt(getattr(item, "SentOn", None))
    if sent:
        out.append(f"보냄: {sent}")
    try:
        imp = {0: "낮음", 1: "보통", 2: "높음"}.get(item.Importance, "")
        if imp and imp != "보통":
            out.append(f"중요도: {imp}")
    except Exception:
        pass
    cats = _clean(getattr(item, "Categories", ""))
    if cats:
        out.append(f"분류: {cats}")
    try:
        if item.UnRead:
            out.append("상태: 안읽음")
    except Exception:
        pass

    atts = []
    try:
        for a in item.Attachments:
            atts.append(a.FileName)
    except Exception:
        pass
    if atts:
        out.append(f"첨부({len(atts)}): {', '.join(atts)}  → save_attachments로 저장")

    out.append("")
    if include_html:
        body = _clean(getattr(item, "HTMLBody", "") or getattr(item, "Body", ""))
    else:
        body = _clean(getattr(item, "Body", ""))
    out.append(_truncate(body, max(200, max_chars)))
    return "\n".join(out)


@mcp.tool()
@outlook_tool
def get_open_or_selected() -> str:
    """지금 Outlook에서 열려 있거나 목록에서 선택한 메일을 읽습니다.

    "지금 보고 있는 메일", "방금 클릭한 그 메일"처럼 사용자가 화면에서 지칭할 때
    entry_id 없이 바로 그 메일을 잡습니다. 열린 항목(Inspector)을 먼저 보고,
    없으면 목록(Explorer)의 선택 항목을 봅니다.

    Returns:
        열린/선택된 메일 요약(여러 개면 목록). 하나뿐이면 본문 앞부분까지.
    """
    app = _app()
    # 1) 창으로 열려 있는 항목
    try:
        insp = app.ActiveInspector()
        if insp is not None:
            item = insp.CurrentItem
            if _is_mail(item):
                return "열려 있는 메일:\n\n" + _mail_summary(item) + "\n\n" + \
                    _truncate(_clean(getattr(item, "Body", "")), 2000)
    except Exception:
        pass
    # 2) 목록에서 선택된 항목들
    try:
        sel = app.ActiveExplorer().Selection
        mails = [sel.Item(i) for i in range(1, sel.Count + 1) if _is_mail(sel.Item(i))]
    except Exception:
        mails = []
    if not mails:
        return "지금 열려 있거나 선택된 메일이 없습니다. Outlook에서 메일을 선택한 뒤 다시 호출하세요."
    if len(mails) == 1:
        m = mails[0]
        return "선택된 메일:\n\n" + _mail_summary(m) + "\n\n" + \
            _truncate(_clean(getattr(m, "Body", "")), 2000)
    body = "\n".join(_mail_summary(m) for m in mails[:MAX_ITEMS])
    return f"선택된 메일 {len(mails)}통:\n\n{body}"


@mcp.tool()
@outlook_tool
def get_conversation(entry_id: str, limit: int = 30) -> str:
    """한 메일이 속한 대화(스레드) 전체를 시간순으로 모읍니다.

    "이 메일 앞뒤로 오간 내용 정리해줘" 같은 요청에 사용합니다.

    Args:
        entry_id: 기준 메일 id.
        limit: 최대 표시 개수.

    Returns:
        같은 대화에 속한 메일 요약 목록(시간순).
    """
    item = _get_item(entry_id)
    limit = max(1, min(limit, MAX_ITEMS))
    collected = []

    # 1) Conversation API 우선
    try:
        conv = item.GetConversation()
        if conv is not None:
            def add_tree(items):
                for it in items:
                    collected.append(it)
                    try:
                        add_tree(conv.GetChildren(it))
                    except Exception:
                        pass
            add_tree(conv.GetRootItems())
    except Exception:
        collected = []

    # 2) 폴백: 같은 폴더에서 ConversationID가 같은 메일을 훑는다.
    if not collected:
        try:
            conv_id = item.ConversationID
            for it in item.Parent.Items:
                if getattr(it, "ConversationID", None) == conv_id:
                    collected.append(it)
        except Exception:
            collected = [item]

    def sort_key(it):
        try:
            return it.ReceivedTime
        except Exception:
            return getattr(it, "SentOn", None)
    try:
        collected.sort(key=sort_key)
    except Exception:
        pass

    mails = [m for m in collected if _is_mail(m)][:limit]
    if not mails:
        return "대화에 속한 메일을 찾지 못했습니다."
    topic = _clean(getattr(item, "ConversationTopic", "")) or _clean(getattr(item, "Subject", ""))
    body = "\n".join(_mail_summary(m) for m in mails)
    return f"대화: {topic}  |  {len(mails)}통\n\n{body}"


# ═══════════════════════════════ C. 첨부파일 ═══════════════════════════════


@mcp.tool()
@outlook_tool
def list_attachments(entry_id: str) -> str:
    """메일의 첨부파일 목록을 조회합니다.

    Args:
        entry_id: 메일 id.

    Returns:
        첨부별 이름·크기·인라인 여부. 저장은 save_attachments로.
    """
    item = _get_item(entry_id)
    try:
        n = item.Attachments.Count
    except Exception:
        return "이 항목에는 첨부가 없습니다."
    if n == 0:
        return "첨부가 없습니다."
    out = [f"첨부 {n}개:"]
    for i in range(1, n + 1):
        a = item.Attachments.Item(i)
        try:
            size = f"{a.Size:,} bytes"
        except Exception:
            size = "?"
        out.append(f"  [{i}] {a.FileName}  ({size})")
    return "\n".join(out)


@mcp.tool()
@outlook_tool
def save_attachments(
    entry_id: str, out_dir: str, name_filter: str = "", index: int = 0
) -> str:
    """메일의 첨부파일을 디스크에 저장합니다.

    Args:
        entry_id: 메일 id.
        out_dir: 저장할 폴더 경로(없으면 생성).
        name_filter: 파일명에 이 문자열이 포함된 첨부만 저장(부분 일치).
        index: 특정 첨부 하나만 저장할 때 번호(1부터). 0(기본)이면 조건에 맞는 전체.

    Returns:
        저장한 파일 경로 목록.
    """
    item = _get_item(entry_id)
    out_dir = os.path.abspath(os.path.expanduser(out_dir))
    os.makedirs(out_dir, exist_ok=True)
    try:
        n = item.Attachments.Count
    except Exception:
        n = 0
    if n == 0:
        return "저장할 첨부가 없습니다."

    saved = []
    nf = name_filter.lower()
    for i in range(1, n + 1):
        if index and i != index:
            continue
        a = item.Attachments.Item(i)
        fname = a.FileName
        if nf and nf not in fname.lower():
            continue
        path = os.path.join(out_dir, fname)
        try:
            a.SaveAsFile(path)
            saved.append(path)
        except pythoncom.com_error as e:
            saved.append(f"(실패: {fname} - {_com_message(e)})")
    if not saved:
        return "조건에 맞는 첨부가 없습니다."
    return "저장 완료:\n" + "\n".join("  " + s for s in saved)


@mcp.tool()
@outlook_tool
def export_message(entry_id: str, path: str, format: str = "msg") -> str:
    """메일 한 통을 파일로 내보냅니다(.msg/.txt/.html 등).

    Args:
        entry_id: 메일 id.
        path: 저장 경로.
        format: msg(기본)/txt/html/rtf/mhtml.

    Returns:
        저장 결과와 경로.
    """
    item = _get_item(entry_id)
    fmt = OL_SAVE_AS.get(format.lower())
    if fmt is None:
        return f"format은 {', '.join(OL_SAVE_AS)} 중 하나여야 합니다."
    path = os.path.abspath(os.path.expanduser(path))
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    item.SaveAs(path, fmt)
    return f"저장 완료: {path}"


# ═══════════════════════ D. 작성 (초안까지만, 🟡) ═══════════════════════


@mcp.tool()
@outlook_tool
def resolve_recipient(name: str) -> str:
    """이름/부분 문자열을 GAL·연락처에서 실제 수신자로 확인합니다.

    발송 전 "이 이름이 누구 주소로 가는지" 검증하는 용도입니다. 회사 보안 정책상
    주소 조회가 막혀 있으면 확인만 되고 주소는 비어 나올 수 있습니다.

    Args:
        name: 확인할 표시 이름 또는 부분 문자열.

    Returns:
        확인 여부와 (가능하면) 표시 이름·SMTP 주소.
    """
    ns = _ns()
    r = ns.CreateRecipient(name)
    r.Resolve()
    if not r.Resolved:
        return f"'{name}'을(를) 확인하지 못했습니다. 이름을 더 정확히 지정하거나 주소를 직접 쓰세요."
    smtp = ""
    for getter in (
        lambda: r.AddressEntry.GetExchangeUser().PrimarySmtpAddress,
        lambda: r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS),
        lambda: r.Address,
    ):
        try:
            v = getter()
            if v:
                smtp = str(v)
                break
        except Exception:
            continue
    return f"확인됨: {r.Name}" + (f"  <{smtp}>" if smtp else "  (주소는 정책상 조회 불가)")


def _fill_mail(mail, to, cc, bcc, subject, body, html, attachments, importance):
    """MailItem 공통 필드 채우기."""
    if to:
        mail.To = "; ".join(_split(to))
    if cc:
        mail.CC = "; ".join(_split(cc))
    if bcc:
        mail.BCC = "; ".join(_split(bcc))
    if subject:
        mail.Subject = subject
    if html:
        mail.HTMLBody = body
    elif body:
        mail.Body = body
    imp = OL_IMPORTANCE.get((importance or "").lower())
    if imp is not None:
        mail.Importance = imp
    for p in _split(attachments):
        path = os.path.abspath(os.path.expanduser(p))
        if os.path.exists(path):
            mail.Attachments.Add(path)
        else:
            raise OutlookError(f"첨부 파일을 찾지 못했습니다: {path}")


@mcp.tool()
@outlook_tool
def create_draft(
    to: str = "",
    subject: str = "",
    body: str = "",
    cc: str = "",
    bcc: str = "",
    html: bool = False,
    attachments: str = "",
    importance: str = "",
    display: bool = True,
) -> str:
    """새 메일 초안을 만듭니다(보내지 않습니다).

    실제 발송은 이 버전에 없습니다 — 초안을 만들어 두면 사용자가 Outlook에서 확인 후
    직접 보내거나, 다음 버전의 승인형 발송 도구로 보냅니다.

    Args:
        to/cc/bcc: 수신자(세미콜론/쉼표로 여러 명). 이름 또는 이메일 주소.
        subject: 제목.
        body: 본문.
        html: body를 HTML로 다룰지.
        attachments: 첨부할 파일 경로(세미콜론/쉼표로 여러 개).
        importance: low/normal/high.
        display: 만든 초안을 화면에 띄울지(기본 True).

    Returns:
        만들어진 초안의 entry_id와 요약.
    """
    mail = _app().CreateItem(OL_MAIL_ITEM)
    _fill_mail(mail, to, cc, bcc, subject, body, html, attachments, importance)
    mail.Save()  # 임시 보관함에 저장
    if display:
        try:
            mail.Display(False)
        except Exception:
            pass
    return (
        "초안을 만들었습니다(발송하지 않음).\n"
        f"  받는사람: {mail.To or '(비어 있음)'}\n"
        f"  제목: {mail.Subject or '(비어 있음)'}\n"
        f"  id: {mail.EntryID}"
    )


@mcp.tool()
@outlook_tool
def respond_message(
    entry_id: str,
    mode: str = "reply",
    body: str = "",
    to: str = "",
    html: bool = False,
    send: bool = False,
    confirm: bool = False,
) -> str:
    """받은 메일에 답장/전체답장/전달합니다. 기본은 '초안'만, send=True면 발송(🔴).

    send=False(기본)면 초안만 만들어 두고 사용자가 Outlook에서 직접 보냅니다.
    send=True는 실제 발송이라 🔴 동작이며 confirm=True가 있어야 나갑니다. confirm 없이
    send=True면 받는사람/제목 프리뷰만 돌려주고 보내지 않습니다.

    Args:
        entry_id: 원본 메일 id.
        mode: reply(답장) / reply_all(전체답장) / forward(전달).
        body: 맨 위에 추가할 내용(원문은 아래에 인용됨).
        to: forward일 때 받는 사람(세미콜론/쉼표).
        html: body를 HTML로 다룰지.
        send: True면 실제 발송(🔴). False면 초안만.
        confirm: send=True일 때 실제 발송하려면 True. 없으면 프리뷰만.

    Returns:
        (send=False) 만들어진 초안의 entry_id / (send=True) 발송 결과 또는 프리뷰.
    """
    item = _get_item(entry_id)
    mode = (mode or "reply").lower()
    if mode == "reply":
        draft = item.Reply()
    elif mode in ("reply_all", "replyall", "all"):
        draft = item.ReplyAll()
    elif mode == "forward":
        draft = item.Forward()
        if to:
            draft.To = "; ".join(_split(to))
    else:
        return "mode는 reply / reply_all / forward 중 하나여야 합니다."

    if body:
        if html:
            draft.HTMLBody = body + (draft.HTMLBody or "")
        else:
            draft.Body = body + "\n\n" + (draft.Body or "")

    if send:
        recips = "; ".join(_recipients(draft, (1, 2))) or draft.To or "(비어 있음)"
        if not confirm:
            return _preview(
                f"메일 {mode} 발송",
                [f"받는사람: {recips}", f"제목: {draft.Subject}",
                 f"본문 앞부분: {_truncate(_clean(draft.Body), 300)}"],
                "(respond_message ... send=true, confirm=true)",
            )
        draft.Send()
        return f"발송 완료({mode}).\n  받는사람: {recips}\n  제목: {draft.Subject}"

    draft.Save()
    try:
        draft.Display(False)
    except Exception:
        pass
    return (
        f"{mode} 초안을 만들었습니다(발송하지 않음).\n"
        f"  받는사람: {draft.To or '(비어 있음)'}\n"
        f"  제목: {draft.Subject}\n"
        f"  id: {draft.EntryID}"
    )


@mcp.tool()
@outlook_tool
def set_message_state(
    entry_ids: str,
    read: str = "",
    category: str = "",
    flag: str = "",
) -> str:
    """메일의 읽음 상태·분류·플래그를 바꿉니다(되돌릴 수 있는 로컬 변경).

    Args:
        entry_ids: 대상 메일 id들(세미콜론/쉼표로 여러 개).
        read: 'read'로 읽음, 'unread'로 안읽음 처리. 비우면 변경 안 함.
        category: 지정할 분류 이름. '-'를 주면 분류 해제. 비우면 변경 안 함.
        flag: 'flag'로 후속 작업 표시, 'complete'로 완료, 'clear'로 해제.

    Returns:
        처리한 메일 수와 적용한 변경 요약.
    """
    ids = _split(entry_ids)
    if not ids:
        return "entry_ids가 비어 있습니다."
    changed = 0
    for eid in ids:
        try:
            item = _ns().GetItemFromID(eid)
        except Exception:
            continue
        if read.lower() in ("read", "읽음"):
            item.UnRead = False
        elif read.lower() in ("unread", "안읽음"):
            item.UnRead = True
        if category == "-":
            item.Categories = ""
        elif category:
            item.Categories = category
        fl = flag.lower()
        try:
            if fl == "flag":
                item.FlagStatus = 2  # olFlagMarked
            elif fl == "complete":
                item.FlagStatus = 1  # olFlagComplete
            elif fl == "clear":
                item.FlagStatus = 0  # olNoFlag
        except Exception:
            pass
        item.Save()
        changed += 1
    parts = []
    if read:
        parts.append(f"읽음={read}")
    if category:
        parts.append("분류해제" if category == "-" else f"분류='{category}'")
    if flag:
        parts.append(f"플래그={flag}")
    return f"{changed}개 메일에 적용: {', '.join(parts) or '(변경 없음)'}"


# ═══════════════════════════════ F. 캘린더 ═══════════════════════════════


@mcp.tool()
@outlook_tool
def list_appointments(
    start: str = "",
    end: str = "",
    include_recurrences: bool = True,
    limit: int = 50,
) -> str:
    """기간 내 일정을 조회합니다.

    Args:
        start: 시작일. 'YYYY-MM-DD'. 생략하면 오늘.
        end: 종료일. 생략하면 start + 7일.
        include_recurrences: 반복 일정을 각 발생 건으로 펼칠지(기본 True).
        limit: 최대 표시 개수.

    Returns:
        일정 요약(시작~끝/제목/장소/주최자/entry_id).
    """
    start_dt = _parse_dt(start) or datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    end_dt = _parse_dt(end, end_of_day=True) or (start_dt + timedelta(days=7))

    cal = _ns().GetDefaultFolder(OL_FOLDER["calendar"])
    items = cal.Items
    if include_recurrences:
        items.Sort("[Start]")
        items.IncludeRecurrences = True
    else:
        items.Sort("[Start]")
    flt = f"[Start] >= '{_ol_time(start_dt)}' AND [Start] <= '{_ol_time(end_dt)}'"
    try:
        items = items.Restrict(flt)
    except pythoncom.com_error:
        pass

    limit = max(1, min(limit, MAX_ITEMS))
    out = []
    for appt in items:
        try:
            s = _fmt_dt(appt.Start)
            e = _fmt_dt(appt.End)
        except Exception:
            continue
        subject = _clean(getattr(appt, "Subject", "")) or "(제목 없음)"
        loc = _clean(getattr(appt, "Location", ""))
        loc_str = f"  @ {loc}" if loc else ""
        organizer = _clean(getattr(appt, "Organizer", ""))
        org_str = f"  주최: {organizer}" if organizer else ""
        eid = getattr(appt, "EntryID", "")
        out.append(f"• {s} ~ {e}  {subject}{loc_str}{org_str}\n    id: {eid}")
        if len(out) >= limit:
            break
    head = f"일정: {_fmt_dt(start_dt)} ~ {_fmt_dt(end_dt)}  |  {len(out)}건"
    if not out:
        return head + "\n\n해당 기간에 일정이 없습니다."
    return f"{head}\n\n" + "\n".join(out)


@mcp.tool()
@outlook_tool
def read_appointment(entry_id: str, max_chars: int = MAX_BODY) -> str:
    """일정 하나의 상세를 읽습니다(본문·참석자·반복 여부 포함).

    Args:
        entry_id: 일정 id.
        max_chars: 본문 최대 글자 수.

    Returns:
        제목/시간/장소/주최자/참석자/본문.
    """
    appt = _get_item(entry_id)
    out = [f"제목: {_clean(getattr(appt, 'Subject', '')) or '(제목 없음)'}"]
    out.append(f"시작: {_fmt_dt(getattr(appt, 'Start', None))}")
    out.append(f"종료: {_fmt_dt(getattr(appt, 'End', None))}")
    loc = _clean(getattr(appt, "Location", ""))
    if loc:
        out.append(f"장소: {loc}")
    org = _clean(getattr(appt, "Organizer", ""))
    if org:
        out.append(f"주최자: {org}")
    req = _recipients(appt, (1,))
    if req:
        out.append(f"참석자: {', '.join(req)}")
    try:
        if appt.IsRecurring:
            out.append("반복: 예")
    except Exception:
        pass
    out.append("")
    out.append(_truncate(_clean(getattr(appt, "Body", "")), max(200, max_chars)))
    return "\n".join(out)


@mcp.tool()
@outlook_tool
def get_free_busy(recipient: str, start: str = "", days: int = 7) -> str:
    """참석자의 여유/바쁨(Free/Busy) 정보를 조회합니다.

    회의를 잡기 전에 상대가 언제 비어 있는지 볼 때 사용합니다. 회사 정책상 조회가
    제한되면 확인만 되고 결과가 비어 나올 수 있습니다.

    Args:
        recipient: 대상 이름 또는 이메일 주소.
        start: 조회 시작일 'YYYY-MM-DD'. 생략하면 오늘.
        days: 조회할 일수(기본 7).

    Returns:
        30분 단위 상태 문자열(0=여유,1=미정,2=바쁨,3=부재중)과 해석.
    """
    ns = _ns()
    r = ns.CreateRecipient(recipient)
    r.Resolve()
    if not r.Resolved:
        return f"'{recipient}'을(를) 확인하지 못했습니다."
    start_dt = _parse_dt(start) or datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    try:
        # FreeBusy(Start, MinPerChar, CompleteFormat)
        raw = r.FreeBusy(start_dt, 30, True)
    except pythoncom.com_error as e:
        return f"여유/바쁨 정보를 가져오지 못했습니다(정책 제한일 수 있음): {_com_message(e)}"
    legend = {"0": "여유", "1": "미정", "2": "바쁨", "3": "부재중"}
    # 하루 = 48칸(30분). days만큼만 표시.
    per_day = 48
    out = [f"{r.Name}  |  {days}일간 여유/바쁨 (30분 단위)"]
    for d in range(days):
        chunk = raw[d * per_day:(d + 1) * per_day]
        if not chunk:
            break
        day = (start_dt + timedelta(days=d)).strftime("%m-%d")
        out.append(f"  {day}: {chunk}")
    out.append("  범례: " + ", ".join(f"{k}={v}" for k, v in legend.items()))
    return "\n".join(out)


@mcp.tool()
@outlook_tool
def create_appointment(
    subject: str,
    start: str,
    end: str = "",
    location: str = "",
    body: str = "",
    reminder_minutes: int = 15,
    all_day: bool = False,
    busy: str = "busy",
) -> str:
    """개인 일정(약속)을 만듭니다. 참석자 초대는 보내지 않습니다.

    참석자에게 초대장이 나가는 '회의'는 승인 게이팅과 함께 다음 버전에서 추가합니다.

    Args:
        subject: 제목.
        start: 시작 'YYYY-MM-DD HH:MM'(종일이면 'YYYY-MM-DD').
        end: 종료. 생략하면 start + 1시간(종일이면 하루).
        location: 장소.
        body: 메모.
        reminder_minutes: 시작 전 알림(분). 0이면 알림 없음.
        all_day: 종일 일정 여부.
        busy: 표시 상태 free/tentative/busy/oof.

    Returns:
        만든 일정의 entry_id와 요약.
    """
    start_dt = _parse_dt(start)
    if start_dt is None:
        return "start가 필요합니다 (예: 2026-07-20 14:00)."
    end_dt = _parse_dt(end) or (start_dt + timedelta(days=1 if all_day else 0, hours=0 if all_day else 1))

    appt = _app().CreateItem(OL_APPOINTMENT_ITEM)
    appt.Subject = subject
    appt.Start = start_dt
    appt.End = end_dt
    if all_day:
        appt.AllDayEvent = True
    if location:
        appt.Location = location
    if body:
        appt.Body = body
    appt.BusyStatus = OL_BUSY.get(busy.lower(), 2)
    if reminder_minutes and reminder_minutes > 0:
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes
    else:
        appt.ReminderSet = False
    appt.Save()
    return (
        "일정을 만들었습니다.\n"
        f"  {subject}  |  {_fmt_dt(start_dt)} ~ {_fmt_dt(end_dt)}\n"
        f"  id: {appt.EntryID}"
    )


# ═══════════════════════════════ G. 연락처 ═══════════════════════════════


@mcp.tool()
@outlook_tool
def list_contacts(query: str = "", folder: str = "contacts", limit: int = 50) -> str:
    """연락처 목록을 조회하거나 이름/회사/이메일로 검색합니다.

    Args:
        query: 이름/회사/이메일에 포함될 문자열(부분 일치). 생략하면 전체.
        folder: 연락처 폴더(기본 contacts).
        limit: 최대 표시 개수.

    Returns:
        연락처 요약(이름/회사/이메일/entry_id).
    """
    fld = _resolve_folder(folder)
    needle = query.lower()
    limit = max(1, min(limit, MAX_ITEMS))
    out = []
    for c in fld.Items:
        try:
            if getattr(c, "Class", None) != OL_CLASS_CONTACT:
                continue
        except Exception:
            continue
        name = _clean(getattr(c, "FullName", "")) or _clean(getattr(c, "CompanyName", ""))
        company = _clean(getattr(c, "CompanyName", ""))
        email = _clean(getattr(c, "Email1Address", ""))
        if needle and needle not in f"{name} {company} {email}".lower():
            continue
        eid = getattr(c, "EntryID", "")
        line = f"• {name or '(이름 없음)'}"
        if company:
            line += f"  |  {company}"
        if email:
            line += f"  |  {email}"
        out.append(f"{line}\n    id: {eid}")
        if len(out) >= limit:
            break
    if not out:
        return "연락처가 없습니다." + (f" (검색어: '{query}')" if query else "")
    return f"연락처 {len(out)}개:\n\n" + "\n".join(out)


@mcp.tool()
@outlook_tool
def read_contact(entry_id: str) -> str:
    """연락처 한 명의 상세 정보를 읽습니다.

    Args:
        entry_id: 연락처 id.

    Returns:
        이름/회사/직책/이메일/전화/주소 등 사용 가능한 항목.
    """
    c = _get_item(entry_id)
    fields = [
        ("이름", "FullName"), ("회사", "CompanyName"), ("직책", "JobTitle"),
        ("이메일", "Email1Address"), ("이메일2", "Email2Address"),
        ("휴대폰", "MobileTelephoneNumber"), ("회사전화", "BusinessTelephoneNumber"),
        ("부서", "Department"), ("주소", "BusinessAddress"),
    ]
    out = []
    for label, attr in fields:
        try:
            v = _clean(getattr(c, attr, ""))
        except Exception:
            v = ""
        if v:
            out.append(f"  {label}: {v}")
    notes = _clean(getattr(c, "Body", ""))
    if notes:
        out.append(f"  메모: {_truncate(notes, 1000)}")
    if not out:
        return "표시할 연락처 정보가 없습니다."
    return "연락처:\n" + "\n".join(out)


@mcp.tool()
@outlook_tool
def create_contact(
    full_name: str,
    email: str = "",
    company: str = "",
    job_title: str = "",
    mobile: str = "",
    business_phone: str = "",
    notes: str = "",
) -> str:
    """연락처를 새로 만듭니다.

    Args:
        full_name: 전체 이름.
        email: 대표 이메일.
        company: 회사.
        job_title: 직책.
        mobile: 휴대폰.
        business_phone: 회사 전화.
        notes: 메모.

    Returns:
        만든 연락처의 entry_id.
    """
    c = _app().CreateItem(OL_CONTACT_ITEM)
    c.FullName = full_name
    if email:
        c.Email1Address = email
    if company:
        c.CompanyName = company
    if job_title:
        c.JobTitle = job_title
    if mobile:
        c.MobileTelephoneNumber = mobile
    if business_phone:
        c.BusinessTelephoneNumber = business_phone
    if notes:
        c.Body = notes
    c.Save()
    return f"연락처를 만들었습니다: {full_name}\n  id: {c.EntryID}"


# ═══════════════════════════════ H. 작업 ═══════════════════════════════


@mcp.tool()
@outlook_tool
def list_tasks(include_completed: bool = False, limit: int = 50) -> str:
    """작업(To-Do) 목록을 조회합니다.

    Args:
        include_completed: 완료된 작업도 포함할지(기본 False).
        limit: 최대 표시 개수.

    Returns:
        작업 요약(제목/기한/상태/완료율/entry_id).
    """
    fld = _ns().GetDefaultFolder(OL_FOLDER["tasks"])
    limit = max(1, min(limit, MAX_ITEMS))
    out = []
    for t in fld.Items:
        try:
            if getattr(t, "Class", None) != OL_CLASS_TASK:
                continue
            complete = bool(getattr(t, "Complete", False))
        except Exception:
            continue
        if complete and not include_completed:
            continue
        subject = _clean(getattr(t, "Subject", "")) or "(제목 없음)"
        due = _fmt_dt(getattr(t, "DueDate", None))
        pct = getattr(t, "PercentComplete", 0)
        status = "완료" if complete else (f"{pct}%" if pct else "미시작")
        due_str = f"  기한 {due}" if due and "4501" not in due else ""  # 미설정 기한 회피
        eid = getattr(t, "EntryID", "")
        out.append(f"• [{status}] {subject}{due_str}\n    id: {eid}")
        if len(out) >= limit:
            break
    if not out:
        return "작업이 없습니다."
    return f"작업 {len(out)}개:\n\n" + "\n".join(out)


@mcp.tool()
@outlook_tool
def create_task(
    subject: str,
    due: str = "",
    body: str = "",
    reminder: str = "",
) -> str:
    """작업(To-Do)을 새로 만듭니다.

    Args:
        subject: 제목.
        due: 기한 'YYYY-MM-DD'.
        body: 메모.
        reminder: 알림 시각 'YYYY-MM-DD HH:MM'.

    Returns:
        만든 작업의 entry_id.
    """
    t = _app().CreateItem(OL_TASK_ITEM)
    t.Subject = subject
    due_dt = _parse_dt(due)
    if due_dt:
        t.DueDate = due_dt
    if body:
        t.Body = body
    rem_dt = _parse_dt(reminder)
    if rem_dt:
        t.ReminderSet = True
        t.ReminderTime = rem_dt
    t.Save()
    return f"작업을 만들었습니다: {subject}\n  id: {t.EntryID}"


@mcp.tool()
@outlook_tool
def complete_task(entry_id: str) -> str:
    """작업을 완료 처리합니다.

    Args:
        entry_id: 작업 id.

    Returns:
        처리 결과.
    """
    t = _get_item(entry_id)
    try:
        t.MarkComplete()
    except Exception:
        t.Status = OL_TASK_COMPLETE
        t.PercentComplete = 100
    t.Save()
    return f"작업을 완료 처리했습니다: {_clean(getattr(t, 'Subject', ''))}"


# ════════════ E. 🔴 발송·이동·삭제·회의 (되돌리기 어려움 — confirm 게이팅) ════════════
#
# 이 구역의 도구는 밖으로 나가거나 되돌리기 어려운 동작을 합니다. 모두 confirm=True를
# 받아야 실행되며, confirm 없이 부르면 프리뷰(누구에게/무슨 제목/무슨 동작)만 돌려줍니다.
# 클라이언트는 이 이름들을 HumanInTheLoopMiddleware의 INTERRUPT_ON에 올려 두면 좋습니다:
#   send_email, send_draft, respond_message, move_message, delete_message,
#   create_meeting, respond_meeting


@mcp.tool()
@outlook_tool
def send_email(
    to: str,
    subject: str = "",
    body: str = "",
    cc: str = "",
    bcc: str = "",
    html: bool = False,
    attachments: str = "",
    importance: str = "",
    confirm: bool = False,
) -> str:
    """🔴 새 메일을 작성해 즉시 발송합니다. (confirm=True 필요)

    confirm 없이 호출하면 받는사람/제목/본문 앞부분 프리뷰만 돌려주고 보내지 않습니다.
    초안만 만들어 두고 싶으면 create_draft를 쓰세요.

    Args:
        to/cc/bcc: 수신자(세미콜론/쉼표로 여러 명). 이름 또는 이메일 주소.
        subject: 제목.
        body: 본문.
        html: body를 HTML로 다룰지.
        attachments: 첨부할 파일 경로(세미콜론/쉼표).
        importance: low/normal/high.
        confirm: 실제 발송하려면 True. 없으면 프리뷰만.

    Returns:
        발송 결과, 또는 (confirm 없을 때) 발송 프리뷰.
    """
    if not _split(to) and not _split(cc) and not _split(bcc):
        return "받는사람(to/cc/bcc)이 비어 있습니다."
    mail = _app().CreateItem(OL_MAIL_ITEM)
    _fill_mail(mail, to, cc, bcc, subject, body, html, attachments, importance)

    if not confirm:
        details = [f"받는사람: {mail.To or '(비어 있음)'}"]
        if mail.CC:
            details.append(f"참조: {mail.CC}")
        details.append(f"제목: {mail.Subject or '(비어 있음)'}")
        atts = _split(attachments)
        if atts:
            details.append(f"첨부: {', '.join(os.path.basename(a) for a in atts)}")
        details.append(f"본문 앞부분: {_truncate(_clean(body), 300)}")
        return _preview("새 메일 발송", details, "(send_email ... confirm=true)")

    mail.Send()
    return (
        "발송 완료.\n"
        f"  받는사람: {mail.To or '(비어 있음)'}\n"
        f"  제목: {mail.Subject or '(제목 없음)'}"
    )


@mcp.tool()
@outlook_tool
def send_draft(entry_id: str, confirm: bool = False) -> str:
    """🔴 이미 만들어 둔 초안을 발송합니다. (confirm=True 필요)

    create_draft나 respond_message로 만든 초안의 entry_id를 받아 그대로 보냅니다.
    confirm 없이 부르면 받는사람/제목 프리뷰만 돌려줍니다.

    Args:
        entry_id: 발송할 초안 id.
        confirm: 실제 발송하려면 True.

    Returns:
        발송 결과, 또는 (confirm 없을 때) 발송 프리뷰.
    """
    item = _get_item(entry_id)
    if not _is_mail(item):
        return "이 항목은 메일이 아니라 발송할 수 없습니다."
    to = item.To or "; ".join(_recipients(item, (1,))) or "(비어 있음)"
    if not confirm:
        details = [f"받는사람: {to}", f"제목: {item.Subject or '(제목 없음)'}"]
        cc = getattr(item, "CC", "")
        if cc:
            details.append(f"참조: {cc}")
        details.append(f"본문 앞부분: {_truncate(_clean(getattr(item, 'Body', '')), 300)}")
        return _preview("초안 발송", details, "(send_draft ... confirm=true)")
    item.Send()
    return f"발송 완료.\n  받는사람: {to}\n  제목: {item.Subject or '(제목 없음)'}"


@mcp.tool()
@outlook_tool
def move_message(entry_ids: str, folder: str, confirm: bool = False) -> str:
    """🔴 메일을 다른 폴더로 이동합니다. (confirm=True 필요)

    Args:
        entry_ids: 옮길 메일 id들(세미콜론/쉼표로 여러 개).
        folder: 대상 폴더 별칭(inbox/sent/deleted 등) 또는 '계정\\폴더' 경로.
        confirm: 실제 이동하려면 True. 없으면 프리뷰만.

    Returns:
        이동 결과, 또는 (confirm 없을 때) 프리뷰.
    """
    ids = _split(entry_ids)
    if not ids:
        return "entry_ids가 비어 있습니다."
    dest = _resolve_folder(folder)
    if not confirm:
        subjects = []
        for eid in ids[:10]:
            try:
                subjects.append(_clean(getattr(_ns().GetItemFromID(eid), "Subject", "")) or "(제목 없음)")
            except Exception:
                subjects.append("(찾을 수 없는 항목)")
        more = f" 외 {len(ids) - 10}통" if len(ids) > 10 else ""
        return _preview(
            f"메일 {len(ids)}통을 '{dest.Name}' 폴더로 이동",
            [f"- {s}" for s in subjects] + ([more] if more else []),
            "(move_message ... confirm=true)",
        )
    moved = 0
    for eid in ids:
        try:
            _ns().GetItemFromID(eid).Move(dest)
            moved += 1
        except Exception:
            continue
    return f"{moved}/{len(ids)}통을 '{dest.Name}' 폴더로 이동했습니다."


@mcp.tool()
@outlook_tool
def delete_message(
    entry_ids: str, permanent: bool = False, confirm: bool = False
) -> str:
    """🔴 메일을 삭제합니다. 기본은 지운편지함으로, permanent=True면 영구 삭제. (confirm=True 필요)

    Args:
        entry_ids: 삭제할 메일 id들(세미콜론/쉼표로 여러 개).
        permanent: True면 지운편지함을 거치지 않고 영구 삭제(복구 어려움).
        confirm: 실제 삭제하려면 True. 없으면 프리뷰만.

    Returns:
        삭제 결과, 또는 (confirm 없을 때) 프리뷰.
    """
    ids = _split(entry_ids)
    if not ids:
        return "entry_ids가 비어 있습니다."
    kind = "영구 삭제" if permanent else "지운편지함으로 삭제"
    if not confirm:
        subjects = []
        for eid in ids[:10]:
            try:
                subjects.append(_clean(getattr(_ns().GetItemFromID(eid), "Subject", "")) or "(제목 없음)")
            except Exception:
                subjects.append("(찾을 수 없는 항목)")
        more = f" 외 {len(ids) - 10}통" if len(ids) > 10 else ""
        return _preview(
            f"메일 {len(ids)}통 {kind}",
            [f"- {s}" for s in subjects] + ([more] if more else []),
            "(delete_message ... confirm=true)",
        )
    deleted_fld = _ns().GetDefaultFolder(OL_FOLDER["deleted"]) if permanent else None
    done = 0
    for eid in ids:
        try:
            item = _ns().GetItemFromID(eid)
            item.Delete()  # 먼저 지운편지함으로
            if permanent and deleted_fld is not None:
                # 지운편지함에서 같은 항목을 찾아 한 번 더 삭제(영구)
                try:
                    for it in list(deleted_fld.Items):
                        if getattr(it, "Subject", None) == getattr(item, "Subject", None):
                            it.Delete()
                            break
                except Exception:
                    pass
            done += 1
        except Exception:
            continue
    return f"{done}/{len(ids)}통을 {kind} 처리했습니다."


@mcp.tool()
@outlook_tool
def create_meeting(
    subject: str,
    start: str,
    attendees: str,
    end: str = "",
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
    reminder_minutes: int = 15,
    busy: str = "busy",
    confirm: bool = False,
) -> str:
    """🔴 참석자에게 초대장을 보내는 '회의'를 만듭니다. (confirm=True 필요)

    참석자에게 실제로 초대 메일이 나가므로 🔴 동작입니다. 초대 없이 내 캘린더에만
    넣으려면 create_appointment를 쓰세요. confirm 없이 부르면 프리뷰만 돌려줍니다.

    Args:
        subject: 제목.
        start: 시작 'YYYY-MM-DD HH:MM'.
        attendees: 필수 참석자(세미콜론/쉼표로 여러 명). 이름 또는 이메일.
        end: 종료. 생략하면 start + 1시간.
        location: 장소.
        body: 본문/안건.
        optional_attendees: 선택 참석자.
        reminder_minutes: 시작 전 알림(분). 0이면 없음.
        busy: 표시 상태 free/tentative/busy/oof.
        confirm: 실제 초대 발송하려면 True. 없으면 프리뷰만.

    Returns:
        발송 결과와 entry_id, 또는 (confirm 없을 때) 프리뷰.
    """
    start_dt = _parse_dt(start)
    if start_dt is None:
        return "start가 필요합니다 (예: 2026-07-20 14:00)."
    end_dt = _parse_dt(end) or (start_dt + timedelta(hours=1))
    req = _split(attendees)
    opt = _split(optional_attendees)
    if not req and not opt:
        return "참석자(attendees)가 비어 있습니다."

    if not confirm:
        details = [
            f"제목: {subject}",
            f"일시: {_fmt_dt(start_dt)} ~ {_fmt_dt(end_dt)}",
            f"필수 참석자: {', '.join(req) or '(없음)'}",
        ]
        if opt:
            details.append(f"선택 참석자: {', '.join(opt)}")
        if location:
            details.append(f"장소: {location}")
        return _preview("회의 초대 발송", details, "(create_meeting ... confirm=true)")

    appt = _app().CreateItem(OL_APPOINTMENT_ITEM)
    appt.Subject = subject
    appt.Start = start_dt
    appt.End = end_dt
    appt.MeetingStatus = OL_MEETING  # 회의로 전환 → Send가 초대장 발송
    if location:
        appt.Location = location
    if body:
        appt.Body = body
    appt.BusyStatus = OL_BUSY.get(busy.lower(), 2)
    if reminder_minutes and reminder_minutes > 0:
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes
    else:
        appt.ReminderSet = False
    for name in req:
        r = appt.Recipients.Add(name)
        r.Type = 1  # olRequired
    for name in opt:
        r = appt.Recipients.Add(name)
        r.Type = 2  # olOptional
    try:
        appt.Recipients.ResolveAll()
    except Exception:
        pass
    appt.Send()
    return (
        "회의 초대를 발송했습니다.\n"
        f"  {subject}  |  {_fmt_dt(start_dt)} ~ {_fmt_dt(end_dt)}\n"
        f"  참석자: {', '.join(req + opt)}\n"
        f"  id: {appt.EntryID}"
    )


@mcp.tool()
@outlook_tool
def respond_meeting(
    entry_id: str,
    response: str = "accept",
    send_response: bool = True,
    confirm: bool = False,
) -> str:
    """🔴 받은 회의 요청에 수락/미정/거절로 응답합니다. (confirm=True 필요)

    주최자에게 응답 메일이 나가므로 🔴 동작입니다. confirm 없이 부르면 어떤 회의에
    어떻게 응답할지 프리뷰만 돌려줍니다.

    Args:
        entry_id: 회의 요청(또는 해당 일정) id.
        response: accept(수락) / tentative(미정) / decline(거절).
        send_response: 주최자에게 응답을 보낼지(False면 조용히 처리).
        confirm: 실제 응답하려면 True. 없으면 프리뷰만.

    Returns:
        응답 결과, 또는 (confirm 없을 때) 프리뷰.
    """
    resp = OL_MEETING_RESPONSE.get((response or "").lower())
    if resp is None:
        return "response는 accept / tentative / decline 중 하나여야 합니다."
    item = _get_item(entry_id)

    # 받은편지함의 회의 요청(MeetingItem)이면 연결된 일정을 가져온다.
    appt = item
    try:
        if getattr(item, "Class", None) == OL_CLASS_MEETING_REQUEST:
            appt = item.GetAssociatedAppointment(False)
    except Exception:
        appt = item

    subject = _clean(getattr(appt, "Subject", "")) or "(제목 없음)"
    when = f"{_fmt_dt(getattr(appt, 'Start', None))} ~ {_fmt_dt(getattr(appt, 'End', None))}"
    label = {3: "수락", 2: "미정", 4: "거절"}[resp]

    if not confirm:
        return _preview(
            f"회의 '{subject}'에 '{label}'(으)로 응답",
            [f"일시: {when}", f"주최자에게 응답 발송: {'예' if send_response else '아니오'}"],
            "(respond_meeting ... confirm=true)",
        )

    try:
        reply = appt.Respond(resp, True, False)  # (Response, fNoUI, fAdditionalTextDialog)
        if send_response and reply is not None:
            reply.Send()
    except pythoncom.com_error as e:
        return f"회의 응답에 실패했습니다: {_com_message(e)}"
    return f"회의 '{subject}'에 '{label}'(으)로 응답했습니다. ({when})"


if __name__ == "__main__":
    import argparse
    import sys

    parser = argparse.ArgumentParser(description="Outlook 메일/일정/연락처/작업 MCP 서버")
    parser.add_argument(
        "--transport",
        choices=["stdio", "http", "sse"],
        default=os.getenv("OUTLOOK_MCP_TRANSPORT", "stdio"),
        help=(
            "stdio(기본): Claude 등 로컬 클라이언트가 프로세스를 직접 실행해 붙는다. "
            "http: n8n 등 네트워크 클라이언트가 URL로 접속한다. "
            "sse: 구버전 n8n MCP 노드용."
        ),
    )
    parser.add_argument("--host", default=os.getenv("OUTLOOK_MCP_HOST", "127.0.0.1"))
    parser.add_argument(
        "--port", type=int, default=int(os.getenv("OUTLOOK_MCP_PORT", "8088"))
    )
    args = parser.parse_args()

    if not COM_AVAILABLE:
        print(f"경고: pywin32를 불러올 수 없습니다 ({COM_IMPORT_ERROR}).", file=sys.stderr)
        print("서버는 실행되지만 모든 도구가 안내 메시지만 반환합니다.", file=sys.stderr)

    if args.transport in ("http", "sse"):
        # COM 특성상 반드시 사용자가 로그인한 세션에서 실행해야 Outlook 사서함이 보인다
        # (서비스나 다른 세션에서는 안 보임).
        path = "/mcp/" if args.transport == "http" else "/sse/"
        url = f"http://{args.host}:{args.port}{path}"
        print(f"Outlook MCP 서버 시작 ({args.transport}) — {url}", file=sys.stderr)
        mcp.run(transport=args.transport, host=args.host, port=args.port)
    else:
        print("Outlook MCP 서버 시작 (stdio)", file=sys.stderr)
        mcp.run(transport="stdio")
