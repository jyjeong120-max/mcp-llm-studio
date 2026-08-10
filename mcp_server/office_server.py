"""office_server.py

Word / Excel / PowerPoint 문서를 읽고, 셋 다 편집(쓰기)까지 지원하는 MCP 서버입니다.
pywin32(COM)로 설치된 Office를 직접 구동하므로, 파일 경로뿐 아니라
지금 화면에 열려 있는 문서(저장 전 편집 내용 포함)도 그대로 읽습니다.

STDIO 트랜스포트라, 클라이언트가 필요한 시점에 이 스크립트를 직접 실행합니다.
streamable-http 서버들과 달리 별도 터미널에서 실행해 둘 필요가 없습니다.

Windows + Office 설치가 전제입니다. 둘 중 하나라도 없으면 서버는 그대로 실행되고,
각 도구가 실패 사유를 담은 안내 메시지를 반환합니다.

Word/Excel/PowerPoint 모두 읽기+쓰기 도구가 있습니다.
outlook_server와 같은 3티어 안전 등급을 따릅니다:
    🟢 읽기 — 모든 read_*/describe_*/find_*/inspect_*/list_slide_shapes/get_shape_size
       도구 (문서를 변경하지 않음)
    🟡 메모리 수정(저장 안 함) — Excel write_excel_cell/write_excel_range,
       Word replace_in_word/insert_word_text, PowerPoint set_powerpoint_text/
       replace_in_powerpoint/set_shape_size/set_shape_position/format_shape/
       add_text_box/add_shape/add_slide/delete_shape. **사용자 세션에 열려 있는**
       문서만 수정하고 디스크에는 쓰지 않는다(저장 전까지 파일 원본은 그대로). Excel의
       COM 수정은 실행 취소(Ctrl+Z) 스택에 쌓이지 않아 복구용으로 이전 값을 응답에 담아
       돌려주고, Word/PowerPoint는 저장 전이라 Ctrl+Z나 '저장 없이 닫기'로 되돌릴 수 있다.
    🔴 디스크 기록 — save_workbook(Excel)/save_document(Word)/save_presentation(PPT)
       덮어쓰기 저장. confirm=True 없이는 실행되지 않고 무엇을 할지 프리뷰만 돌려준다
       (_preview — outlook/catia와 같은 게이트).

path 인자 규칙 (모든 문서 도구 공통)
    path=None  -> 해당 앱에서 지금 활성화된 문서를 읽습니다.
    path 지정  -> 이미 열려 있으면 그 세션을 그대로 읽고,
                  아니면 백그라운드에서 읽기 전용으로 열었다가 닫습니다.

password 인자 규칙
    열기 암호가 걸린 문서는 password로 암호를 넘깁니다. 이미 열려 있는 문서는
    사용자가 암호를 풀어 둔 상태이므로 필요하지 않습니다.
    암호가 틀리거나 빠지면 대화상자 대신 오류 메시지를 돌려줍니다(NO_PASSWORD 참고).
"""

from __future__ import annotations

import os
import threading
from contextlib import contextmanager
from datetime import datetime

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
    name="office",
    instructions=(
        "Word/Excel/PowerPoint 문서를 읽고 수정하는 MCP 서버입니다. "
        "path를 생략하면 지금 열려 있는 활성 문서를 읽습니다. "
        "어떤 문서가 열려 있는지 모르면 list_open_documents를 먼저 호출하세요. "
        "Word 특정 섹션만 읽으려면 read_word_section을, PowerPoint 도형을 수정하기 전 "
        "번호·크기를 확인하려면 list_slide_shapes를 쓰세요. "
        "쓰기(Excel write_excel_*, Word replace_in_word/insert_word_text, PowerPoint "
        "set_powerpoint_text/set_shape_size/add_shape 등)는 열려 있는 문서의 메모리만 "
        "바꾸고 저장하지 않습니다 — 디스크 저장은 save_workbook/save_document/"
        "save_presentation이 담당하며 confirm=True 없이 부르면 프리뷰만 돌려줍니다."
    ),
)

# 출력이 컨텍스트를 통째로 삼키지 않도록 하는 기본 상한. 도구 인자로 조정할 수 있다.
MAX_ROWS = 200
MAX_COLS = 50
MAX_CHARS = 20000
MAX_MATCHES = 50

# COM 상수. win32com.client.constants는 타입 라이브러리를 먼저 만들어야 해서 직접 쓴다.
MSO_FALSE = 0
MSO_TRUE = -1

# 암호가 지정되지 않았을 때 대신 넘기는 값.
# 암호를 아예 넘기지 않거나 빈 문자열을 넘기면, 보호된 문서를 열 때 Office가
# 암호 입력 대화상자를 띄운다. 화면이 없는 서버에서는 그 대화상자를 닫을 방법이
# 없어 COM 호출이 영영 멈춘다(에이전트도 같이 멈춘다). 실제 암호일 리 없는
# 문자열을 넘기면 대화상자 대신 "암호가 틀렸다"는 오류가 나므로 잡아서 처리할 수 있다.
# 두 가지 제약이 있다 — 둘 다 어기면 실제로 깨진다:
#   - NUL 문자 금지: Office가 빈 암호로 보고 다시 대화상자를 띄운다.
#   - 15자 이하: Excel의 암호 최대 길이가 15자라, 넘기면 암호 오류가 아니라
#     Open 자체가 실패해서 암호 없는 파일조차 열지 못한다.
NO_PASSWORD = "~~nopw~~"
WD_BODY_TEXT = 10  # wdOutlineLevelBodyText — 이 값 미만이면 제목 단락
WD_GOTO_PAGE = 1  # wdGoToPage
WD_GOTO_ABSOLUTE = 1  # wdGoToAbsolute
WD_STAT_PAGES = 2  # wdStatisticPages
# Word 쓰기용 상수 (win32com.client.constants를 쓸 수 없어 값으로 직접 둔다)
WD_FIND_STOP = 0  # wdFindStop — 문서 끝에서 멈춤(개수 세기용)
WD_FIND_CONTINUE = 1  # wdFindContinue — 끝까지 이어서 바꾸기
WD_REPLACE_ONE = 1  # wdReplaceOne
WD_REPLACE_ALL = 2  # wdReplaceAll
WD_COLLAPSE_START = 1  # wdCollapseStart
WD_COLLAPSE_END = 0  # wdCollapseEnd
WD_NO_PROTECTION = -1  # wdNoProtection — 이 값이 아니면 편집 제한이 걸려 있다
WD_REVISION_TYPES = {  # WdRevisionType
    0: "변경 없음",
    1: "삽입",
    2: "삭제",
    3: "속성 변경",
    4: "단락 번호 변경",
    5: "필드 표시 변경",
    6: "조정",
    7: "충돌",
    8: "스타일 변경",
    9: "대체",
    10: "단락 서식 변경",
    11: "표 속성 변경",
    12: "구역 속성 변경",
    13: "스타일 정의 변경",
    14: "이동(원위치)",
    15: "이동(새 위치)",
    16: "셀 삽입",
    17: "셀 삭제",
    18: "셀 병합",
    19: "셀 분할",
    20: "충돌 삽입",
    21: "충돌 삭제",
}

# PowerPoint 쓰기용 상수 (win32com.client.constants를 쓸 수 없어 값으로 직접 둔다)
MSO_TEXT_HORIZONTAL = 1  # msoTextOrientationHorizontal
PP_ALIGN = {"left": 1, "center": 2, "right": 3, "justify": 4}  # PpParagraphAlignment
# MsoShapeType — list_slide_shapes에서 도형 종류를 사람이 읽을 이름으로.
MSO_SHAPE_TYPES = {
    1: "자동도형", 2: "콜아웃", 3: "차트", 4: "코멘트", 5: "자유형", 6: "그룹",
    7: "임베디드개체", 8: "폼컨트롤", 9: "선", 11: "OLE개체", 12: "미디어",
    13: "그림", 14: "개체틀", 17: "텍스트상자", 19: "표", 20: "워드아트",
    21: "잉크", 24: "다이어그램", 26: "3D모델",
}
# 자주 쓰는 MsoAutoShapeType (add_shape의 friendly 이름 → 값).
# COM이 거부하는 값을 주면 com_error가 나므로, 그때는 유효 이름 목록으로 안내한다.
PPT_AUTO_SHAPES = {
    "rectangle": 1, "사각형": 1,
    "rounded_rectangle": 5, "둥근사각형": 5,
    "oval": 9, "ellipse": 9, "타원": 9, "원": 9,
    "triangle": 7, "삼각형": 7,
    "right_triangle": 8, "직각삼각형": 8,
    "diamond": 4, "마름모": 4,
    "parallelogram": 2, "평행사변형": 2,
    "trapezoid": 3, "사다리꼴": 3,
    "pentagon": 12, "오각형": 12,
    "hexagon": 10, "육각형": 10,
    "octagon": 6, "팔각형": 6,
    "cross": 11, "십자": 11,
    "can": 13, "원통": 13,
    "cube": 14, "정육면체": 14,
    "heart": 21, "하트": 21,
    "sun": 23, "해": 23,
    "moon": 24, "달": 24,
    "smiley": 17, "웃는얼굴": 17,
    "lightning": 22, "번개": 22,
    "arc": 25, "호": 25,
    "right_arrow": 33, "오른쪽화살표": 33,
    "left_arrow": 34, "왼쪽화살표": 34,
    "up_arrow": 35, "위쪽화살표": 35,
    "down_arrow": 36, "아래쪽화살표": 36,
    "left_right_arrow": 37, "좌우화살표": 37,
    "star": 92, "별": 92,
    "chevron": 52, "갈매기": 52,
}
# PpSlideLayout (add_slide의 friendly 이름 → 값).
PPT_SLIDE_LAYOUTS = {
    "blank": 12, "빈": 12,
    "title": 1, "제목": 1,
    "title_only": 11, "제목만": 11,
    "text": 2, "제목과내용": 2,
    "two_column": 3, "2단": 3,
    "object": 7, "개체": 7,
}
# 단위 → 포인트(pt) 환산. PowerPoint COM의 Left/Top/Width/Height는 포인트 단위다.
PPT_UNIT_TO_PT = {
    "pt": 1.0, "point": 1.0, "points": 1.0,
    "cm": 28.3464567, "mm": 2.83464567,
    "in": 72.0, "inch": 72.0, "inches": 72.0,
}

_APPS = {
    "word": ("Word.Application", "Documents"),
    "excel": ("Excel.Application", "Workbooks"),
    "ppt": ("PowerPoint.Application", "Presentations"),
}


class OfficeError(Exception):
    """도구가 사용자에게 그대로 돌려줄 안내 메시지를 담은 예외."""


# 지금 읽는 문서가 사용자 세션에서 온 것인지(True) 우리가 백그라운드로 연 것인지(False).
# FastMCP가 도구를 워커 스레드에서 돌리므로 스레드별로 따로 둔다.
_ctx = threading.local()


def _require_com():
    if not COM_AVAILABLE:
        raise OfficeError(
            "Office 문서를 읽지 못했습니다. pywin32를 불러올 수 없습니다"
            f"({COM_IMPORT_ERROR}). Windows에서 `pip install pywin32`로 설치한 뒤 "
            "다시 요청하세요."
        )


def office_tool(fn):
    """COM 초기화와 예외 처리를 감싸는 도구 데코레이터.

    FastMCP는 동기 도구를 워커 스레드에서 실행한다. COM은 스레드마다
    CoInitialize가 필요하므로 매 호출마다 초기화하고 해제한다.
    OfficeError는 안내 메시지로, 나머지 COM 오류는 사유를 붙여 반환한다.
    """
    from functools import wraps

    @wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            _require_com()
        except OfficeError as e:
            return str(e)

        pythoncom.CoInitialize()
        try:
            return fn(*args, **kwargs)
        except OfficeError as e:
            return str(e)
        except pythoncom.com_error as e:
            return f"Office 호출이 실패했습니다: {_com_message(e)}"
        except Exception as e:
            return f"문서를 읽지 못했습니다: {type(e).__name__}: {e}"
        finally:
            pythoncom.CoUninitialize()

    return wrapper


def _com_message(e) -> str:
    """com_error에서 사람이 읽을 만한 설명만 뽑아낸다."""
    try:
        excepinfo = e.excepinfo
        if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
            return str(excepinfo[2]).strip()
    except Exception:
        pass
    return str(e)


def _get_running_app(kind: str):
    """이미 실행 중인 Office 앱에 붙는다. 실행 중이 아니면 None."""
    prog_id = _APPS[kind][0]
    try:
        return win32com.client.GetActiveObject(prog_id)
    except (pythoncom.com_error, AttributeError):
        return None


def _same_file(a: str, b: str) -> bool:
    """열린 문서의 FullName과 사용자가 준 경로가 같은 파일인지 비교한다.

    OneDrive/SharePoint에 있는 문서의 FullName은 로컬 경로가 아니라 URL이라
    전체 경로가 어긋난다. 그때는 파일명으로 대신 맞춰 본다.
    """
    if not a or not b:
        return False
    if os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b)):
        return True
    return os.path.normcase(os.path.basename(a)) == os.path.normcase(os.path.basename(b))


def _find_in_rot(path: str):
    """실행 중인 개체 테이블(ROT)에서 해당 경로의 문서를 직접 찾는다.

    Office는 열어 둔 문서를 전체 경로를 이름으로 삼아 ROT에 등록한다.
    GetActiveObject는 앱 인스턴스를 하나만 돌려주므로, Excel처럼 인스턴스가
    여러 개 떠 있으면 사용자가 보고 있는 문서를 놓친다. ROT를 직접 훑으면
    어느 인스턴스에 속했든 문서 객체를 찾을 수 있다.
    """
    target = os.path.normcase(os.path.abspath(path))
    try:
        context = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()
        for moniker in rot:
            try:
                name = moniker.GetDisplayName(context, None)
            except pythoncom.com_error:
                continue
            # ROT에는 Office 외의 항목도 들어 있다. 경로가 정확히 같을 때만 받는다.
            if not name or os.path.normcase(os.path.abspath(name)) != target:
                continue
            try:
                obj = rot.GetObject(moniker)
                return win32com.client.Dispatch(
                    obj.QueryInterface(pythoncom.IID_IDispatch)
                )
            except pythoncom.com_error:
                continue
    except Exception:
        return None
    return None


def _find_open_doc(kind: str, path: str):
    """지정 경로의 문서가 이미 열려 있으면 그 COM 객체를 돌려준다."""
    doc = _find_in_rot(path)
    if doc is not None:
        return doc

    # ROT에 없는 경우(OneDrive/SharePoint 문서는 FullName이 URL이라 경로가 어긋난다)
    # 실행 중인 앱의 문서 목록에서 파일명으로 다시 찾아본다.
    app = _get_running_app(kind)
    if app is None:
        return None
    collection = getattr(app, _APPS[kind][1])
    try:
        for doc in collection:
            if _same_file(doc.FullName, path):
                return doc
    except pythoncom.com_error:
        return None
    return None


def _active_doc(kind: str):
    """앱에서 지금 활성화된 문서를 돌려준다. 없으면 안내와 함께 실패."""
    label = {"word": "Word", "excel": "Excel", "ppt": "PowerPoint"}[kind]
    app = _get_running_app(kind)
    if app is None:
        raise OfficeError(
            f"{label}이(가) 실행되어 있지 않아 활성 문서를 읽을 수 없습니다. "
            f"{label}에서 문서를 열거나, path 인자에 파일 경로를 지정하세요."
        )
    try:
        collection = getattr(app, _APPS[kind][1])
        if collection.Count == 0:
            raise OfficeError(
                f"{label}은 실행 중이지만 열린 문서가 없습니다. "
                "문서를 열거나 path 인자에 파일 경로를 지정하세요."
            )
        if kind == "word":
            return app.ActiveDocument
        if kind == "excel":
            return app.ActiveWorkbook
        return app.ActivePresentation
    except pythoncom.com_error as e:
        raise OfficeError(f"{label} 활성 문서를 가져오지 못했습니다: {_com_message(e)}")


def _is_password_error(msg: str) -> bool:
    """Office가 돌려준 오류가 '암호가 틀렸다'는 뜻인지 판단한다.

    오류 문구는 Office 표시 언어를 따라가므로(한국어판: '암호가 잘못되었습니다',
    영문판: 'The password you supplied is not correct') 코드가 아니라 키워드로 본다.
    못 알아보면 Office 원문을 그대로 사용자에게 보여주므로 손해는 없다.
    """
    low = msg.lower()
    return "암호" in msg or "password" in low


def _open_background(kind: str, path: str, password: str | None):
    """사용자 세션과 분리된 백그라운드 인스턴스에서 읽기 전용으로 연다.

    DispatchEx로 별도 프로세스를 띄우므로, 읽고 닫는 동안 사용자가 보고 있는
    Office 창에는 영향을 주지 않는다.

    인자는 모두 '위치 인자'로 넘긴다. pywin32의 디스패치는 Open의 Password를
    키워드로 주면 조용히 흘려버리고, 그러면 Office가 암호 대화상자를 띄운 채
    멈춘다(암호가 맞아도 마찬가지다). 위치로 넘겨야 실제로 전달된다.
    """
    pw = password or NO_PASSWORD
    app = win32com.client.DispatchEx(_APPS[kind][0])
    try:
        if kind == "word":
            app.Visible = False
            app.DisplayAlerts = 0
            # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument
            # 뒤쪽 선택 인자(Visible 등)까지 채워 넣으면 타입 불일치로 Open이 실패한다.
            # 앱을 이미 숨겨 뒀으므로 문서 창은 어차피 뜨지 않는다.
            doc = app.Documents.Open(path, False, True, False, pw)
        elif kind == "excel":
            app.Visible = False
            app.DisplayAlerts = False
            # Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword,
            # IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify,
            # Converter, AddToMru
            # UpdateLinks=0: 외부 링크를 갱신하려 들면서 멈추는 걸 막는다.
            # IgnoreReadOnlyRecommended=True: '읽기 전용 권장' 안내창을 막는다.
            doc = app.Workbooks.Open(
                path, 0, True, None, pw, pw, True,
                None, None, None, None, None, False,
            )
        else:
            # PowerPoint는 Application.Visible=False를 거부하므로 WithWindow=False로 연다.
            # Open에는 암호 인자가 아예 없다. 파일명 뒤에 ::암호:: 를 붙이는 것이
            # PowerPoint가 지원하는 유일한 방법이며, 암호가 없는 파일에 붙여도 무해하다.
            doc = app.Presentations.Open(
                f"{path}::{pw}::", MSO_TRUE, MSO_FALSE, MSO_FALSE
            )
    except pythoncom.com_error as e:
        try:
            app.Quit()
        except Exception:
            pass
        msg = _com_message(e)
        name = os.path.basename(path)
        if _is_password_error(msg):
            if password:
                raise OfficeError(
                    f"'{name}'의 암호가 맞지 않습니다. password 인자를 확인하세요."
                )
            raise OfficeError(
                f"'{name}'은(는) 암호로 보호된 문서입니다. "
                "password 인자에 열기 암호를 지정하세요."
            )
        raise OfficeError(f"'{path}' 파일을 열지 못했습니다: {msg}")
    return app, doc


@contextmanager
def _document(kind: str, path: str = "", password: str = ""):
    """문서를 확보해서 넘겨주고, 우리가 연 것만 정리한다.

    사용자가 이미 열어 둔 문서는 절대 닫지 않는다 — 편집 중인 창이
    도구 호출 때문에 사라지면 안 된다.

    password는 백그라운드로 열 때만 쓰인다. 이미 열려 있는 문서는 사용자가
    암호를 풀어 둔 상태이므로 다시 필요하지 않다.

    path가 비어 있으면(None 또는 "") 활성 문서를 읽는다 — 스키마 호환을 위해
    선택 인자를 str="" 로 두므로, 클라이언트가 빈 문자열을 보내도 활성 문서로 본다.
    """
    label = {"word": "Word", "excel": "Excel", "ppt": "PowerPoint"}[kind]

    if not path:
        _ctx.user_session = True
        yield _active_doc(kind)
        return

    path = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(path):
        opened = _find_open_doc(kind, path)
        if opened is None:
            raise OfficeError(
                f"'{path}' 경로에 파일이 없고, {label}에도 열려 있지 않습니다. "
                "경로를 확인하세요."
            )
        _ctx.user_session = True
        yield opened
        return

    already_open = _find_open_doc(kind, path)
    if already_open is not None:
        # 사용자 세션에 열려 있는 문서 — 저장 전 편집 내용까지 그대로 읽는다.
        _ctx.user_session = True
        yield already_open
        return

    _ctx.user_session = False
    app, doc = _open_background(kind, path, password)
    try:
        yield doc
    finally:
        try:
            if kind == "excel":
                doc.Close(SaveChanges=False)
            elif kind == "word":
                doc.Close(SaveChanges=0)
            else:
                doc.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass


# ─────────────────────────────── 포맷 헬퍼 ───────────────────────────────


def _fmt_value(v) -> str:
    """셀/필드 값 하나를 사람이 읽을 문자열로 만든다."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        if (v.hour, v.minute, v.second) == (0, 0, 0):
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _md_escape(s: str) -> str:
    """마크다운 표 셀 안에서 구분자와 줄바꿈이 표를 깨뜨리지 않게 한다."""
    return s.replace("|", "\\|").replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")


def _address(rng) -> str:
    """Range의 주소를 'A1:D20' 형태 문자열로 얻는다.

    pywin32의 동적 디스패치는 인자가 모두 선택적인 Address를 메서드가 아니라
    속성으로 노출해서, 접근하는 순간 이미 '$A$1:$D$20' 문자열을 돌려준다.
    makepy 캐시가 있는 환경에서는 호출 가능한 메서드가 된다 — 양쪽을 모두 받는다.
    """
    a = rng.Address
    if callable(a):
        a = a(0, 0)
    return str(a).replace("$", "")


def _col_letter(n: int) -> str:
    """1 -> A, 27 -> AA. 셀 주소를 COM 호출 없이 계산한다."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _as_grid(value):
    """Range.Value를 항상 2차원 튜플로 정규화한다.

    Excel은 여러 셀이면 튜플의 튜플, 단일 셀이면 스칼라를 준다.
    """
    if value is None:
        return [[None]]
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return [list(row) for row in value]
        return [list(value)]
    return [[value]]


def _grid_to_markdown(grid, start_row: int, start_col: int, max_rows: int, max_cols: int) -> str:
    """2차원 값 배열을 행/열 머리표가 붙은 마크다운 표로 만든다.

    시트 좌표(1행, A열)를 그대로 머리표에 넣어야, 에이전트가 읽은 값을
    다시 셀 주소로 지목할 수 있다.
    """
    total_rows, total_cols = len(grid), max(len(r) for r in grid)
    rows = grid[:max_rows]
    cols = min(total_cols, max_cols)

    header = ["", *[_col_letter(start_col + c) for c in range(cols)]]
    lines = ["| " + " | ".join(header) + " |"]
    lines.append("|" + "|".join(["---"] * len(header)) + "|")
    for i, row in enumerate(rows):
        cells = [_md_escape(_fmt_value(v)) for v in row[:cols]]
        lines.append("| " + " | ".join([str(start_row + i), *cells]) + " |")

    out = "\n".join(lines)
    notes = []
    if total_rows > max_rows:
        notes.append(f"행 {total_rows}개 중 {max_rows}개만 표시 (max_rows로 조정)")
    if total_cols > max_cols:
        notes.append(f"열 {total_cols}개 중 {cols}개만 표시 (max_cols로 조정)")
    if notes:
        out += "\n\n(" + " / ".join(notes) + ")"
    return out


def _truncate(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + f"\n\n…(전체 {len(text):,}자 중 {limit:,}자만 표시)"


def _clean(text: str) -> str:
    """Word/PPT 텍스트의 COM 특수문자를 일반 문자로 바꾼다."""
    if not text:
        return ""
    return (
        text.replace("\x07", " ")  # 셀/행 종료 표시
        .replace("\x0b", "\n")  # 수동 줄바꿈
        .replace("\r", "\n")
        .strip()
    )


def _parse_index_range(spec: str | None, total: int, label: str) -> list[int]:
    """'3', '2-5', '1,4,7' 형태를 1-based 인덱스 목록으로 바꾼다."""
    if not spec:
        return list(range(1, total + 1))
    out: list[int] = []
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        try:
            if "-" in part:
                a, b = part.split("-", 1)
                out.extend(range(int(a), int(b) + 1))
            else:
                out.append(int(part))
        except ValueError:
            raise OfficeError(f"{label} 범위 형식이 잘못되었습니다: '{spec}' (예: '3', '2-5', '1,4,7')")
    picked = [i for i in out if 1 <= i <= total]
    if not picked:
        raise OfficeError(f"'{spec}'에 해당하는 {label}이(가) 없습니다. 전체 {total}개입니다.")
    return picked


def _resolve_sheet(wb, sheet: str | int | None):
    """시트를 이름 또는 번호로 찾는다. 생략하면 활성 시트."""
    if sheet is None or sheet == "":
        return wb.ActiveSheet
    try:
        return wb.Worksheets(int(sheet) if str(sheet).isdigit() else sheet)
    except pythoncom.com_error:
        names = [ws.Name for ws in wb.Worksheets]
        raise OfficeError(
            f"'{sheet}' 시트를 찾을 수 없습니다. 이 통합문서의 시트: {', '.join(names)}"
        )


def _resolve_range(ws, cell_range: str | None):
    """범위 문자열을 Range로 바꾼다. 생략하면 데이터가 있는 영역 전체."""
    if not cell_range:
        return ws.UsedRange
    try:
        return ws.Range(cell_range)
    except pythoncom.com_error:
        raise OfficeError(
            f"'{cell_range}'는 올바른 범위가 아닙니다. 'A1:D20' 또는 'B5' 형식으로 지정하세요."
        )


def _doc_label(doc, path: str | None) -> str:
    """어떤 문서를 읽었는지 결과 머리에 적어 준다 (활성 문서일 때 특히 중요)."""
    try:
        name = doc.Name
    except Exception:
        return path or "(알 수 없는 문서)"

    # Saved는 사용자 세션 문서일 때만 의미가 있다. 백그라운드로 연 문서는
    # 페이지 수 조회 같은 읽기 동작만으로도 Word가 재페이지네이션을 하며
    # Saved를 False로 바꿔 버려서, 그대로 믿으면 "편집 중"이라고 오인한다.
    saved = ""
    if getattr(_ctx, "user_session", False):
        try:
            if not doc.Saved:
                saved = " *저장 안 됨(편집 중)"
        except Exception:
            pass
    origin = "" if path else " [활성 문서]"
    return f"{name}{origin}{saved}"


# ─────────────────────────────── 공통 도구 ───────────────────────────────


@mcp.tool()
@office_tool
def list_open_documents() -> str:
    """지금 Office에 열려 있는 Word/Excel/PowerPoint 문서 목록을 조회합니다.

    사용자가 "지금 보고 있는 문서", "열어둔 엑셀"처럼 파일 경로 없이 지칭할 때
    가장 먼저 호출합니다. 각 문서의 전체 경로와 저장 여부를 알려주므로,
    이후 도구에 넘길 path를 여기서 얻을 수 있습니다.

    Returns:
        앱별 열린 문서 목록 (이름, 전체 경로, 저장 여부, 활성 문서 표시).
        실행 중인 Office 앱이 없으면 그 사실을 알리는 메시지.
    """
    lines = []
    for kind, label in (("word", "Word"), ("excel", "Excel"), ("ppt", "PowerPoint")):
        app = _get_running_app(kind)
        if app is None:
            lines.append(f"[{label}] 실행 중이 아닙니다.")
            continue
        try:
            collection = getattr(app, _APPS[kind][1])
            if collection.Count == 0:
                lines.append(f"[{label}] 실행 중이지만 열린 문서가 없습니다.")
                continue
            try:
                active_name = {
                    "word": lambda: app.ActiveDocument.Name,
                    "excel": lambda: app.ActiveWorkbook.Name,
                    "ppt": lambda: app.ActivePresentation.Name,
                }[kind]()
            except Exception:
                active_name = None

            lines.append(f"[{label}] {collection.Count}개")
            for doc in collection:
                name = doc.Name
                try:
                    full = doc.FullName
                except Exception:
                    full = name
                # 한 번도 저장하지 않은 새 문서는 FullName이 파일명과 같다.
                loc = full if full != name else "(저장되지 않은 새 문서)"
                marks = []
                if name == active_name:
                    marks.append("활성")
                try:
                    if not doc.Saved:
                        marks.append("저장 안 됨")
                except Exception:
                    pass
                suffix = f"  <{', '.join(marks)}>" if marks else ""
                lines.append(f"  - {name}{suffix}\n      {loc}")
        except pythoncom.com_error as e:
            lines.append(f"[{label}] 목록을 가져오지 못했습니다: {_com_message(e)}")

    return "\n".join(lines)


@mcp.tool()
@office_tool
def read_document_properties(
    kind: str, path: str = "", password: str = ""
) -> str:
    """문서의 메타데이터(작성자, 수정일, 페이지/단어 수 등)를 조회합니다.

    "이 문서 누가 마지막으로 고쳤어?", "언제 만든 거야?" 같은 질문에 사용합니다.
    Word/Excel/PowerPoint 모두 지원합니다.

    Args:
        kind: 문서 종류. 'word', 'excel', 'ppt' 중 하나.
        path: 문서 파일 경로. 생략하면 해당 앱의 활성 문서를 읽습니다.
        password: 문서에 열기 암호가 걸려 있을 때 지정합니다.
            이미 열려 있는 문서를 읽을 때는 필요 없습니다.

    Returns:
        제목/작성자/최종 수정자/생성·수정 일시/회사 등 사용 가능한 속성 목록.
    """
    kind = (kind or "").strip().lower()
    if kind not in _APPS:
        return "kind는 'word', 'excel', 'ppt' 중 하나여야 합니다."

    fields = [
        "Title", "Subject", "Author", "Last author", "Manager", "Company",
        "Category", "Keywords", "Comments", "Revision number",
        "Creation date", "Last save time", "Last print date",
        "Number of pages", "Number of words", "Number of characters",
        "Number of slides", "Template",
    ]
    with _document(kind, path, password) as doc:
        out = [f"문서: {_doc_label(doc, path)}"]
        try:
            out.append(f"경로: {doc.FullName}")
        except Exception:
            pass
        out.append("")
        for name in fields:
            try:
                value = doc.BuiltinDocumentProperties(name).Value
            except Exception:
                continue  # 이 문서 종류에 없는 속성은 조용히 건너뛴다
            text = _fmt_value(value)
            if text:
                out.append(f"  {name}: {text}")
        return "\n".join(out)


# ─────────────────────────────── Excel 도구 ───────────────────────────────


@mcp.tool()
@office_tool
def describe_excel_workbook(path: str = "", password: str = "") -> str:
    """Excel 통합문서의 전체 구조를 개요로 조회합니다.

    시트마다 데이터가 어디에 얼마나 있는지, 차트/피벗테이블/표가 있는지 알려줍니다.
    큰 파일을 통째로 읽기 전에 먼저 호출해서 어느 시트의 어느 범위를 읽을지
    정하는 용도입니다.

    Args:
        path: .xlsx/.xlsm/.xls 파일 경로. 생략하면 Excel의 활성 통합문서.
        password: 통합문서에 열기 암호가 걸려 있을 때 지정합니다.
            이미 열려 있는 통합문서를 읽을 때는 필요 없습니다.

    Returns:
        시트별 사용 범위·행열 수·차트/피벗/표 개수, 명명된 범위 개수.
    """
    with _document("excel", path, password) as wb:
        out = [f"통합문서: {_doc_label(wb, path)}"]
        try:
            out.append(f"경로: {wb.FullName}")
        except Exception:
            pass
        out.append(f"시트 {wb.Worksheets.Count}개\n")

        for ws in wb.Worksheets:
            used = ws.UsedRange
            try:
                addr = _address(used)
                rows, cols = used.Rows.Count, used.Columns.Count
                empty = used.Cells.Count == 1 and used.Value is None
            except pythoncom.com_error:
                out.append(f"  [{ws.Name}] 사용 범위를 읽지 못했습니다.")
                continue

            visible = "" if ws.Visible == -1 else "  (숨김)"
            if empty:
                out.append(f"  [{ws.Name}]{visible} 빈 시트")
                continue

            extras = []
            for label, getter in (
                ("차트", lambda: ws.ChartObjects().Count),
                ("피벗테이블", lambda: ws.PivotTables().Count),
                ("표", lambda: ws.ListObjects.Count),
            ):
                try:
                    n = getter()
                    if n:
                        extras.append(f"{label} {n}개")
                except Exception:
                    pass
            extra = f"  |  {', '.join(extras)}" if extras else ""
            out.append(f"  [{ws.Name}]{visible} {addr}  ({rows}행 × {cols}열){extra}")

        try:
            if wb.Names.Count:
                out.append(f"\n명명된 범위 {wb.Names.Count}개 (read_excel_named_ranges로 조회)")
        except Exception:
            pass
        return "\n".join(out)


@mcp.tool()
@office_tool
def read_excel_range(
    path: str = "",
    sheet: str = "",
    cell_range: str = "",
    max_rows: int = MAX_ROWS,
    max_cols: int = MAX_COLS,
    password: str = "",
) -> str:
    """Excel 시트의 셀 값을 읽어 표로 반환합니다.

    수식이 아니라 계산된 값을 읽습니다. 날짜 셀은 날짜로 변환해서 보여줍니다.
    범위를 생략하면 데이터가 있는 영역 전체를 읽으므로, 큰 시트라면
    describe_excel_workbook으로 크기를 먼저 확인하세요.

    Args:
        path: 파일 경로. 생략하면 Excel의 활성 통합문서(저장 전 편집 내용 포함).
        sheet: 시트 이름 또는 번호(1부터). 생략하면 활성 시트.
        cell_range: 'A1:D20' 같은 범위. 생략하면 사용 중인 영역 전체.
        max_rows: 표시할 최대 행 수. 기본 200.
        max_cols: 표시할 최대 열 수. 기본 50.
        password: 통합문서에 열기 암호가 걸려 있을 때 지정합니다.
            이미 열려 있는 통합문서를 읽을 때는 필요 없습니다.

    Returns:
        시트 좌표(행 번호·열 문자)가 붙은 마크다운 표. 상한을 넘으면 잘라내고 안내를 붙입니다.
    """
    with _document("excel", path, password) as wb:
        ws = _resolve_sheet(wb, sheet)
        rng = _resolve_range(ws, cell_range)
        grid = _as_grid(rng.Value)
        header = (
            f"통합문서: {_doc_label(wb, path)}  |  시트: {ws.Name}  |  "
            f"범위: {_address(rng)}"
        )
        if len(grid) == 1 and len(grid[0]) == 1 and grid[0][0] is None:
            return header + "\n\n(빈 범위입니다)"
        table = _grid_to_markdown(grid, rng.Row, rng.Column, max(1, max_rows), max(1, max_cols))
        return f"{header}\n\n{table}"


@mcp.tool()
@office_tool
def read_excel_formulas(
    path: str = "",
    sheet: str = "",
    cell_range: str = "",
    password: str = "",
) -> str:
    """Excel 시트에 들어 있는 수식을 셀 주소와 함께 조회합니다.

    계산 결과가 아니라 수식 자체를 봅니다. 계산 로직 검증, 잘못된 참조 추적,
    다른 사람이 만든 시트 인수인계처럼 "이 숫자가 어떻게 나온 건지" 봐야 할 때
    사용합니다. 값이 직접 입력된 셀은 제외하고 수식이 있는 셀만 나열합니다.

    Args:
        path: 파일 경로. 생략하면 활성 통합문서.
        sheet: 시트 이름 또는 번호. 생략하면 활성 시트.
        cell_range: 검사할 범위. 생략하면 사용 중인 영역 전체.
        password: 통합문서에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        'A1: =SUM(B1:B10)' 형식의 수식 목록과 총 개수.
    """
    with _document("excel", path, password) as wb:
        ws = _resolve_sheet(wb, sheet)
        rng = _resolve_range(ws, cell_range)
        grid = _as_grid(rng.Formula)
        base_row, base_col = rng.Row, rng.Column

        found = []
        for r, row in enumerate(grid):
            for c, v in enumerate(row):
                if isinstance(v, str) and v.startswith("="):
                    addr = f"{_col_letter(base_col + c)}{base_row + r}"
                    found.append(f"  {addr}: {v}")

        header = (
            f"통합문서: {_doc_label(wb, path)}  |  시트: {ws.Name}  |  "
            f"범위: {_address(rng)}"
        )
        if not found:
            return f"{header}\n\n이 범위에는 수식이 없습니다. 모두 직접 입력된 값입니다."
        body = "\n".join(found[:MAX_MATCHES * 4])
        note = (
            f"\n\n(수식 {len(found)}개 중 {MAX_MATCHES * 4}개만 표시. "
            "cell_range로 범위를 좁혀 보세요.)"
            if len(found) > MAX_MATCHES * 4
            else ""
        )
        return f"{header}\n\n수식 {len(found)}개:\n{body}{note}"


@mcp.tool()
@office_tool
def inspect_excel_cell(
    cell: str,
    path: str = "",
    sheet: str = "",
    password: str = "",
) -> str:
    """Excel 셀 하나를 값·수식·서식·메모·참조 관계까지 상세 조회합니다.

    "이 셀 값이 왜 이래?"를 파헤칠 때 사용합니다. 수식이 참조하는 셀(참조되는 셀)과
    이 셀을 참조하는 셀(참조하는 셀)까지 알려주므로 오류 추적에 유용합니다.

    Args:
        cell: 셀 주소. 'B7' 형식.
        path: 파일 경로. 생략하면 활성 통합문서.
        sheet: 시트 이름 또는 번호. 생략하면 활성 시트.
        password: 통합문서에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        표시 텍스트, 실제 값, 수식, 표시 형식, 메모, 참조 관계.
    """
    with _document("excel", path, password) as wb:
        ws = _resolve_sheet(wb, sheet)
        rng = _resolve_range(ws, cell)
        if rng.Cells.Count != 1:
            raise OfficeError(f"셀 하나만 지정하세요. '{cell}'은 {rng.Cells.Count}개 셀입니다.")

        out = [f"통합문서: {_doc_label(wb, path)}  |  시트: {ws.Name}  |  셀: {_address(rng)}", ""]

        def add(label, getter):
            try:
                v = getter()
            except Exception:
                return
            text = _fmt_value(v)
            if text:
                out.append(f"  {label}: {text}")

        add("표시 텍스트", lambda: rng.Text)
        add("값", lambda: rng.Value)
        formula = None
        try:
            formula = rng.Formula
        except Exception:
            pass
        if isinstance(formula, str) and formula.startswith("="):
            out.append(f"  수식: {formula}")
        add("표시 형식", lambda: rng.NumberFormat)

        # 메모는 기존 메모(Comment)와 최신 스레드 메모(CommentThreaded)로 나뉜다.
        for label, attr in (("메모", "Comment"), ("스레드 메모", "CommentThreaded")):
            try:
                c = getattr(rng, attr)
                if c is not None:
                    text = c.Text() if attr == "Comment" else c.Text
                    if text:
                        out.append(f"  {label}: {_clean(str(text))}")
            except Exception:
                pass

        # Precedents/Dependents는 대상이 없으면 예외를 던진다 — 없는 게 정상이다.
        for label, attr in (("참조되는 셀", "Precedents"), ("이 셀을 참조", "Dependents")):
            try:
                out.append(f"  {label}: {_address(getattr(rng, attr))}")
            except Exception:
                pass

        return "\n".join(out)


@mcp.tool()
@office_tool
def find_in_excel(
    query: str,
    path: str = "",
    sheet: str = "",
    password: str = "",
) -> str:
    """Excel 통합문서에서 값을 검색해 셀 주소를 찾습니다.

    "매출 합계가 어느 셀에 있지?"처럼 위치를 모를 때 사용합니다.
    대소문자를 구분하지 않는 부분 일치로 찾습니다.

    Args:
        query: 찾을 문자열.
        path: 파일 경로. 생략하면 활성 통합문서.
        sheet: 특정 시트만 검색. 생략하면 모든 시트를 검색합니다.
        password: 통합문서에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        일치한 셀의 시트/주소/값 목록.
    """
    if not query:
        return "검색어가 비어 있습니다."
    needle = query.lower()

    with _document("excel", path, password) as wb:
        sheets = [_resolve_sheet(wb, sheet)] if sheet else list(wb.Worksheets)
        hits = []
        for ws in sheets:
            used = ws.UsedRange
            grid = _as_grid(used.Value)
            base_row, base_col = used.Row, used.Column
            for r, row in enumerate(grid):
                for c, v in enumerate(row):
                    if v is None:
                        continue
                    text = _fmt_value(v)
                    if needle in text.lower():
                        addr = f"{_col_letter(base_col + c)}{base_row + r}"
                        hits.append(f"  [{ws.Name}] {addr}: {text}")
                        if len(hits) >= MAX_MATCHES:
                            break
                if len(hits) >= MAX_MATCHES:
                    break
            if len(hits) >= MAX_MATCHES:
                break

        header = f"통합문서: {_doc_label(wb, path)}  |  검색어: '{query}'"
        if not hits:
            return f"{header}\n\n일치하는 셀이 없습니다."
        note = f"\n\n(최대 {MAX_MATCHES}개까지만 표시)" if len(hits) >= MAX_MATCHES else ""
        return f"{header}\n\n{len(hits)}개 발견:\n" + "\n".join(hits) + note


@mcp.tool()
@office_tool
def read_excel_named_ranges(path: str = "", password: str = "") -> str:
    """Excel의 명명된 범위와 표(ListObject) 목록을 조회합니다.

    수식에 '매출_합계' 같은 이름이 나올 때 그게 실제로 어느 범위인지 확인하거나,
    잘 만들어진 통합문서의 데이터 구조를 파악할 때 사용합니다.

    Args:
        path: 파일 경로. 생략하면 활성 통합문서.
        password: 통합문서에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        이름별 참조 범위와 현재 값, 시트별 표 이름·범위·행 수.
    """
    with _document("excel", path, password) as wb:
        out = [f"통합문서: {_doc_label(wb, path)}", ""]

        names = []
        try:
            for nm in wb.Names:
                try:
                    refers = nm.RefersTo
                except Exception:
                    refers = "?"
                # #REF! 등 깨진 이름은 값 조회에서 예외가 난다 — 이름만 남긴다.
                try:
                    value = _fmt_value(nm.RefersToRange.Value)
                    if len(value) > 60:
                        value = value[:60] + "…"
                    names.append(f"  {nm.Name}  ->  {refers}   = {value}")
                except Exception:
                    names.append(f"  {nm.Name}  ->  {refers}")
        except Exception:
            pass
        out.append(f"명명된 범위 {len(names)}개" + (":" if names else ""))
        out.extend(names)

        tables = []
        try:
            for ws in wb.Worksheets:
                for lo in ws.ListObjects:
                    try:
                        rows = lo.ListRows.Count
                        cols = ", ".join(h.Name for h in lo.HeaderRowRange.Cells)
                        tables.append(
                            f"  [{ws.Name}] {lo.Name}  {_address(lo.Range)}  "
                            f"({rows}행)\n      열: {cols}"
                        )
                    except Exception:
                        tables.append(f"  [{ws.Name}] {lo.Name}")
        except Exception:
            pass
        out.append("")
        out.append(f"표(ListObject) {len(tables)}개" + (":" if tables else ""))
        out.extend(tables)

        return "\n".join(out)


# ─────────────────────────────── Word 도구 ───────────────────────────────


# ─────────────── Excel 쓰기 (🟡 메모리 수정 / 🔴 저장 — 3티어) ───────────────
# 쓰기는 반드시 '사용자 세션에 열려 있는' 통합문서에만 한다. _document가 안 열린
# 파일을 여는 백그라운드 인스턴스는 읽기 전용이라, 거기에 쓰면 닫을 때 조용히
# 버려진다 — 그래서 쓰기 도구는 _document를 쓰지 않고 _writable_workbook을 쓴다.


def _preview(action: str, details: list[str], tool_hint: str = "") -> str:
    """🔴 도구가 confirm 없이 호출됐을 때 돌려줄 '실행 전 확인' 프리뷰.

    실제 동작은 하지 않고 무엇을 할지 요약만 보여준다(outlook/catia와 같은 게이트).
    사용자가 확인한 뒤 같은 도구를 confirm=True로 다시 부르면 그때 실행된다.
    """
    lines = [f"⚠️ 승인 필요 — 아직 실행하지 않았습니다: {action}", ""]
    lines.extend(f"  {d}" for d in details)
    lines.append("")
    lines.append("이대로 진행하려면 같은 도구를 confirm=true 로 다시 호출하세요. " + (tool_hint or ""))
    return "\n".join(lines).rstrip()


def _writable_workbook(path: str):
    """수정 대상 통합문서를 돌려준다 — 반드시 사용자 세션에 열려 있는 것만.

    path 비움 → 활성 통합문서. path 지정 → Excel에 열려 있으면 그 문서, 아니면
    안내와 함께 실패한다(백그라운드로 열어 수정하면 변경이 버려지므로 열지 않는다).
    읽기 전용으로 열린 통합문서도 거절한다.
    """
    if not path:
        wb = _active_doc("excel")
    else:
        p = os.path.abspath(os.path.expanduser(path))
        wb = _find_open_doc("excel", p)
        if wb is None:
            raise OfficeError(
                f"'{p}'이(가) Excel에 열려 있지 않습니다. 쓰기 도구는 열려 있는 통합문서만 "
                "수정합니다(안 열린 파일을 백그라운드로 열어 쓰면 변경이 버려집니다). "
                "Excel에서 파일을 연 뒤 다시 시도하세요."
            )
    try:
        ro = bool(wb.ReadOnly)
    except Exception:  # noqa: BLE001 — 확인 불가면 일단 진행(쓰기 시점에 오류로 드러남)
        ro = False
    if ro:
        raise OfficeError(f"'{wb.Name}'은(는) 읽기 전용으로 열려 있어 수정할 수 없습니다.")
    _ctx.user_session = True
    return wb


def _coerce_cell_value(value, as_text: bool):
    """셀에 넣을 값을 정한다. 숫자 문자열은 숫자로, as_text면 문자 그대로.

    as_text일 때 '='로 시작하는 문자열은 아포스트로피를 붙여 수식이 아닌 텍스트로
    저장한다(Excel 관례 — 셀 값에는 아포스트로피가 보이지 않는다).
    """
    if value is None:
        return None  # None은 셀 비우기
    s = str(value)
    if as_text:
        return "'" + s if s.startswith("=") else s
    if isinstance(value, (int, float, bool)):
        return value
    t = s.strip()
    try:
        return int(t)
    except ValueError:
        pass
    try:
        return float(t)
    except ValueError:
        return s


@mcp.tool()
@office_tool
def write_excel_cell(
    cell: str, value: str, path: str = "", sheet: str = "", as_text: bool = False
) -> str:
    """🟡 열려 있는 Excel 통합문서의 셀 하나에 값이나 수식을 씁니다 (저장하지 않음).

    디스크에는 쓰지 않습니다 — 파일 저장은 save_workbook(confirm 필요)이 담당합니다.
    주의: COM으로 바꾼 내용은 Excel의 실행 취소(Ctrl+Z)에 쌓이지 않습니다. 대신
    이전 값을 응답에 돌려주므로, 되돌리려면 그 값으로 다시 쓰면 됩니다.

    Args:
        cell: 'B7' 형식 셀 주소.
        value: 쓸 값. '=SUM(B1:B5)'처럼 =로 시작하면 수식으로 넣습니다.
               숫자 문자열은 숫자로 변환합니다(그대로 문자로 넣으려면 as_text=True).
        path: 파일 경로. 생략하면 활성 통합문서. **Excel에 열려 있어야 합니다.**
        sheet: 시트 이름 또는 번호. 생략하면 활성 시트.
        as_text: True면 숫자/수식 변환 없이 문자열 그대로 씁니다.

    Returns:
        바꾸기 전 값(복구용)과 바꾼 후 표시 텍스트.
    """
    wb = _writable_workbook(path)
    ws = _resolve_sheet(wb, sheet)
    rng = _resolve_range(ws, cell)
    if rng.Cells.Count != 1:
        raise OfficeError(
            f"셀 하나만 지정하세요. '{cell}'은 {rng.Cells.Count}개 셀입니다. "
            "여러 셀은 write_excel_range를 쓰세요."
        )
    try:
        old = rng.Formula if str(rng.Formula).startswith("=") else _fmt_value(rng.Value)
    except Exception:  # noqa: BLE001 — 이전 값 표시는 복구용 부가정보
        old = "(읽기 실패)"
    s = str(value)
    if not as_text and s.strip().startswith("="):
        rng.Formula = s.strip()
        kind = "수식"
    else:
        rng.Value = _coerce_cell_value(value, as_text)
        kind = "값"
    return (
        f"셀 {kind} 기록: {ws.Name}!{_address(rng)} (통합문서: {_doc_label(wb, path)})\n"
        f"  이전: {old if old != '' else '(빈 셀)'}\n"
        f"  이후: {_fmt_value(rng.Text)}\n"
        "아직 저장하지 않았습니다 — 파일에 반영하려면 save_workbook을 호출하세요."
    )


@mcp.tool()
@office_tool
def write_excel_range(
    start_cell: str, values: str, path: str = "", sheet: str = "", as_text: bool = False
) -> str:
    """🟡 열려 있는 Excel 통합문서에 2차원 값 배열을 씁니다 (저장하지 않음).

    start_cell을 왼쪽 위 모서리로 해서 배열 크기만큼 씁니다. 디스크에는 쓰지
    않으며(save_workbook이 담당), 덮어써진 이전 값을 응답에 돌려줍니다(복구용 —
    COM 수정은 Excel의 Ctrl+Z에 쌓이지 않습니다).

    Args:
        start_cell: 시작(왼쪽 위) 셀 주소. 'B2' 형식.
        values: JSON 2차원 배열. 예: '[["이름","점수"],["철수",90],["영희",85]]'.
                행마다 열 수가 같아야 합니다. null은 그 셀을 비웁니다.
                '='로 시작하는 문자열은 수식으로 들어갑니다.
        path: 파일 경로. 생략하면 활성 통합문서. **Excel에 열려 있어야 합니다.**
        sheet: 시트 이름 또는 번호. 생략하면 활성 시트.
        as_text: True면 숫자/수식 변환 없이 전부 문자열로 씁니다.

    Returns:
        기록한 범위 주소와 덮어써진 이전 값 표(복구용).
    """
    import json

    try:
        grid = json.loads(values)
    except json.JSONDecodeError as e:
        raise OfficeError(f"values JSON 파싱 오류: {e} — 예: [[\"이름\",\"점수\"],[\"철수\",90]]")
    if not isinstance(grid, list) or not grid or not all(isinstance(r, list) and r for r in grid):
        raise OfficeError("values는 비어 있지 않은 2차원 배열이어야 합니다. 예: [[1,2],[3,4]]")
    cols = len(grid[0])
    if any(len(r) != cols for r in grid):
        raise OfficeError("모든 행의 열 수가 같아야 합니다(직사각형 배열).")
    rows = len(grid)
    if rows * cols > 10000:
        raise OfficeError(f"한 번에 최대 10,000셀까지 씁니다 (요청: {rows}×{cols}={rows * cols}셀).")

    wb = _writable_workbook(path)
    ws = _resolve_sheet(wb, sheet)
    anchor = _resolve_range(ws, start_cell)
    if anchor.Cells.Count != 1:
        raise OfficeError(f"start_cell은 셀 하나여야 합니다. '{start_cell}'은 {anchor.Cells.Count}개 셀입니다.")
    target = anchor.Resize(rows, cols)

    old_grid = _as_grid(target.Value)
    old_table = _grid_to_markdown(old_grid, target.Row, target.Column, 30, 20)

    target.Value = tuple(tuple(_coerce_cell_value(v, as_text) for v in row) for row in grid)
    return (
        f"범위 기록: {ws.Name}!{_address(target)} ({rows}행 × {cols}열, "
        f"통합문서: {_doc_label(wb, path)})\n\n"
        f"덮어써진 이전 값(복구용):\n{old_table}\n\n"
        "아직 저장하지 않았습니다 — 파일에 반영하려면 save_workbook을 호출하세요."
    )


@mcp.tool()
@office_tool
def save_workbook(path: str = "", confirm: bool = False) -> str:
    """🔴 열려 있는 Excel 통합문서를 현재 경로에 저장합니다(덮어쓰기). (confirm=True 필요)

    write_excel_cell/write_excel_range로 바꾼 내용을 디스크에 반영하는 단계입니다.
    confirm 없이 부르면 어떤 파일을 덮어쓸지 프리뷰만 돌려줍니다.

    Args:
        path: 파일 경로. 생략하면 활성 통합문서. Excel에 열려 있어야 합니다.
        confirm: 실제 저장하려면 True. 없으면 프리뷰만.
    """
    wb = _writable_workbook(path)
    try:
        full = wb.FullName
    except Exception:  # noqa: BLE001
        full = ""
    if not full or full == wb.Name or not os.path.isabs(full):
        return (
            "이 통합문서는 아직 디스크에 저장된 적이 없어 경로가 없습니다. "
            "Excel에서 먼저 '다른 이름으로 저장'해 경로를 정하세요."
        )
    changed = ""
    try:
        changed = "변경 있음(미저장)" if not wb.Saved else "변경 없음(이미 저장됨)"
    except Exception:  # noqa: BLE001
        pass
    if not confirm:
        details = [f"통합문서: {wb.Name}", f"경로: {full} (덮어쓰기)"]
        if changed:
            details.append(f"상태: {changed}")
        return _preview("통합문서 저장(덮어쓰기)", details, "(save_workbook ... confirm=true)")
    wb.Save()
    return f"저장 완료.\n  경로: {full}"


def _word_page_count(doc) -> int:
    """문서의 페이지 수. 백그라운드 인스턴스는 먼저 재페이지네이션해야 정확하다."""
    try:
        doc.Repaginate()
    except Exception:
        pass
    try:
        return max(1, int(doc.ComputeStatistics(WD_STAT_PAGES)))
    except Exception:
        return 1


def _word_page_text(doc, page: int, total: int) -> str:
    """지정 페이지의 본문 텍스트.

    GoTo로 그 페이지의 시작 위치를 잡고, 다음 페이지 시작 직전까지를 잘라낸다.
    마지막 페이지는 다음 페이지가 없으므로 문서 끝까지 읽는다.
    """
    start = doc.GoTo(WD_GOTO_PAGE, WD_GOTO_ABSOLUTE, page).Start
    if page >= total:
        end = doc.Content.End
    else:
        end = doc.GoTo(WD_GOTO_PAGE, WD_GOTO_ABSOLUTE, page + 1).Start
    if end <= start:
        return ""
    return doc.Range(start, end).Text


@mcp.tool()
@office_tool
def read_word_document(
    path: str = "",
    max_chars: int = MAX_CHARS,
    pages: str = "",
    password: str = "",
) -> str:
    """Word 문서의 본문을 텍스트로 읽습니다. 페이지를 지정하면 그 페이지만 읽습니다.

    문서 내용을 요약·번역·검토해야 할 때 사용합니다. 긴 문서라면
    read_word_outline으로 구조를 먼저 본 뒤 pages로 필요한 페이지만 읽는 편이 낫습니다.

    Args:
        path: .docx/.doc 파일 경로. 생략하면 Word의 활성 문서(저장 전 편집 내용 포함).
        max_chars: 반환할 최대 글자 수. 기본 20000.
        pages: 특정 페이지만 읽을 때 '1', '2-3', '1,4' 형식. 생략하면 전체.
            페이지 경계는 Word가 현재 레이아웃(여백·글꼴·프린터 설정)으로 나눈 것이라,
            환경에 따라 조금씩 달라질 수 있습니다.
        password: 문서에 열기 암호가 걸려 있을 때 지정합니다.
            이미 열려 있는 문서를 읽을 때는 필요 없습니다.

    Returns:
        문서 이름/페이지·단어 수와 본문 텍스트. pages 지정 시 페이지별로 구분해서 반환.
    """
    with _document("word", path, password) as doc:
        head = f"문서: {_doc_label(doc, path)}"

        if pages:
            total = _word_page_count(doc)
            indices = _parse_index_range(pages, total, "페이지")
            parts = []
            for p in indices:
                page_text = _clean(_word_page_text(doc, p, total))
                parts.append(f"── {p}페이지 ──\n{page_text or '(내용 없음)'}")
            head += f"  |  전체 {total}페이지 중 {len(indices)}페이지"
            body = "\n\n".join(parts)
            return f"{head}\n\n{_truncate(body, max(100, max_chars))}"

        text = _clean(doc.Content.Text)
        stats = []
        for label, name in (("페이지", "Number of pages"), ("단어", "Number of words")):
            try:
                stats.append(f"{label} {doc.BuiltinDocumentProperties(name).Value}")
            except Exception:
                pass
        if stats:
            head += f"  |  {', '.join(stats)}"
        return f"{head}\n\n{_truncate(text, max(100, max_chars))}"


@mcp.tool()
@office_tool
def read_word_outline(path: str = "", password: str = "") -> str:
    """Word 문서의 제목 구조(목차)를 조회합니다.

    긴 보고서/제안서에서 어디에 무슨 내용이 있는지 먼저 파악할 때 사용합니다.
    제목 스타일 이름이 아니라 개요 수준으로 판단하므로 한글판/영문판 Word 모두 동작합니다.

    Args:
        path: 파일 경로. 생략하면 활성 문서.
        password: 문서에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        수준별로 들여쓴 제목 목록과 각 제목이 있는 페이지 번호.
    """
    with _document("word", path, password) as doc:
        lines = []
        for para in doc.Paragraphs:
            try:
                level = int(para.OutlineLevel)
            except Exception:
                continue
            if level >= WD_BODY_TEXT:
                continue  # 본문 단락
            text = _clean(para.Range.Text)
            if not text:
                continue
            try:
                page = para.Range.Information(3)  # wdActiveEndPageNumber
                page_mark = f"  (p.{page})"
            except Exception:
                page_mark = ""
            lines.append("  " * (level - 1) + f"- {text}{page_mark}")

        head = f"문서: {_doc_label(doc, path)}"
        if not lines:
            return (
                f"{head}\n\n제목 스타일이 지정된 단락이 없습니다. "
                "read_word_document로 본문을 직접 읽으세요."
            )
        return f"{head}\n\n제목 {len(lines)}개:\n" + "\n".join(lines)


@mcp.tool()
@office_tool
def read_word_tables(
    path: str = "",
    table_index: int = 0,
    password: str = "",
) -> str:
    """Word 문서 안의 표를 마크다운 표로 읽습니다.

    계약서의 조건표, 보고서의 실적표처럼 본문 텍스트로 읽으면 뭉개지는
    표 데이터를 정확히 가져올 때 사용합니다.

    Args:
        path: 파일 경로. 생략하면 활성 문서.
        table_index: 특정 표만 읽을 때 번호(1부터). 0(기본)이면 모든 표.
        password: 문서에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        표별 마크다운 표. 병합된 셀이 있으면 해당 표는 행 단위 텍스트로 대체합니다.
    """
    with _document("word", path, password) as doc:
        total = doc.Tables.Count
        head = f"문서: {_doc_label(doc, path)}"
        if total == 0:
            return f"{head}\n\n표가 없습니다."
        # table_index=0(기본)은 '전체'를 뜻한다. 지정했을 때만 범위를 검사한다.
        if table_index and not (1 <= table_index <= total):
            raise OfficeError(f"표 번호는 1~{total} 사이여야 합니다. (전체 {total}개)")

        indices = [table_index] if table_index else range(1, total + 1)
        out = [f"{head}  |  표 {total}개", ""]
        for i in indices:
            tbl = doc.Tables(i)
            rows, cols = tbl.Rows.Count, tbl.Columns.Count
            out.append(f"[표 {i}]  {rows}행 × {cols}열")
            try:
                grid = []
                for r in range(1, rows + 1):
                    grid.append([_clean(tbl.Cell(r, c).Range.Text) for c in range(1, cols + 1)])
                header = ["| " + " | ".join(_md_escape(v) for v in grid[0]) + " |"]
                header.append("|" + "|".join(["---"] * cols) + "|")
                for row in grid[1:]:
                    header.append("| " + " | ".join(_md_escape(v) for v in row) + " |")
                out.append("\n".join(header))
            except pythoncom.com_error:
                # Cell(r, c)는 병합된 셀 자리에서 예외를 던진다. 행 단위로 물러선다.
                for r in range(1, rows + 1):
                    cells = [_clean(c.Range.Text) for c in tbl.Rows(r).Cells]
                    out.append("  " + " | ".join(cells))
                out.append("  (병합된 셀이 있어 행 단위로 표시했습니다)")
            out.append("")
        return "\n".join(out)


@mcp.tool()
@office_tool
def read_word_comments(path: str = "", password: str = "") -> str:
    """Word 문서의 검토 메모(댓글)를 조회합니다.

    "검토 의견 정리해줘", "리뷰어가 뭐라고 했어?" 같은 요청에 사용합니다.
    각 메모가 어느 본문에 달렸는지, 누가 언제 달았는지까지 가져옵니다.

    Args:
        path: 파일 경로. 생략하면 활성 문서.
        password: 문서에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        메모별 작성자·날짜·대상 본문·메모 내용.
    """
    with _document("word", path, password) as doc:
        head = f"문서: {_doc_label(doc, path)}"
        total = doc.Comments.Count
        if total == 0:
            return f"{head}\n\n검토 메모가 없습니다."

        out = [f"{head}  |  메모 {total}개", ""]
        for i in range(1, total + 1):
            cm = doc.Comments(i)
            author = getattr(cm, "Author", "") or "(작성자 없음)"
            try:
                date = _fmt_value(cm.Date)
            except Exception:
                date = ""
            try:
                page = f"p.{cm.Scope.Information(3)}  "
            except Exception:
                page = ""
            scope = _truncate(_clean(cm.Scope.Text), 200)
            text = _clean(cm.Range.Text)
            out.append(f"[{i}] {author}  {date}  {page}")
            out.append(f"    대상 본문: {scope}")
            out.append(f"    메모: {text}")
            out.append("")
        return "\n".join(out)


@mcp.tool()
@office_tool
def read_word_revisions(path: str = "", password: str = "") -> str:
    """Word 문서의 변경 내용 추적(수정 이력)을 조회합니다.

    "뭐가 바뀌었어?", "이번 개정에서 수정된 부분만 알려줘" 같은 요청에 사용합니다.
    변경 내용 추적이 켜진 상태로 편집된 문서에서만 결과가 나옵니다.

    Args:
        path: 파일 경로. 생략하면 활성 문서.
        password: 문서에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        변경별 유형(삽입/삭제/서식변경 등)·작성자·날짜·해당 텍스트.
    """
    with _document("word", path, password) as doc:
        head = f"문서: {_doc_label(doc, path)}"
        total = doc.Revisions.Count
        if total == 0:
            return (
                f"{head}\n\n추적된 변경 내용이 없습니다. "
                "(변경 내용 추적이 꺼져 있었거나, 이미 모두 적용/취소되었습니다)"
            )

        out = [f"{head}  |  변경 {total}개", ""]
        shown = min(total, MAX_MATCHES)
        for i in range(1, shown + 1):
            rev = doc.Revisions(i)
            try:
                kind = WD_REVISION_TYPES.get(int(rev.Type), f"유형{rev.Type}")
            except Exception:
                kind = "알 수 없음"
            author = getattr(rev, "Author", "") or "(작성자 없음)"
            try:
                date = _fmt_value(rev.Date)
            except Exception:
                date = ""
            try:
                text = _truncate(_clean(rev.Range.Text), 300)
            except Exception:
                text = ""
            out.append(f"[{i}] {kind}  |  {author}  {date}")
            if text:
                out.append(f"    {text}")
        if total > shown:
            out.append(f"\n(변경 {total}개 중 {shown}개만 표시)")
        return "\n".join(out)


@mcp.tool()
@office_tool
def find_in_word(
    query: str,
    path: str = "",
    context_chars: int = 100,
    password: str = "",
) -> str:
    """Word 문서에서 문자열을 검색해 위치와 앞뒤 문맥을 찾습니다.

    긴 문서에서 특정 조항이나 키워드가 어디에 있는지 찾을 때 사용합니다.
    대소문자를 구분하지 않는 부분 일치로 찾습니다.

    Args:
        query: 찾을 문자열.
        path: 파일 경로. 생략하면 활성 문서.
        context_chars: 일치 지점 앞뒤로 함께 보여줄 글자 수. 기본 100.
        password: 문서에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        일치한 단락 번호·페이지와 앞뒤 문맥이 붙은 발췌.
    """
    if not query:
        return "검색어가 비어 있습니다."
    needle = query.lower()

    with _document("word", path, password) as doc:
        hits = []
        for idx, para in enumerate(doc.Paragraphs, start=1):
            text = _clean(para.Range.Text)
            if not text or needle not in text.lower():
                continue
            pos = text.lower().find(needle)
            start = max(0, pos - context_chars)
            end = min(len(text), pos + len(query) + context_chars)
            snippet = ("…" if start else "") + text[start:end] + ("…" if end < len(text) else "")
            try:
                page = f"  p.{para.Range.Information(3)}"
            except Exception:
                page = ""
            hits.append(f"  [단락 {idx}]{page}\n      {snippet}")
            if len(hits) >= MAX_MATCHES:
                break

        head = f"문서: {_doc_label(doc, path)}  |  검색어: '{query}'"
        if not hits:
            return f"{head}\n\n일치하는 내용이 없습니다."
        note = f"\n\n(최대 {MAX_MATCHES}개까지만 표시)" if len(hits) >= MAX_MATCHES else ""
        return f"{head}\n\n{len(hits)}개 발견:\n" + "\n".join(hits) + note


@mcp.tool()
@office_tool
def read_word_section(
    heading: str,
    path: str = "",
    include_subsections: bool = True,
    occurrence: int = 1,
    max_chars: int = MAX_CHARS,
    password: str = "",
) -> str:
    """Word 문서에서 특정 제목(섹션)에 해당하는 부분만 골라 읽습니다.

    긴 문서 전체를 읽지 않고 "○○ 항목만", "결론 부분만"처럼 한 섹션만 볼 때
    사용합니다. read_word_outline으로 제목 구조를 먼저 확인하면 heading을 정하기
    쉽습니다. 제목은 제목 스타일(개요 수준)이 지정된 단락만 대상으로 하며,
    부분 일치(대소문자 무시)로 찾습니다.

    Args:
        heading: 읽을 섹션의 제목(부분 문자열 가능). 예: '3. 시험 결과'.
        path: .docx/.doc 파일 경로. 생략하면 Word의 활성 문서(저장 전 편집 내용 포함).
        include_subsections: True(기본)면 하위 제목의 내용까지 포함합니다.
            False면 같은/상위 제목이 아니라 다음 제목이 처음 나오기 전까지,
            즉 이 제목 바로 아래 본문만 읽습니다.
        occurrence: 같은 제목이 여러 번 나올 때 몇 번째를 읽을지(1부터). 기본 1.
        max_chars: 반환할 최대 글자 수. 기본 20000.
        password: 문서에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        해당 섹션의 제목과 본문 텍스트. 제목을 못 찾으면 문서의 제목 목록을 안내합니다.
    """
    if not heading or not heading.strip():
        return "heading(읽을 섹션 제목)이 비어 있습니다."
    needle = heading.strip().lower()

    with _document("word", path, password) as doc:
        head = f"문서: {_doc_label(doc, path)}"
        paras = list(doc.Paragraphs)

        # (인덱스, 개요수준, 제목텍스트) — 제목 스타일이 지정된 단락만 모은다.
        headings = []
        for i, para in enumerate(paras):
            try:
                level = int(para.OutlineLevel)
            except Exception:
                continue
            if level >= WD_BODY_TEXT:
                continue
            text = _clean(para.Range.Text)
            if text:
                headings.append((i, level, text))

        matches = [h for h in headings if needle in h[2].lower()]
        if not matches:
            if not headings:
                return (
                    f"{head}\n\n제목 스타일이 지정된 단락이 없어 섹션을 고를 수 없습니다. "
                    "read_word_document로 본문을 직접 읽거나 find_in_word로 검색하세요."
                )
            sample = "\n".join(f"  - {t}" for _, _, t in headings[:30])
            more = f"\n  …(제목 {len(headings)}개 중 30개만 표시)" if len(headings) > 30 else ""
            return (
                f"{head}\n\n'{heading}'과(와) 일치하는 제목을 찾지 못했습니다. "
                f"문서의 제목 목록:\n{sample}{more}"
            )

        occ = occurrence if occurrence and occurrence >= 1 else 1
        if occ > len(matches):
            return (
                f"{head}\n\n'{heading}'과(와) 일치하는 제목은 {len(matches)}개뿐인데 "
                f"{occ}번째를 요청했습니다. occurrence를 1~{len(matches)}로 지정하세요."
            )
        start_i, level, htext = matches[occ - 1]

        # 시작 제목 다음 단락부터, 이 섹션의 끝(같은/상위 수준 제목)까지 모은다.
        collected: list[tuple[str, str]] = []  # (종류, 텍스트) — 종류: 'h'=하위제목 'b'=본문
        for para in paras[start_i + 1:]:
            try:
                lv = int(para.OutlineLevel)
            except Exception:
                lv = WD_BODY_TEXT
            is_heading = lv < WD_BODY_TEXT
            if is_heading:
                if include_subsections and lv > level:
                    # 하위 제목 — 섹션에 포함해 소제목으로 표시하고 계속 읽는다.
                    sub = _clean(para.Range.Text)
                    if sub:
                        collected.append(("h", sub))
                    continue
                break  # 같은/상위 수준 제목이거나 하위 미포함 → 섹션 끝
            txt = _clean(para.Range.Text)
            if txt:
                collected.append(("b", txt))

        lines = [f"── {htext} ──"]
        for kind, t in collected:
            if kind == "h":
                lines.append("")
                lines.append(f"[{t}]")
            else:
                lines.append(t)
        body = "\n".join(lines)
        if not collected:
            body += "\n(이 섹션에는 본문 내용이 없습니다.)"

        if len(matches) > 1:
            head += f"  |  섹션: '{htext}' ({occ}/{len(matches)}번째 일치)"
        else:
            head += f"  |  섹션: '{htext}'"
        return f"{head}\n\n{_truncate(body, max(100, max_chars))}"


# ─────────────── Word 쓰기 (🟡 메모리 수정 / 🔴 저장 — 3티어) ───────────────
# Excel 쓰기와 같은 원칙: 반드시 '사용자 세션에 열려 있는' 문서에만 쓴다.
# _document가 안 열린 파일을 여는 백그라운드 인스턴스는 읽기 전용이라 거기에 쓰면
# 닫을 때 조용히 버려진다 — 그래서 쓰기 도구는 _document가 아니라
# _writable_document으로 사용자 세션 문서만 잡는다(_writable_workbook과 같은 구조).


def _writable_document(path: str):
    """수정 대상 Word 문서를 돌려준다 — 반드시 사용자 세션에 열려 있는 것만.

    path 비움 → 활성 문서. path 지정 → Word에 열려 있으면 그 문서, 아니면 안내와
    함께 실패한다(백그라운드로 열어 수정하면 변경이 버려지므로 열지 않는다).
    읽기 전용이거나 편집 제한이 걸린 문서도 거절한다.
    """
    if not path:
        doc = _active_doc("word")
    else:
        p = os.path.abspath(os.path.expanduser(path))
        doc = _find_open_doc("word", p)
        if doc is None:
            raise OfficeError(
                f"'{p}'이(가) Word에 열려 있지 않습니다. 쓰기 도구는 열려 있는 문서만 "
                "수정합니다(안 열린 파일을 백그라운드로 열어 쓰면 변경이 버려집니다). "
                "Word에서 파일을 연 뒤 다시 시도하세요."
            )
    try:
        if bool(doc.ReadOnly):
            raise OfficeError(f"'{doc.Name}'은(는) 읽기 전용으로 열려 있어 수정할 수 없습니다.")
    except OfficeError:
        raise
    except Exception:  # noqa: BLE001 — 확인 불가면 일단 진행(쓰기 시점에 오류로 드러남)
        pass
    try:
        if int(doc.ProtectionType) != WD_NO_PROTECTION:
            raise OfficeError(
                f"'{doc.Name}'에는 편집 제한(문서 보호)이 걸려 있어 수정할 수 없습니다. "
                "Word에서 '편집 제한 중지'로 보호를 푼 뒤 다시 시도하세요."
            )
    except OfficeError:
        raise
    except Exception:  # noqa: BLE001 — ProtectionType을 못 읽으면 그냥 진행
        pass
    _ctx.user_session = True
    return doc


def _insert_text(rng, text: str, as_new_paragraph: bool, before: bool):
    """Range 위치에 텍스트를 넣는다. as_new_paragraph면 단락 구분(\\r)을 함께 넣는다."""
    if as_new_paragraph:
        rng.InsertBefore(text + "\r") if before else rng.InsertAfter("\r" + text)
    else:
        rng.InsertBefore(text) if before else rng.InsertAfter(text)


@mcp.tool()
@office_tool
def replace_in_word(
    find_text: str,
    replace_text: str,
    path: str = "",
    match_case: bool = False,
    whole_word: bool = False,
    replace_all: bool = True,
) -> str:
    """🟡 열려 있는 Word 문서에서 텍스트를 찾아 바꿉니다 (저장하지 않음).

    "○○를 △△로 다 바꿔줘"처럼 문서의 특정 부분을 수정할 때 사용합니다.
    디스크에는 쓰지 않으며(파일 저장은 save_document, confirm 필요), 바꾼 내용은
    저장 전이라 Word에서 Ctrl+Z나 '저장 없이 닫기'로 되돌릴 수 있습니다.

    Args:
        find_text: 찾을 텍스트.
        replace_text: 바꿀 텍스트. 빈 문자열이면 찾은 텍스트를 삭제합니다.
        path: 파일 경로. 생략하면 활성 문서. **Word에 열려 있어야 합니다.**
        match_case: True면 대소문자를 구분합니다. 기본 False.
        whole_word: True면 단어 전체가 일치할 때만 바꿉니다. 기본 False.
        replace_all: True(기본)면 모두 바꾸고, False면 첫 번째 하나만 바꿉니다.

    Returns:
        바꾼 곳 수와 저장 안내. 일치가 없으면 그 사실을 알립니다.
    """
    if not find_text:
        return "find_text(찾을 텍스트)가 비어 있습니다."

    doc = _writable_document(path)
    head = f"문서: {_doc_label(doc, path)}"

    # 실제 치환과 같은 조건으로 먼저 개수를 센다(치환 결과에 find_text가 다시
    # 생겨도 개수가 흔들리지 않도록 치환 전에 별도 범위로 센다).
    probe = doc.Content.Duplicate
    pf = probe.Find
    pf.ClearFormatting()
    count = 0
    while pf.Execute(find_text, match_case, whole_word, False, False, False, True, WD_FIND_STOP, False):
        count += 1
        probe.Collapse(WD_COLLAPSE_END)
        if count >= 100000:  # 폭주 방지 상한
            break
    if count == 0:
        return f"{head}\n\n'{find_text}'을(를) 찾지 못해 바꾼 내용이 없습니다."

    rng = doc.Content
    rf = rng.Find
    rf.ClearFormatting()
    rf.Replacement.ClearFormatting()
    mode = WD_REPLACE_ALL if replace_all else WD_REPLACE_ONE
    rf.Execute(
        find_text, match_case, whole_word, False, False, False,
        True, WD_FIND_CONTINUE, False, replace_text, mode,
    )
    done = count if replace_all else 1
    verb = "삭제" if replace_text == "" else "치환"
    return (
        f"찾아 바꾸기 완료({verb}): '{find_text}' → "
        f"{'(삭제)' if replace_text == '' else repr(replace_text)}  ({head})\n"
        f"  바꾼 곳: {done}곳" + (f" (일치 {count}곳 중 첫 1곳만)" if not replace_all and count > 1 else "") + "\n"
        "아직 저장하지 않았습니다 — 파일에 반영하려면 save_document를 호출하세요."
    )


@mcp.tool()
@office_tool
def insert_word_text(
    text: str,
    path: str = "",
    position: str = "end",
    anchor: str = "",
    match_case: bool = False,
    as_new_paragraph: bool = True,
) -> str:
    """🟡 열려 있는 Word 문서의 특정 위치에 텍스트를 입력합니다 (저장하지 않음).

    문서 끝/처음에 문단을 덧붙이거나, 특정 문구(anchor)를 기준으로 그 앞·뒤에
    끼워 넣거나, anchor 자체를 교체할 수 있습니다. 디스크에는 쓰지 않으며
    (파일 저장은 save_document, confirm 필요), 저장 전이라 Word에서 Ctrl+Z나
    '저장 없이 닫기'로 되돌릴 수 있습니다.

    Args:
        text: 넣을 텍스트.
        path: 파일 경로. 생략하면 활성 문서. **Word에 열려 있어야 합니다.**
        position: 넣을 위치.
            'end'(기본) 문서 끝 / 'start' 문서 처음 /
            'after' anchor 뒤 / 'before' anchor 앞 / 'replace' anchor를 text로 교체.
            after/before/replace는 anchor가 필요합니다.
        anchor: 기준이 될 기존 문구(부분 문자열). position이 after/before/replace일 때 사용.
        match_case: anchor를 찾을 때 대소문자 구분. 기본 False.
        as_new_paragraph: True(기본)면 새 문단으로 넣습니다(줄바꿈 포함).
            False면 기존 문단에 이어 붙입니다. position이 'replace'면 무시됩니다.

    Returns:
        수행한 동작 요약과 저장 안내.
    """
    if not text:
        return "text(넣을 텍스트)가 비어 있습니다."
    pos = (position or "end").strip().lower()
    valid = {"end", "start", "after", "before", "replace"}
    if pos not in valid:
        return f"position은 {', '.join(sorted(valid))} 중 하나여야 합니다. (받은 값: '{position}')"
    if pos in ("after", "before", "replace") and not anchor:
        return f"position='{pos}'에는 기준 문구 anchor가 필요합니다."

    doc = _writable_document(path)
    head = f"문서: {_doc_label(doc, path)}"

    if pos in ("after", "before", "replace"):
        rng = doc.Content
        f = rng.Find
        f.ClearFormatting()
        # 찾기 전용(치환 인자 생략) — 성공하면 rng가 찾은 텍스트 범위로 바뀐다.
        if not f.Execute(anchor, match_case, False, False, False, False, True, WD_FIND_STOP, False):
            raise OfficeError(f"'{anchor}'을(를) 문서에서 찾지 못했습니다. anchor를 확인하세요.")
        if pos == "replace":
            rng.Text = text
            action = f"'{anchor}'을(를) 교체"
        elif pos == "after":
            rng.Collapse(WD_COLLAPSE_END)
            _insert_text(rng, text, as_new_paragraph, before=False)
            action = f"'{anchor}' 뒤에 삽입"
        else:  # before
            rng.Collapse(WD_COLLAPSE_START)
            _insert_text(rng, text, as_new_paragraph, before=True)
            action = f"'{anchor}' 앞에 삽입"
    else:
        rng = doc.Content
        if pos == "end":
            rng.Collapse(WD_COLLAPSE_END)
            _insert_text(rng, text, as_new_paragraph, before=False)
            action = "문서 끝에 추가"
        else:  # start
            rng.Collapse(WD_COLLAPSE_START)
            _insert_text(rng, text, as_new_paragraph, before=True)
            action = "문서 처음에 추가"

    preview = text if len(text) <= 60 else text[:60] + "…"
    return (
        f"입력 완료: {action}  ({head})\n"
        f"  넣은 내용: {preview}\n"
        "아직 저장하지 않았습니다 — 파일에 반영하려면 save_document를 호출하세요."
    )


@mcp.tool()
@office_tool
def save_document(path: str = "", confirm: bool = False) -> str:
    """🔴 열려 있는 Word 문서를 현재 경로에 저장합니다(덮어쓰기). (confirm=True 필요)

    replace_in_word/insert_word_text로 바꾼 내용을 디스크에 반영하는 단계입니다.
    confirm 없이 부르면 어떤 파일을 덮어쓸지 프리뷰만 돌려줍니다.

    Args:
        path: 파일 경로. 생략하면 활성 문서. Word에 열려 있어야 합니다.
        confirm: 실제 저장하려면 True. 없으면 프리뷰만.
    """
    doc = _writable_document(path)
    try:
        full = doc.FullName
    except Exception:  # noqa: BLE001
        full = ""
    if not full or full == doc.Name or not os.path.isabs(full):
        return (
            "이 문서는 아직 디스크에 저장된 적이 없어 경로가 없습니다. "
            "Word에서 먼저 '다른 이름으로 저장'해 경로를 정하세요."
        )
    changed = ""
    try:
        changed = "변경 있음(미저장)" if not doc.Saved else "변경 없음(이미 저장됨)"
    except Exception:  # noqa: BLE001
        pass
    if not confirm:
        details = [f"문서: {doc.Name}", f"경로: {full} (덮어쓰기)"]
        if changed:
            details.append(f"상태: {changed}")
        return _preview("문서 저장(덮어쓰기)", details, "(save_document ... confirm=true)")
    doc.Save()
    return f"저장 완료.\n  경로: {full}"


# ──────────────────────────── PowerPoint 도구 ────────────────────────────


def _slide_title(slide) -> str:
    """슬라이드 제목. 제목 개체틀이 없으면 첫 텍스트 도형으로 대신한다."""
    try:
        if slide.Shapes.HasTitle:
            return _clean(slide.Shapes.Title.TextFrame.TextRange.Text)
    except Exception:
        pass
    try:
        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                return _clean(shape.TextFrame.TextRange.Text).split("\n")[0]
    except Exception:
        pass
    return "(제목 없음)"


@mcp.tool()
@office_tool
def read_powerpoint_outline(path: str = "", password: str = "") -> str:
    """PowerPoint 발표자료의 슬라이드별 제목 목록을 조회합니다.

    발표자료의 전체 흐름을 먼저 파악하거나, 몇 번 슬라이드를 자세히 볼지
    고를 때 사용합니다. 슬라이드별 도형·표·차트 유무도 함께 알려줍니다.

    Args:
        path: .pptx/.ppt 파일 경로. 생략하면 PowerPoint의 활성 발표자료.
        password: 발표자료에 열기 암호가 걸려 있을 때 지정합니다.
            이미 열려 있는 발표자료를 읽을 때는 필요 없습니다.

    Returns:
        슬라이드 번호·제목·구성 요소 요약 목록.
    """
    with _document("ppt", path, password) as pres:
        total = pres.Slides.Count
        out = [f"발표자료: {_doc_label(pres, path)}  |  슬라이드 {total}장", ""]
        for slide in pres.Slides:
            marks = []
            try:
                tables = sum(1 for s in slide.Shapes if s.HasTable)
                charts = sum(1 for s in slide.Shapes if s.HasChart)
                if tables:
                    marks.append(f"표 {tables}")
                if charts:
                    marks.append(f"차트 {charts}")
            except Exception:
                pass
            try:
                notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
                if _clean(notes):
                    marks.append("노트")
            except Exception:
                pass
            extra = f"   <{', '.join(marks)}>" if marks else ""
            out.append(f"  {slide.SlideIndex:>3}. {_slide_title(slide)}{extra}")
        return "\n".join(out)


@mcp.tool()
@office_tool
def read_powerpoint_slides(
    path: str = "", slides: str = "", password: str = ""
) -> str:
    """PowerPoint 슬라이드의 본문 텍스트를 도형 단위로 읽습니다.

    발표자료 내용을 요약·검토하거나 문서로 옮길 때 사용합니다.
    전체를 읽으면 길어지므로 slides로 범위를 좁힐 수 있습니다.

    Args:
        path: 파일 경로. 생략하면 활성 발표자료.
        slides: 읽을 슬라이드. '3', '2-5', '1,4,7' 형식. 생략하면 전체.
        password: 발표자료에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        슬라이드별 제목과 텍스트 도형 내용.
    """
    with _document("ppt", path, password) as pres:
        total = pres.Slides.Count
        indices = _parse_index_range(slides, total, "슬라이드")
        out = [f"발표자료: {_doc_label(pres, path)}  |  슬라이드 {total}장 중 {len(indices)}장", ""]

        for i in indices:
            slide = pres.Slides(i)
            out.append(f"── 슬라이드 {i}: {_slide_title(slide)} ──")
            texts = []
            for shape in slide.Shapes:
                try:
                    if not (shape.HasTextFrame and shape.TextFrame.HasText):
                        continue
                    text = _clean(shape.TextFrame.TextRange.Text)
                except Exception:
                    continue
                if text:
                    texts.append(text)
            if texts:
                # 제목은 이미 머리에 찍었으므로 본문만 남긴다.
                body = [t for t in texts[1:]] if len(texts) > 1 else texts
                out.extend(f"  {line}" for t in body for line in t.split("\n"))
            else:
                out.append("  (텍스트 없음)")
            out.append("")
        return _truncate("\n".join(out), MAX_CHARS)


@mcp.tool()
@office_tool
def read_powerpoint_notes(
    path: str = "", slides: str = "", password: str = ""
) -> str:
    """PowerPoint의 발표자 노트를 조회합니다.

    슬라이드 본문에는 없는 실제 발표 스크립트나 배경 설명이 노트에 있습니다.
    발표 준비, 발표자료 인수인계, 회의록 작성에 사용합니다.

    Args:
        path: 파일 경로. 생략하면 활성 발표자료.
        slides: 읽을 슬라이드. '3', '2-5', '1,4,7' 형식. 생략하면 전체.
        password: 발표자료에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        노트가 있는 슬라이드의 번호·제목·노트 내용.
    """
    with _document("ppt", path, password) as pres:
        total = pres.Slides.Count
        indices = _parse_index_range(slides, total, "슬라이드")
        out = [f"발표자료: {_doc_label(pres, path)}", ""]

        found = 0
        for i in indices:
            slide = pres.Slides(i)
            try:
                # 노트 페이지의 두 번째 개체틀이 본문 노트다 (첫 번째는 슬라이드 이미지).
                notes = _clean(slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text)
            except Exception:
                notes = ""
            if not notes:
                continue
            found += 1
            out.append(f"── 슬라이드 {i}: {_slide_title(slide)} ──")
            out.extend(f"  {line}" for line in notes.split("\n"))
            out.append("")

        if found == 0:
            return f"발표자료: {_doc_label(pres, path)}\n\n발표자 노트가 있는 슬라이드가 없습니다."
        out.insert(1, f"노트가 있는 슬라이드 {found}장")
        return _truncate("\n".join(out), MAX_CHARS)


@mcp.tool()
@office_tool
def read_powerpoint_tables(
    path: str = "", slides: str = "", password: str = ""
) -> str:
    """PowerPoint 슬라이드 안의 표를 마크다운 표로 읽습니다.

    발표자료의 실적표·비교표처럼 텍스트로 읽으면 뭉개지는 데이터를
    정확히 가져올 때 사용합니다.

    Args:
        path: 파일 경로. 생략하면 활성 발표자료.
        slides: 읽을 슬라이드. '3', '2-5', '1,4,7' 형식. 생략하면 전체.
        password: 발표자료에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        슬라이드별 표를 마크다운 표로 변환한 결과.
    """
    with _document("ppt", path, password) as pres:
        total = pres.Slides.Count
        indices = _parse_index_range(slides, total, "슬라이드")
        out = [f"발표자료: {_doc_label(pres, path)}", ""]

        found = 0
        for i in indices:
            slide = pres.Slides(i)
            for shape in slide.Shapes:
                try:
                    if not shape.HasTable:
                        continue
                except Exception:
                    continue
                found += 1
                table = shape.Table
                rows, cols = table.Rows.Count, table.Columns.Count
                out.append(f"── 슬라이드 {i} / 표 {found}  ({rows}행 × {cols}열) ──")
                grid = []
                for r in range(1, rows + 1):
                    grid.append(
                        [
                            _md_escape(_clean(table.Cell(r, c).Shape.TextFrame.TextRange.Text))
                            for c in range(1, cols + 1)
                        ]
                    )
                lines = ["| " + " | ".join(grid[0]) + " |", "|" + "|".join(["---"] * cols) + "|"]
                lines.extend("| " + " | ".join(row) + " |" for row in grid[1:])
                out.append("\n".join(lines))
                out.append("")

        if found == 0:
            return f"발표자료: {_doc_label(pres, path)}\n\n표가 있는 슬라이드가 없습니다."
        out.insert(1, f"표 {found}개")
        return _truncate("\n".join(out), MAX_CHARS)


# ─────────── PowerPoint 쓰기 (🟡 메모리 수정 / 🔴 저장 — 3티어) ───────────
# Word/Excel과 같은 원칙: 쓰기는 반드시 '사용자 세션에 열려 있는' 발표자료에만
# 한다. _document가 백그라운드로 여는 인스턴스는 읽기용이라, 거기에 쓰면 닫을 때
# 조용히 버려진다 — 그래서 쓰기 도구는 _document 대신 _writable_presentation을 쓴다.
# 좌표·크기는 PowerPoint COM이 쓰는 포인트(pt) 단위이고, unit 인자로 cm/mm/in도 받는다.


def _writable_presentation(path: str):
    """수정 대상 발표자료를 돌려준다 — 반드시 사용자 세션에 열려 있는 것만.

    path 비움 → 활성 발표자료. path 지정 → PowerPoint에 열려 있으면 그 발표자료,
    아니면 안내와 함께 실패한다(백그라운드로 열어 수정하면 변경이 버려지므로 열지 않는다).
    읽기 전용으로 열린 발표자료도 거절한다.
    """
    if not path:
        pres = _active_doc("ppt")
    else:
        p = os.path.abspath(os.path.expanduser(path))
        pres = _find_open_doc("ppt", p)
        if pres is None:
            raise OfficeError(
                f"'{p}'이(가) PowerPoint에 열려 있지 않습니다. 쓰기 도구는 열려 있는 "
                "발표자료만 수정합니다(안 열린 파일을 백그라운드로 열어 쓰면 변경이 "
                "버려집니다). PowerPoint에서 파일을 연 뒤 다시 시도하세요."
            )
    try:
        if bool(pres.ReadOnly):
            raise OfficeError(f"'{pres.Name}'은(는) 읽기 전용으로 열려 있어 수정할 수 없습니다.")
    except OfficeError:
        raise
    except Exception:  # noqa: BLE001 — 확인 불가면 일단 진행(쓰기 시점에 오류로 드러남)
        pass
    _ctx.user_session = True
    return pres


def _resolve_slide(pres, slide):
    """1-based 슬라이드 번호를 Slide 객체로 바꾼다."""
    total = pres.Slides.Count
    try:
        i = int(slide)
    except (TypeError, ValueError):
        raise OfficeError(f"슬라이드 번호는 정수여야 합니다(받은 값: '{slide}').")
    if not (1 <= i <= total):
        raise OfficeError(f"슬라이드 {slide}이(가) 없습니다. 이 발표자료는 전체 {total}장입니다.")
    return pres.Slides(i)


def _resolve_shape(slide, shape):
    """도형을 1-based 번호 또는 이름으로 찾는다.

    list_slide_shapes가 알려 주는 번호/이름을 그대로 받는다. 못 찾으면 그 슬라이드의
    도형 목록을 곁들여 안내한다.
    """
    shapes = slide.Shapes
    s = str(shape).strip()
    if s.isdigit():
        idx = int(s)
        if not (1 <= idx <= shapes.Count):
            raise OfficeError(
                f"이 슬라이드에는 도형이 {shapes.Count}개뿐입니다(요청: {idx}번). "
                "list_slide_shapes로 번호를 확인하세요."
            )
        return shapes.Item(idx)
    try:
        return shapes.Item(s)
    except pythoncom.com_error:
        names = []
        for k in range(1, shapes.Count + 1):
            try:
                names.append(f"{k}:{shapes.Item(k).Name}")
            except Exception:
                names.append(str(k))
        raise OfficeError(
            f"'{shape}' 도형을 찾을 수 없습니다. 이 슬라이드의 도형: {', '.join(names) or '(없음)'}"
        )


def _to_points(value: float, unit: str) -> float:
    """길이 값을 지정 단위에서 포인트(pt)로 환산한다."""
    u = (unit or "pt").strip().lower()
    if u not in PPT_UNIT_TO_PT:
        raise OfficeError(
            f"unit은 {', '.join(sorted(set(PPT_UNIT_TO_PT)))} 중 하나여야 합니다(받은 값: '{unit}')."
        )
    try:
        return float(value) * PPT_UNIT_TO_PT[u]
    except (TypeError, ValueError):
        raise OfficeError(f"길이 값이 숫자가 아닙니다: '{value}'.")


def _fmt_len(pt: float) -> str:
    """포인트 길이를 pt와 cm로 함께 보여 준다(사람이 크기를 가늠하기 쉽게)."""
    try:
        return f"{pt:.1f}pt ({pt / 28.3464567:.2f}cm)"
    except Exception:  # noqa: BLE001
        return f"{pt}pt"


def _hex_to_ole(color: str) -> int:
    """'#RRGGBB' 또는 'RRGGBB'를 Office가 쓰는 OLE 색값(0xBBGGRR)으로 바꾼다."""
    h = str(color).strip().lstrip("#")
    if len(h) != 6:
        raise OfficeError(f"색은 '#RRGGBB' 6자리 16진수로 지정하세요(받은 값: '{color}').")
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError:
        raise OfficeError(f"색 '{color}'을(를) 16진수로 해석할 수 없습니다(예: '#1F4E79').")
    return r + (g << 8) + (b << 16)


def _shape_kind(shape) -> str:
    """도형의 종류를 사람이 읽을 이름으로. 표/차트/텍스트는 더 구체적으로."""
    try:
        if shape.HasTable:
            return "표"
    except Exception:  # noqa: BLE001
        pass
    try:
        if shape.HasChart:
            return "차트"
    except Exception:  # noqa: BLE001
        pass
    try:
        return MSO_SHAPE_TYPES.get(int(shape.Type), f"종류{int(shape.Type)}")
    except Exception:  # noqa: BLE001
        return "도형"


@mcp.tool()
@office_tool
def list_slide_shapes(slide: int, path: str = "", password: str = "") -> str:
    """🟢 한 슬라이드의 도형 목록을 번호·이름·종류·위치·크기와 함께 조회합니다.

    도형을 수정(set_powerpoint_text/set_shape_size/set_shape_position/format_shape)
    하거나 지우기 전에, 대상 도형의 '번호'나 '이름'을 먼저 확인하는 용도입니다.
    위치·크기는 포인트(pt)와 cm로 함께 보여 줍니다.

    Args:
        slide: 슬라이드 번호(1부터). read_powerpoint_outline로 확인하세요.
        path: 파일 경로. 생략하면 활성 발표자료.
        password: 발표자료에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        도형별 번호·이름·종류·위치(left,top)·크기(width,height)·텍스트 미리보기.
    """
    with _document("ppt", path, password) as pres:
        slide_obj = _resolve_slide(pres, slide)
        shapes = slide_obj.Shapes
        out = [
            f"발표자료: {_doc_label(pres, path)}  |  슬라이드 {slide}: {_slide_title(slide_obj)}",
            f"도형 {shapes.Count}개",
            "",
        ]
        for k in range(1, shapes.Count + 1):
            shape = shapes.Item(k)
            try:
                name = shape.Name
            except Exception:  # noqa: BLE001
                name = "?"
            kind = _shape_kind(shape)
            try:
                pos = f"위치({_fmt_len(shape.Left)}, {_fmt_len(shape.Top)})"
                size = f"크기({_fmt_len(shape.Width)} × {_fmt_len(shape.Height)})"
            except Exception:  # noqa: BLE001 — 선/그룹 등 일부는 위치·크기가 없다
                pos, size = "위치(?)", "크기(?)"
            out.append(f"  [{k}] {name}  · {kind}")
            out.append(f"      {pos}  {size}")
            try:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    txt = _clean(shape.TextFrame.TextRange.Text)
                    if txt:
                        preview = txt if len(txt) <= 60 else txt[:60] + "…"
                        out.append(f"      텍스트: {preview}")
            except Exception:  # noqa: BLE001
                pass
        if shapes.Count == 0:
            out.append("  (도형 없음)")
        return _truncate("\n".join(out), MAX_CHARS)


@mcp.tool()
@office_tool
def set_powerpoint_text(slide: int, shape: str, text: str, path: str = "") -> str:
    """🟡 열려 있는 발표자료에서 한 도형의 텍스트를 통째로 바꿉니다 (저장하지 않음).

    제목/본문/텍스트 상자 등 텍스트를 담을 수 있는 도형의 내용을 새 text로 교체합니다.
    디스크에는 쓰지 않으며(저장은 save_presentation, confirm 필요), 저장 전이라
    PowerPoint에서 Ctrl+Z로 되돌릴 수 있습니다.

    Args:
        slide: 슬라이드 번호(1부터).
        shape: 대상 도형의 번호 또는 이름(list_slide_shapes로 확인).
        text: 새 텍스트. 여러 줄은 개행(\\n)으로 구분합니다.
        path: 파일 경로. 생략하면 활성 발표자료. **PowerPoint에 열려 있어야 합니다.**

    Returns:
        바꾸기 전/후 텍스트 요약과 저장 안내.
    """
    pres = _writable_presentation(path)
    slide_obj = _resolve_slide(pres, slide)
    shape_obj = _resolve_shape(slide_obj, shape)
    try:
        has_tf = bool(shape_obj.HasTextFrame)
    except Exception:  # noqa: BLE001
        has_tf = False
    if not has_tf:
        raise OfficeError(
            f"'{shape_obj.Name}' 도형은 텍스트를 담을 수 없습니다(그림/선 등). "
            "텍스트 상자나 개체틀을 대상으로 지정하세요."
        )
    tr = shape_obj.TextFrame.TextRange
    try:
        old = _clean(tr.Text)
    except Exception:  # noqa: BLE001
        old = "(읽기 실패)"
    # PowerPoint는 개행을 \r로 다룬다.
    tr.Text = str(text).replace("\r\n", "\n").replace("\n", "\r")
    old_p = old if len(old) <= 60 else old[:60] + "…"
    new_p = str(text) if len(str(text)) <= 60 else str(text)[:60] + "…"
    return (
        f"텍스트 교체: 슬라이드 {slide} / '{shape_obj.Name}' (발표자료: {_doc_label(pres, path)})\n"
        f"  이전: {old_p if old else '(빈 텍스트)'}\n"
        f"  이후: {new_p}\n"
        "아직 저장하지 않았습니다 — 파일에 반영하려면 save_presentation을 호출하세요."
    )


@mcp.tool()
@office_tool
def replace_in_powerpoint(
    find_text: str,
    replace_text: str,
    path: str = "",
    slides: str = "",
    match_case: bool = False,
    whole_word: bool = False,
    replace_all: bool = True,
) -> str:
    """🟡 열려 있는 발표자료의 텍스트를 찾아 바꿉니다 (여러 슬라이드, 저장하지 않음).

    "○○를 △△로 다 바꿔줘"처럼 발표자료 전반의 문구를 일괄 수정할 때 씁니다.
    각 도형의 TextRange.Replace를 써서 서식은 유지한 채 글자만 바꿉니다.
    디스크에는 쓰지 않으며(저장은 save_presentation), Ctrl+Z로 되돌릴 수 있습니다.

    Args:
        find_text: 찾을 텍스트.
        replace_text: 바꿀 텍스트. 빈 문자열이면 찾은 텍스트를 삭제합니다.
        path: 파일 경로. 생략하면 활성 발표자료. **PowerPoint에 열려 있어야 합니다.**
        slides: 대상 슬라이드. '3', '2-5', '1,4,7' 형식. 생략하면 전체.
        match_case: True면 대소문자를 구분합니다. 기본 False.
        whole_word: True면 단어 전체가 일치할 때만 바꿉니다. 기본 False.
        replace_all: True(기본)면 모두 바꾸고, False면 슬라이드마다 첫 하나만 바꿉니다.

    Returns:
        바꾼 곳 수와 저장 안내. 일치가 없으면 그 사실을 알립니다.
    """
    if not find_text:
        return "find_text(찾을 텍스트)가 비어 있습니다."
    pres = _writable_presentation(path)
    total = pres.Slides.Count
    indices = _parse_index_range(slides, total, "슬라이드")
    head = f"발표자료: {_doc_label(pres, path)}"

    mc = MSO_TRUE if match_case else MSO_FALSE
    ww = MSO_TRUE if whole_word else MSO_FALSE
    count = 0
    for i in indices:
        slide_obj = pres.Slides(i)
        for shape in slide_obj.Shapes:
            try:
                if not (shape.HasTextFrame and shape.TextFrame.HasText):
                    continue
                tr = shape.TextFrame.TextRange
            except Exception:  # noqa: BLE001
                continue
            matched = False
            # TextRange.Replace(FindWhat, ReplaceWhat, After, MatchCase, WholeWords)는
            # 한 번에 하나를 바꾸고 바뀐 범위를 돌려준다(끝나면 None). 인자는 pywin32가
            # 키워드를 흘릴 수 있어 위치로 넘긴다. After(시작 문자 위치)를 옮겨 이어 찾는다.
            start = 0
            while True:
                try:
                    found = tr.Replace(find_text, replace_text, start, mc, ww)
                except pythoncom.com_error:
                    break
                if not found or _clean(getattr(found, "Text", "")) == "":
                    break
                count += 1
                matched = True
                if not replace_all:
                    break
                try:
                    # 바뀐 텍스트 끝 다음부터 이어 찾는다(치환문에 find_text가 또 있어도 무한루프 방지).
                    start = int(found.Start) + max(len(replace_text), 1)
                except Exception:  # noqa: BLE001
                    break
                if count >= 100000:  # 폭주 방지 상한
                    break
            if matched and not replace_all:
                # replace_all=False는 슬라이드마다 첫 하나만 — 다음 슬라이드로.
                break
    if count == 0:
        return f"{head}\n\n'{find_text}'을(를) 찾지 못해 바꾼 내용이 없습니다."
    verb = "삭제" if replace_text == "" else "치환"
    return (
        f"찾아 바꾸기 완료({verb}): '{find_text}' → "
        f"{'(삭제)' if replace_text == '' else repr(replace_text)}  ({head})\n"
        f"  바꾼 곳: {count}곳 (슬라이드 {len(indices)}장 대상)\n"
        "아직 저장하지 않았습니다 — 파일에 반영하려면 save_presentation을 호출하세요."
    )


@mcp.tool()
@office_tool
def get_shape_size(slide: int, shape: str, path: str = "", password: str = "") -> str:
    """🟢 한 도형의 현재 위치와 크기를 조회합니다 (pt·cm·inch 함께).

    크기를 바꾸기(set_shape_size) 전에 현재 값을 확인하거나, 여러 도형의 크기를
    맞출 때 기준값을 얻는 용도입니다. 도형 하나만 빠르게 볼 때 씁니다
    (슬라이드 전체는 list_slide_shapes).

    Args:
        slide: 슬라이드 번호(1부터).
        shape: 도형 번호 또는 이름(list_slide_shapes로 확인).
        path: 파일 경로. 생략하면 활성 발표자료.
        password: 발표자료에 열기 암호가 걸려 있을 때 지정합니다.

    Returns:
        위치(left, top)·크기(width, height)를 pt/cm/inch로.
    """
    with _document("ppt", path, password) as pres:
        slide_obj = _resolve_slide(pres, slide)
        shape_obj = _resolve_shape(slide_obj, shape)
        try:
            left, top = shape_obj.Left, shape_obj.Top
            width, height = shape_obj.Width, shape_obj.Height
        except Exception:  # noqa: BLE001
            raise OfficeError(
                f"'{shape_obj.Name}' 도형의 위치·크기를 읽을 수 없습니다(선/그룹 등일 수 있음)."
            )

        def _triple(pt):
            return f"{pt:.1f}pt / {pt / 28.3464567:.2f}cm / {pt / 72.0:.2f}in"

        return (
            f"발표자료: {_doc_label(pres, path)}  |  슬라이드 {slide} / '{shape_obj.Name}' "
            f"({_shape_kind(shape_obj)})\n"
            f"  위치 left : {_triple(left)}\n"
            f"  위치 top  : {_triple(top)}\n"
            f"  너비 width : {_triple(width)}\n"
            f"  높이 height: {_triple(height)}"
        )


@mcp.tool()
@office_tool
def set_shape_size(
    slide: int,
    shape: str,
    width: float | None = None,
    height: float | None = None,
    unit: str = "pt",
    lock_aspect: bool = False,
    path: str = "",
) -> str:
    """🟡 한 도형의 크기(너비·높이)를 바꿉니다 (저장하지 않음).

    width/height 중 준 것만 바꿉니다(하나만 주면 그 한 변만). 디스크에는 쓰지 않으며
    (저장은 save_presentation), Ctrl+Z로 되돌릴 수 있습니다.

    Args:
        slide: 슬라이드 번호(1부터).
        shape: 도형 번호 또는 이름(list_slide_shapes로 확인).
        width: 새 너비. 생략하면 너비는 그대로.
        height: 새 높이. 생략하면 높이는 그대로.
        unit: width/height의 단위. 'pt'(기본)/'cm'/'mm'/'in'.
        lock_aspect: True면 가로세로 비율을 고정합니다(한 변만 줘도 다른 변이 따라옴).
            이때는 width나 height 중 하나만 주는 것이 자연스럽습니다.
        path: 파일 경로. 생략하면 활성 발표자료. **PowerPoint에 열려 있어야 합니다.**

    Returns:
        바꾸기 전/후 크기와 저장 안내.
    """
    if width is None and height is None:
        return "width나 height 중 적어도 하나는 지정하세요."
    pres = _writable_presentation(path)
    slide_obj = _resolve_slide(pres, slide)
    shape_obj = _resolve_shape(slide_obj, shape)
    try:
        old_w, old_h = float(shape_obj.Width), float(shape_obj.Height)
    except Exception:  # noqa: BLE001
        raise OfficeError(f"'{shape_obj.Name}' 도형은 크기를 바꿀 수 없습니다(선/그룹 등일 수 있음).")

    try:
        # msoTrue=-1 / msoFalse=0. 비율 고정을 켜면 한 변만 줘도 나머지가 따라온다.
        shape_obj.LockAspectRatio = MSO_TRUE if lock_aspect else MSO_FALSE
    except Exception:  # noqa: BLE001
        pass
    if width is not None:
        shape_obj.Width = _to_points(width, unit)
    if height is not None:
        shape_obj.Height = _to_points(height, unit)
    return (
        f"크기 변경: 슬라이드 {slide} / '{shape_obj.Name}' (발표자료: {_doc_label(pres, path)})\n"
        f"  이전: {_fmt_len(old_w)} × {_fmt_len(old_h)}\n"
        f"  이후: {_fmt_len(shape_obj.Width)} × {_fmt_len(shape_obj.Height)}"
        + (f"  (비율 고정)" if lock_aspect else "") + "\n"
        "아직 저장하지 않았습니다 — 파일에 반영하려면 save_presentation을 호출하세요."
    )


@mcp.tool()
@office_tool
def set_shape_position(
    slide: int, shape: str, left: float | None = None, top: float | None = None,
    unit: str = "pt", path: str = ""
) -> str:
    """🟡 한 도형의 위치(left·top)를 바꿉니다 (저장하지 않음).

    슬라이드 왼쪽 위 모서리 기준의 좌표로 도형을 옮깁니다. left/top 중 준 것만
    바꿉니다. 디스크에는 쓰지 않으며(저장은 save_presentation), Ctrl+Z로 되돌립니다.

    Args:
        slide: 슬라이드 번호(1부터).
        shape: 도형 번호 또는 이름(list_slide_shapes로 확인).
        left: 새 왼쪽 좌표. 생략하면 그대로.
        top: 새 위쪽 좌표. 생략하면 그대로.
        unit: left/top의 단위. 'pt'(기본)/'cm'/'mm'/'in'.
        path: 파일 경로. 생략하면 활성 발표자료. **PowerPoint에 열려 있어야 합니다.**

    Returns:
        바꾸기 전/후 위치와 저장 안내.
    """
    if left is None and top is None:
        return "left나 top 중 적어도 하나는 지정하세요."
    pres = _writable_presentation(path)
    slide_obj = _resolve_slide(pres, slide)
    shape_obj = _resolve_shape(slide_obj, shape)
    try:
        old_l, old_t = float(shape_obj.Left), float(shape_obj.Top)
    except Exception:  # noqa: BLE001
        raise OfficeError(f"'{shape_obj.Name}' 도형은 위치를 바꿀 수 없습니다(선/그룹 등일 수 있음).")
    if left is not None:
        shape_obj.Left = _to_points(left, unit)
    if top is not None:
        shape_obj.Top = _to_points(top, unit)
    return (
        f"위치 변경: 슬라이드 {slide} / '{shape_obj.Name}' (발표자료: {_doc_label(pres, path)})\n"
        f"  이전: left {_fmt_len(old_l)}, top {_fmt_len(old_t)}\n"
        f"  이후: left {_fmt_len(shape_obj.Left)}, top {_fmt_len(shape_obj.Top)}\n"
        "아직 저장하지 않았습니다 — 파일에 반영하려면 save_presentation을 호출하세요."
    )


@mcp.tool()
@office_tool
def format_shape(
    slide: int,
    shape: str,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    font_color: str = "",
    fill_color: str = "",
    align: str = "",
    path: str = "",
) -> str:
    """🟡 한 도형의 글꼴·색·정렬·채우기를 바꿉니다 (저장하지 않음).

    도형 전체 텍스트에 글꼴 크기/굵게/기울임/글자색/정렬을 적용하고, 도형 배경을
    fill_color로 채웁니다. 준 인자만 반영합니다. 디스크에는 쓰지 않습니다.

    Args:
        slide: 슬라이드 번호(1부터).
        shape: 도형 번호 또는 이름(list_slide_shapes로 확인).
        font_size: 글꼴 크기(pt). 생략하면 그대로.
        bold: 굵게 여부(True/False). 생략하면 그대로.
        italic: 기울임 여부(True/False). 생략하면 그대로.
        font_color: 글자색 '#RRGGBB'. 생략하면 그대로.
        fill_color: 도형 배경색 '#RRGGBB'. 생략하면 그대로.
        align: 문단 정렬 'left'/'center'/'right'/'justify'. 생략하면 그대로.
        path: 파일 경로. 생략하면 활성 발표자료. **PowerPoint에 열려 있어야 합니다.**

    Returns:
        적용한 서식 요약과 저장 안내.
    """
    pres = _writable_presentation(path)
    slide_obj = _resolve_slide(pres, slide)
    shape_obj = _resolve_shape(slide_obj, shape)
    applied = []

    if fill_color:
        try:
            shape_obj.Fill.Solid()
            shape_obj.Fill.ForeColor.RGB = _hex_to_ole(fill_color)
            applied.append(f"채우기 {fill_color}")
        except OfficeError:
            raise
        except Exception:  # noqa: BLE001
            raise OfficeError(f"'{shape_obj.Name}' 도형에 채우기 색을 적용할 수 없습니다.")

    text_opts = (font_size is not None or bold is not None or italic is not None
                 or font_color or align)
    if text_opts:
        try:
            has_tf = bool(shape_obj.HasTextFrame)
        except Exception:  # noqa: BLE001
            has_tf = False
        if not has_tf:
            raise OfficeError(
                f"'{shape_obj.Name}' 도형은 텍스트가 없어 글꼴/정렬을 바꿀 수 없습니다."
            )
        tr = shape_obj.TextFrame.TextRange
        font = tr.Font
        if font_size is not None:
            font.Size = float(font_size)
            applied.append(f"크기 {font_size}pt")
        if bold is not None:
            font.Bold = MSO_TRUE if bold else MSO_FALSE
            applied.append("굵게" if bold else "굵게 해제")
        if italic is not None:
            font.Italic = MSO_TRUE if italic else MSO_FALSE
            applied.append("기울임" if italic else "기울임 해제")
        if font_color:
            font.Color.RGB = _hex_to_ole(font_color)
            applied.append(f"글자색 {font_color}")
        if align:
            a = align.strip().lower()
            if a not in PP_ALIGN:
                raise OfficeError(f"align은 {', '.join(PP_ALIGN)} 중 하나여야 합니다(받은 값: '{align}').")
            tr.ParagraphFormat.Alignment = PP_ALIGN[a]
            applied.append(f"정렬 {a}")

    if not applied:
        return "바꿀 서식 인자를 하나도 지정하지 않았습니다(font_size/bold/italic/font_color/fill_color/align)."
    return (
        f"서식 적용: 슬라이드 {slide} / '{shape_obj.Name}' (발표자료: {_doc_label(pres, path)})\n"
        f"  적용: {', '.join(applied)}\n"
        "아직 저장하지 않았습니다 — 파일에 반영하려면 save_presentation을 호출하세요."
    )


def _report_new_shape(shape, slide: int, pres, path: str, kind: str) -> str:
    """add_* 도구가 만든 도형을 번호·이름·위치·크기로 요약해 돌려준다."""
    try:
        num = shape.Parent.Shapes.Count  # 새로 추가된 도형은 맨 끝
    except Exception:  # noqa: BLE001
        num = "?"
    try:
        pos = f"위치({_fmt_len(shape.Left)}, {_fmt_len(shape.Top)}), 크기({_fmt_len(shape.Width)} × {_fmt_len(shape.Height)})"
    except Exception:  # noqa: BLE001
        pos = ""
    return (
        f"{kind} 추가: 슬라이드 {slide} / '{shape.Name}' [{num}번]  "
        f"(발표자료: {_doc_label(pres, path)})\n"
        + (f"  {pos}\n" if pos else "")
        + "아직 저장하지 않았습니다 — 파일에 반영하려면 save_presentation을 호출하세요."
    )


@mcp.tool()
@office_tool
def add_text_box(
    slide: int,
    text: str,
    left: float = 50,
    top: float = 50,
    width: float = 300,
    height: float = 60,
    unit: str = "pt",
    font_size: float | None = None,
    bold: bool | None = None,
    font_color: str = "",
    path: str = "",
) -> str:
    """🟡 슬라이드에 새 텍스트 상자를 만들고 글을 넣습니다 (저장하지 않음).

    자유롭게 배치할 설명·주석·라벨을 추가할 때 씁니다. 위치·크기는 unit 단위,
    글꼴 옵션은 선택입니다. 디스크에는 쓰지 않으며(저장은 save_presentation),
    Ctrl+Z로 되돌릴 수 있습니다.

    Args:
        slide: 슬라이드 번호(1부터).
        text: 넣을 텍스트. 여러 줄은 개행(\\n)으로 구분합니다.
        left: 왼쪽 좌표(기본 50). top: 위쪽 좌표(기본 50).
        width: 너비(기본 300). height: 높이(기본 60).
        unit: 위 좌표·크기의 단위. 'pt'(기본)/'cm'/'mm'/'in'.
        font_size: 글꼴 크기(pt). 생략하면 기본값.
        bold: 굵게 여부. 생략하면 기본값.
        font_color: 글자색 '#RRGGBB'. 생략하면 기본값.
        path: 파일 경로. 생략하면 활성 발표자료. **PowerPoint에 열려 있어야 합니다.**

    Returns:
        새로 만든 텍스트 상자의 번호·이름·위치·크기와 저장 안내.
    """
    pres = _writable_presentation(path)
    slide_obj = _resolve_slide(pres, slide)
    box = slide_obj.Shapes.AddTextbox(
        MSO_TEXT_HORIZONTAL,
        _to_points(left, unit), _to_points(top, unit),
        _to_points(width, unit), _to_points(height, unit),
    )
    tr = box.TextFrame.TextRange
    tr.Text = str(text).replace("\r\n", "\n").replace("\n", "\r")
    if font_size is not None:
        tr.Font.Size = float(font_size)
    if bold is not None:
        tr.Font.Bold = MSO_TRUE if bold else MSO_FALSE
    if font_color:
        tr.Font.Color.RGB = _hex_to_ole(font_color)
    return _report_new_shape(box, slide, pres, path, "텍스트 상자")


@mcp.tool()
@office_tool
def add_shape(
    slide: int,
    shape_type: str,
    left: float = 50,
    top: float = 50,
    width: float = 150,
    height: float = 100,
    unit: str = "pt",
    text: str = "",
    fill_color: str = "",
    path: str = "",
) -> str:
    """🟡 슬라이드에 도형(사각형·원·화살표 등)을 만듭니다 (저장하지 않음).

    다이어그램 블록·강조 도형·화살표 등을 그릴 때 씁니다. 도형 안에 text를 넣거나
    fill_color로 배경을 칠할 수 있습니다. 디스크에는 쓰지 않습니다(저장은
    save_presentation).

    Args:
        slide: 슬라이드 번호(1부터).
        shape_type: 도형 종류. 예: 'rectangle', 'rounded_rectangle', 'oval',
            'triangle', 'diamond', 'right_arrow', 'star' 등(한글 별칭도 가능:
            '사각형','타원','화살표' 등). 지원 목록은 잘못된 값을 주면 안내됩니다.
        left: 왼쪽 좌표(기본 50). top: 위쪽 좌표(기본 50).
        width: 너비(기본 150). height: 높이(기본 100).
        unit: 위 좌표·크기의 단위. 'pt'(기본)/'cm'/'mm'/'in'.
        text: 도형 안에 넣을 텍스트(선택).
        fill_color: 배경색 '#RRGGBB'(선택).
        path: 파일 경로. 생략하면 활성 발표자료. **PowerPoint에 열려 있어야 합니다.**

    Returns:
        새로 만든 도형의 번호·이름·위치·크기와 저장 안내.
    """
    key = str(shape_type).strip().lower()
    if key not in PPT_AUTO_SHAPES:
        valid = ", ".join(sorted(k for k in PPT_AUTO_SHAPES if k.isascii()))
        raise OfficeError(f"shape_type '{shape_type}'을(를) 모릅니다. 지원: {valid}")
    pres = _writable_presentation(path)
    slide_obj = _resolve_slide(pres, slide)
    try:
        shape = slide_obj.Shapes.AddShape(
            PPT_AUTO_SHAPES[key],
            _to_points(left, unit), _to_points(top, unit),
            _to_points(width, unit), _to_points(height, unit),
        )
    except pythoncom.com_error as e:
        raise OfficeError(f"'{shape_type}' 도형을 만들지 못했습니다: {_com_message(e)}")
    if text:
        try:
            shape.TextFrame.TextRange.Text = str(text).replace("\r\n", "\n").replace("\n", "\r")
        except Exception:  # noqa: BLE001
            pass
    if fill_color:
        try:
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = _hex_to_ole(fill_color)
        except OfficeError:
            raise
        except Exception:  # noqa: BLE001
            pass
    return _report_new_shape(shape, slide, pres, path, f"도형({key})")


@mcp.tool()
@office_tool
def add_slide(path: str = "", layout: str = "blank", index: int = 0, title: str = "") -> str:
    """🟡 발표자료에 새 슬라이드를 추가합니다 (저장하지 않음).

    빈/제목만/제목+내용 등 레이아웃을 골라 슬라이드를 끼워 넣습니다. 디스크에는
    쓰지 않으며(저장은 save_presentation), Ctrl+Z로 되돌릴 수 있습니다.

    Args:
        path: 파일 경로. 생략하면 활성 발표자료. **PowerPoint에 열려 있어야 합니다.**
        layout: 레이아웃. 'blank'(기본)/'title'/'title_only'/'text'/'two_column'/'object'
            (한글 별칭: '빈','제목','제목만','제목과내용' 등).
        index: 삽입 위치(1부터). 0(기본)이면 맨 끝에 추가합니다.
        title: 제목 개체틀이 있으면 채울 제목 텍스트(선택).

    Returns:
        새 슬라이드 번호·레이아웃과 저장 안내.
    """
    key = str(layout).strip().lower()
    if key not in PPT_SLIDE_LAYOUTS:
        valid = ", ".join(sorted(k for k in PPT_SLIDE_LAYOUTS if k.isascii()))
        raise OfficeError(f"layout '{layout}'을(를) 모릅니다. 지원: {valid}")
    pres = _writable_presentation(path)
    total = pres.Slides.Count
    pos = total + 1 if not index or int(index) < 1 else min(int(index), total + 1)
    slide_obj = pres.Slides.Add(pos, PPT_SLIDE_LAYOUTS[key])
    if title:
        try:
            if slide_obj.Shapes.HasTitle:
                slide_obj.Shapes.Title.TextFrame.TextRange.Text = str(title)
        except Exception:  # noqa: BLE001
            pass
    return (
        f"슬라이드 추가: {pos}번 (레이아웃 {key}) — 이제 전체 {pres.Slides.Count}장 "
        f"(발표자료: {_doc_label(pres, path)})\n"
        + (f"  제목: {title}\n" if title else "")
        + "아직 저장하지 않았습니다 — 파일에 반영하려면 save_presentation을 호출하세요."
    )


@mcp.tool()
@office_tool
def delete_shape(slide: int, shape: str, path: str = "") -> str:
    """🟡 슬라이드에서 도형 하나를 지웁니다 (저장하지 않음).

    잘못 넣었거나 필요 없는 도형을 제거할 때 씁니다. 디스크에는 쓰지 않으므로
    (저장은 save_presentation) 저장 전 PowerPoint에서 Ctrl+Z로 되돌릴 수 있습니다.
    저장까지 하면 되돌리기 어려우니 주의하세요.

    Args:
        slide: 슬라이드 번호(1부터).
        shape: 지울 도형의 번호 또는 이름(list_slide_shapes로 확인).
        path: 파일 경로. 생략하면 활성 발표자료. **PowerPoint에 열려 있어야 합니다.**

    Returns:
        지운 도형 요약과 저장 안내.
    """
    pres = _writable_presentation(path)
    slide_obj = _resolve_slide(pres, slide)
    shape_obj = _resolve_shape(slide_obj, shape)
    name = shape_obj.Name
    kind = _shape_kind(shape_obj)
    shape_obj.Delete()
    return (
        f"도형 삭제: 슬라이드 {slide} / '{name}' ({kind})  (발표자료: {_doc_label(pres, path)})\n"
        "아직 저장하지 않았습니다 — 파일에 반영하려면 save_presentation을, 되돌리려면 "
        "PowerPoint에서 Ctrl+Z를 쓰세요."
    )


@mcp.tool()
@office_tool
def save_presentation(path: str = "", confirm: bool = False) -> str:
    """🔴 열려 있는 발표자료를 현재 경로에 저장합니다(덮어쓰기). (confirm=True 필요)

    set_powerpoint_text/add_shape 등으로 바꾼 내용을 디스크에 반영하는 단계입니다.
    confirm 없이 부르면 어떤 파일을 덮어쓸지 프리뷰만 돌려줍니다.

    Args:
        path: 파일 경로. 생략하면 활성 발표자료. PowerPoint에 열려 있어야 합니다.
        confirm: 실제 저장하려면 True. 없으면 프리뷰만.
    """
    pres = _writable_presentation(path)
    try:
        full = pres.FullName
    except Exception:  # noqa: BLE001
        full = ""
    if not full or not os.path.isabs(str(full).replace("/", os.sep)):
        return (
            "이 발표자료는 아직 디스크에 저장된 적이 없어 경로가 없습니다. "
            "PowerPoint에서 먼저 '다른 이름으로 저장'해 경로를 정하세요."
        )
    changed = ""
    try:
        changed = "변경 없음(이미 저장됨)" if bool(pres.Saved) else "변경 있음(미저장)"
    except Exception:  # noqa: BLE001
        pass
    if not confirm:
        details = [f"발표자료: {pres.Name}", f"경로: {full} (덮어쓰기)"]
        if changed:
            details.append(f"상태: {changed}")
        return _preview("발표자료 저장(덮어쓰기)", details, "(save_presentation ... confirm=true)")
    pres.Save()
    return f"저장 완료.\n  경로: {full}"


if __name__ == "__main__":
    import argparse
    import sys

    parser = argparse.ArgumentParser(description="Office 문서 읽기 MCP 서버")
    parser.add_argument(
        "--transport",
        choices=["stdio", "http", "sse"],
        default=os.getenv("OFFICE_MCP_TRANSPORT", "stdio"),
        help=(
            "stdio(기본): Claude 등 로컬 클라이언트가 프로세스를 직접 실행해 붙는다. "
            "http: n8n 등 네트워크 클라이언트가 URL로 접속한다(권장). "
            "sse: 구버전 n8n MCP 노드가 Streamable HTTP를 못 쓸 때 사용한다."
        ),
    )
    parser.add_argument("--host", default=os.getenv("OFFICE_MCP_HOST", "127.0.0.1"))
    parser.add_argument(
        "--port", type=int, default=int(os.getenv("OFFICE_MCP_PORT", "8087"))
    )
    args = parser.parse_args()

    if not COM_AVAILABLE:
        print(f"경고: pywin32를 불러올 수 없습니다 ({COM_IMPORT_ERROR}).", file=sys.stderr)
        print("서버는 실행되지만 모든 도구가 안내 메시지만 반환합니다.", file=sys.stderr)

    if args.transport in ("http", "sse"):
        # 네트워크 접속용. n8n의 MCP Client Tool 노드가 이 URL로 붙는다.
        # COM 특성상, 반드시 사용자가 로그인한 그 세션에서 실행해야
        # 열린 문서/활성 Office가 보인다(서비스나 다른 세션에서는 안 보임).
        path = "/mcp/" if args.transport == "http" else "/sse/"
        url = f"http://{args.host}:{args.port}{path}"
        print(f"Office MCP 서버 시작 ({args.transport}) — {url}", file=sys.stderr)
        mcp.run(transport=args.transport, host=args.host, port=args.port)
    else:
        # stdio: stdout이 MCP 프로토콜 채널이므로 로그는 stderr로 보낸다.
        print("Office MCP 서버 시작 (stdio)", file=sys.stderr)
        mcp.run(transport="stdio")
