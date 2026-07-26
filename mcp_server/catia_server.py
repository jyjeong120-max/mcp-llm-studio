"""catia_server.py

CATIA V5 문서를 읽고(🟢) 수정·저장·내보내기(🔴)하는 MCP 서버입니다.
pywin32(COM)로 **이미 실행 중인 CATIA V5 세션에 붙어**, 화면에 열려 있는
Part/Product/Drawing을 그대로 다룹니다. office_server.py / outlook_server.py와
같은 계열이며, 규약도 동일합니다.

전제
    - Windows + CATIA **V5** 설치. (3DEXPERIENCE/V6는 자동화 모델이 달라 이 서버로는
      동작하지 않을 수 있습니다.)
    - 반드시 **사용자가 로그인한 그 세션**에서 실행해야 합니다(COM 제약). 서비스나
      다른 세션에서 띄우면 열린 CATIA가 보이지 않습니다.
    - 이 서버는 **새 CATIA를 띄우지 않고** 이미 떠 있는 인스턴스에만 붙습니다
      (GetActiveObject). CATIA가 꺼져 있으면 각 도구가 안내 메시지를 돌려줍니다.
    - pywin32가 없거나(비-Windows) CATIA가 없으면 서버는 그대로 뜨고, 도구들이
      실패 사유를 담은 안내를 반환합니다(우아한 저하 — import 에러로 죽지 않음).

안전 등급 (outlook_server.py와 동일한 3티어)
    🟢 읽기 — 열린 문서 목록/파라미터 조회/제품(BOM) 트리/부품 요약/스케치 요소 조회/
       현재 화면 선택 조회(get_selection)
       (list_sketch_geometry)/3D 요소·솔리드 피처 조회(list_3d_geometry)/측정
       (measure_element·measure_body)
    🟡 로컬 생성(비파괴, 되돌리기 쉬움) — CATIA 실행(launch_catia), 새 파트(new_part),
       문서 열기(open_document), 스케치 열기(new_sketch), 스케치 요소 그리기
       (sketch_point/line/circle/arc/rectangle/spline), 스케치 구속(sketch_coincidence/
       perpendicular/parallel/horizontal/vertical)과 치수 구속(sketch_dimension),
       3D 와이어프레임(point_3d/line_3d/spline_3d/plane_offset)·로프트(create_loft),
       솔리드 피처(pad/pocket/shaft/groove/mirror_body/fillet·chamfer_selected_edges/
       shell_selected_faces), 어셈블리(new_product/add_component/move_component),
       뷰·유틸(undo/redo/set_visibility/set_color/capture_view). CATIA의 Undo로 되돌릴 수
       있고 디스크에 쓰지 않으므로 confirm 없이 실행한다. (예외: capture_view는 새
       이미지 파일 하나를 만들지만 기존 파일을 절대 덮어쓰지 않아 🟡로 둔다.)
       그린 요소는 이름을 반환/지정할 수 있고, 그 이름으로 구속에서
       지목한다(선 끝점은 '이름.start'/'이름.end', 원 중심은 '이름.center').
    🔴 파괴·디스크 기록·되돌리기 어려움 — set_parameter, 요소 삭제(delete_element),
       save_document, export_document, 그리고 CATIA 전체 종료(quit_catia). 반드시
       confirm=True를 받아야 실행되며,
       confirm 없이 호출하면 무엇을 할지 요약한 '실행 전 확인' 프리뷰만 돌려준다
       (_preview). 클라이언트가 HumanInTheLoopMiddleware의 INTERRUPT_ON에 같은 이름을
       올리면 이중 안전장치가 된다.

path 인자 규칙 (문서를 받는 모든 도구 공통)
    path=None  -> CATIA에서 지금 활성화된 문서(ActiveDocument)를 다룹니다.
    path 지정  -> 이미 열려 있으면 그 문서를, 아니면 CATIA에서 열어(작업 후) 닫습니다.
                  (우리가 연 문서는 저장하지 않는 한 닫을 때 변경이 버려집니다.)
"""

from __future__ import annotations

import json
import math
import os
import re
import sys
from contextlib import contextmanager
from functools import wraps

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
    name="catia",
    instructions=(
        "CATIA V5 문서를 읽고 수정하며 파트를 설계하는 MCP 서버입니다. path를 생략하면 "
        "지금 활성화된 문서를 다룹니다. 어떤 문서가 열려 있는지 모르면 list_open_documents를 "
        "먼저 호출하세요. 설계 흐름: new_part(새 파트) 또는 open_document(파일 열기) → "
        "new_sketch(스케치 열기) → sketch_point/line/circle/spline(그리기, name으로 이름 지정 "
        "권장; 원 여러 개는 sketch_circles로 한 번에) → "
        "sketch_coincidence/perpendicular/parallel/horizontal/vertical(형상 구속)·"
        "sketch_dimension(치수 구속: 길이/거리/각도/반지름 값 지정) → "
        "pad(돌출)/pocket(파냄)/shaft·groove(회전 — axis로 스케치 안 축 선 지정)로 솔리드화. "
        "모따기는 chamfer_selected_edges, 속 비우기는 shell_selected_faces, 대칭 복제는 "
        "mirror_body. 어셈블리는 new_product → add_component(파일 추가) → move_component"
        "(이동). 형상 검증은 measure_element/measure_body(길이·면적·부피·질량), 결과 확인은 "
        "capture_view(화면 캡처), 실수 복구는 undo(재적용은 redo). 사용자가 화면에서 "
        "선택해 둔 항목은 get_selection으로 확인. "
        "3D 와이어프레임은 point_3d/line_3d/spline_3d, 평면 띄우기는 plane_offset, 여러 "
        "단면을 잇는 로프트는 create_loft(단면 스케치들을 plane_offset 평면 위에 만들어 "
        "이름으로 지목), 모서리 라운드는 fillet_selected_edges, 모따기는 "
        "chamfer_selected_edges. 모서리를 클릭 없이 고르려면 list_edges로 모서리"
        "(길이·중점)를 확인해 select_edge(index)로 잡은 뒤 이 도구들을 부른다(면은 "
        "list_faces/select_face). 어떤 요소가 있는지 헷갈리면 기억 대신 list_sketch_geometry(스케치 안)/"
        "list_3d_geometry(3D 요소·피처)로 조회하세요. 여러 스케치를 만들었으면 "
        "list_sketches로 지금 있는 스케치 이름을 확인하고, pad/pocket에 sketch= 로 대상 "
        "스케치 이름을 반드시 지정하세요(비우면 가장 최근 "
        "스케치가 잡혀 엉뚱한 프로파일을 뭅니다). 돌출·파냄 방향은 스케치 평면 법선을 "
        "따르며(xy→+Z, yz→+X, zx→+Y), 반대로 하려면 reverse=True를 씁니다. 특정 면에 "
        "스케치하거나 면을 뚫을 때 사용자 클릭 없이 하려면 list_faces로 면(면적·법선·중심)을 "
        "확인해 원하는 면을 고르고 select_face(index)로 잡은 뒤 new_sketch(plane='selection')/"
        "shell_selected_faces를 부릅니다. 파라미터 변경"
        "(set_parameter)·요소 삭제(delete_element)·저장(save_document)·내보내기"
        "(export_document)·CATIA 종료"
        "(quit_catia)는 되돌리기 어려워 confirm=True가 있어야 실행되며, confirm 없이 부르면 "
        "무엇을 할지 프리뷰만 돌려줍니다. 파트 생성·스케치·그리기는 비파괴(Undo 가능)라 "
        "confirm 없이 바로 실행됩니다."
    ),
)

# 출력이 컨텍스트를 통째로 삼키지 않도록 하는 기본 상한. 도구 인자로 조정할 수 있다.
MAX_ITEMS = 300      # 파라미터/노드 최대 표시 개수
MAX_TREE_NODES = 500  # 제품 트리 최대 노드 수
MAX_CHARS = 20000

# ExportData가 받는 형식 문자열 → 사람이 읽는 설명. CATIA가 파일명에 확장자를 붙인다.
# (설치된 번역기/라이선스에 따라 지원 형식이 다를 수 있다 — 없으면 CATIA가 오류를 낸다.)
EXPORT_FORMATS = {
    "stp": "STEP", "step": "STEP", "igs": "IGES", "iges": "IGES",
    "stl": "STL", "wrl": "VRML", "cgr": "CGR", "model": "CATIA V4 model",
    "3dxml": "3D XML",
}


class CatiaError(Exception):
    """도구가 사용자에게 그대로 돌려줄 안내 메시지를 담은 예외."""


def _require_com():
    if not COM_AVAILABLE:
        raise CatiaError(
            "CATIA에 접근하지 못했습니다. pywin32를 불러올 수 없습니다"
            f"({COM_IMPORT_ERROR}). Windows에서 CATIA V5가 설치·실행된 환경에서 "
            "`pip install pywin32` 후 다시 실행하세요."
        )


def catia_tool(fn):
    """COM 초기화와 예외 처리를 감싸는 도구 데코레이터.

    FastMCP는 동기 도구를 워커 스레드에서 실행한다. COM은 스레드마다 CoInitialize가
    필요하므로 매 호출마다 초기화하고 해제한다. CatiaError는 안내 메시지로, 나머지
    COM 오류는 사유를 붙여 반환한다(도구가 예외로 죽지 않고 항상 문자열을 돌려준다).
    """

    @wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            _require_com()
        except CatiaError as e:
            return str(e)

        pythoncom.CoInitialize()
        try:
            return fn(*args, **kwargs)
        except CatiaError as e:
            return str(e)
        except pythoncom.com_error as e:
            return f"CATIA 호출이 실패했습니다: {_com_message(e)}"
        except Exception as e:  # noqa: BLE001 — 모든 오류를 안내 문자열로 돌려준다
            return f"작업에 실패했습니다: {type(e).__name__}: {e}"
        finally:
            pythoncom.CoUninitialize()

    return wrapper


def _com_message(e) -> str:
    """com_error에서 사람이 읽을 만한 설명만 뽑아낸다."""
    try:
        info = e.excepinfo
        if info and len(info) > 2 and info[2]:
            return str(info[2]).strip()
    except Exception:
        pass
    return str(e)


# ─────────────────────────────── 연결/문서 헬퍼 ───────────────────────────────


def _catia():
    """실행 중인 CATIA V5 인스턴스에 붙는다. 없으면 안내와 함께 실패.

    Dispatch가 아니라 GetActiveObject를 쓴다 — CATIA를 새로 띄우면 라이선스를
    잡고 빈 세션이 떠서 사용자가 보는 문서와 어긋나기 때문이다. 이미 켜져 있는
    그 세션에만 붙는다.
    """
    try:
        return win32com.client.GetActiveObject("CATIA.Application")
    except (pythoncom.com_error, AttributeError):
        raise CatiaError(
            "CATIA V5가 실행되어 있지 않습니다. CATIA를 켜고 문서를 연 뒤 다시 "
            "시도하세요. (이 서버는 새 CATIA를 띄우지 않고 이미 열린 세션에 붙습니다. "
            "3DEXPERIENCE/V6에서는 동작하지 않을 수 있습니다.)"
        )


def _doc_type(doc) -> str:
    """문서 종류를 이름 확장자로 판별한다 (COM 타입명은 pywin32로 얻기 번거롭다)."""
    name = ""
    try:
        name = (doc.Name or "").lower()
    except Exception:
        pass
    if name.endswith(".catpart"):
        return "Part"
    if name.endswith(".catproduct"):
        return "Product"
    if name.endswith(".catdrawing"):
        return "Drawing"
    return "기타"


def _full_name(doc) -> str:
    try:
        return doc.FullName or ""
    except Exception:
        return ""


def _norm(p: str) -> str:
    return os.path.normcase(os.path.abspath(os.path.expanduser(p))) if p else ""


def _find_open_doc(catia, path: str):
    """지정 경로의 문서가 이미 CATIA에 열려 있으면 그 COM 객체를 돌려준다.

    전체 경로 정확 일치를 먼저 찾는다. 그게 없을 때만 파일 이름 일치로 폴백하되,
    호출자가 **절대경로를 명시했으면 폴백하지 않는다** — 다른 폴더의 동명 파일을
    엉뚱하게 잡아 저장·수정하는 사고를 막기 위해서다(save_document 등이 이 함수를 쓴다).
    사용자가 파일 이름만 넘긴 경우(절대경로 아님)에만 이름 일치로 열린 문서를 찾는다.
    """
    if not path:
        return None
    docs = catia.Documents
    try:
        count = docs.Count
    except pythoncom.com_error:
        return None
    target_full = _norm(path)
    base_target = os.path.normcase(os.path.basename(path))
    is_abs = os.path.isabs(os.path.expanduser(path))
    basename_match = None
    for i in range(1, count + 1):
        try:
            doc = docs.Item(i)
        except pythoncom.com_error:
            continue
        full = _full_name(doc)
        if full and _norm(full) == target_full:
            return doc  # 전체 경로 정확 일치 — 최우선
        if basename_match is None and base_target:
            fb = os.path.normcase(os.path.basename(full)) if full else ""
            nb = os.path.normcase(os.path.basename(_safe(doc, "Name")))
            if base_target in (fb, nb):
                basename_match = doc
    # 절대경로를 줬는데 그 경로가 안 열려 있으면, 동명 파일로 착각하지 않고 None을
    # 돌려준다(호출자가 그 경로의 파일을 직접 열게 한다).
    return None if is_abs else basename_match


def _active_document(catia):
    """CATIA에서 지금 활성화된 문서를 돌려준다. 없으면 안내와 함께 실패."""
    try:
        if catia.Documents.Count == 0:
            raise CatiaError(
                "CATIA는 실행 중이지만 열린 문서가 없습니다. 문서를 열거나 path 인자에 "
                "파일 경로를 지정하세요."
            )
        return catia.ActiveDocument
    except pythoncom.com_error as e:
        raise CatiaError(f"활성 문서를 가져오지 못했습니다: {_com_message(e)}")


@contextmanager
def _document(catia, path):
    """작업할 문서를 확보한다.

    path 없음 -> 활성 문서. path 있고 이미 열려 있음 -> 그 문서(사용자 세션).
    path 있고 안 열려 있음 -> CATIA에서 열고, 블록이 끝나면 닫는다(저장하지 않으면
    변경은 버려진다). 우리가 연 문서만 닫는다 — 사용자가 열어 둔 문서는 건드리지 않는다.
    """
    if not path:
        yield _active_document(catia)
        return

    opened = _find_open_doc(catia, path)
    if opened is not None:
        yield opened  # 사용자 세션에 열려 있는 문서 — 그대로 다룬다
        return

    p = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(p):
        raise CatiaError(f"'{p}' 경로에 파일이 없고 CATIA에도 열려 있지 않습니다. 경로를 확인하세요.")
    doc = catia.Documents.Open(p)
    try:
        yield doc
    finally:
        try:
            doc.Close()
        except Exception:
            pass


def _safe(obj, attr: str) -> str:
    """COM 속성을 안전하게 문자열로 읽는다. 실패하면 빈 문자열."""
    try:
        v = getattr(obj, attr)
        return "" if v is None else str(v).strip()
    except Exception:
        return ""


def _doc_label(doc) -> str:
    name = _safe(doc, "Name") or "(이름 없음)"
    full = _full_name(doc)
    kind = _doc_type(doc)
    return f"{name} [{kind}]" + (f"  ({full})" if full and full != name else "")


def _truncate(text: str, limit: int = MAX_CHARS) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + f"\n\n…(전체 {len(text):,}자 중 {limit:,}자만 표시)"


def _preview(action: str, details: list[str], tool_hint: str = "") -> str:
    """🔴 도구가 confirm 없이 호출됐을 때 돌려줄 '실행 전 확인' 프리뷰.

    실제 동작은 하지 않고 무엇을 할지 요약만 보여준다. 사용자가 확인한 뒤 같은
    도구를 confirm=True로 다시 부르면 그때 실행된다.
    """
    lines = [f"⚠️ 승인 필요 — 아직 실행하지 않았습니다: {action}", ""]
    lines.extend(f"  {d}" for d in details)
    lines.append("")
    lines.append("이대로 진행하려면 같은 도구를 confirm=true 로 다시 호출하세요. " + (tool_hint or ""))
    return "\n".join(lines).rstrip()


# ─────────────────────────────── 파라미터 헬퍼 ───────────────────────────────


def _parameter_container(doc):
    """Part/Product의 파라미터 컬렉션을 돌려준다. Drawing 등은 지원하지 않는다."""
    kind = _doc_type(doc)
    if kind == "Part":
        container = doc.Part
    elif kind == "Product":
        container = doc.Product
    else:
        raise CatiaError(
            f"이 문서({kind})에는 파라미터 트리가 없습니다. Part(.CATPart) 또는 "
            "Product(.CATProduct)에서만 파라미터를 다룰 수 있습니다."
        )
    try:
        return container.Parameters
    except pythoncom.com_error as e:
        raise CatiaError(f"파라미터 컬렉션을 열지 못했습니다: {_com_message(e)}")


def _param_value(p) -> str:
    """파라미터 값을 문자열로 읽는다. 길이/각도 등은 단위 포함(ValueAsString) 우선."""
    try:
        vas = p.ValueAsString
        return str(vas() if callable(vas) else vas)
    except Exception:
        pass
    try:
        return str(p.Value)
    except Exception:
        return "(값을 읽을 수 없음)"


def _update(doc) -> None:
    """파라미터 변경 후 형상을 재생성한다. 실패해도 조용히 넘어간다(값은 이미 반영됨)."""
    try:
        kind = _doc_type(doc)
        if kind == "Part":
            doc.Part.Update()
        elif kind == "Product":
            doc.Product.Update()
    except Exception:
        pass


# ─────────────────────────────── 파트·스케치 헬퍼 ───────────────────────────────

# 기준 평면 이름 → OriginElements 속성. 스케치를 얹을 기본 평면들.
_PLANES = {"xy": "PlaneXY", "yz": "PlaneYZ", "zx": "PlaneZX"}

# 기준 평면 → (스케치 2D 좌표가 매핑되는 3D 축, 돌출/파냄 법선 방향).
# pad/pocket은 스케치 평면의 법선 방향으로 재료를 만들거나 파낸다 — 모델이 어느
# 평면에 그려야 형상이 원하는 방향으로 서는지 판단하도록 도구 응답에 실어 보낸다.
_PLANE_AXES = {
    "xy": ("(x,y)→(X,Y)", "+Z"),
    "yz": ("(x,y)→(Y,Z)", "+X"),
    "zx": ("(x,y)→(Z,X)", "+Y"),
}


def _active_part(catia):
    """활성 문서의 Part를 돌려준다((doc, part)). Part가 아니면 안내와 함께 실패.

    스케치/그리기 도구는 '지금 활성화된 파트'에서 작업한다 — new_part나 open_document,
    또는 CATIA에서 사용자가 연 파트가 활성 상태여야 한다.
    """
    doc = _active_document(catia)
    if _doc_type(doc) != "Part":
        raise CatiaError(
            "활성 문서가 Part(.CATPart)가 아닙니다. new_part로 파트를 만들거나 CATIA에서 "
            "파트를 활성화한 뒤 다시 시도하세요."
        )
    try:
        return doc, doc.Part
    except pythoncom.com_error as e:
        raise CatiaError(f"활성 파트를 가져오지 못했습니다: {_com_message(e)}")


def _get_sketch(part, name: str = ""):
    """스케치를 돌려준다. name이 있으면 그 이름, 없으면 가장 최근 스케치(MainBody).

    스케치가 하나도 없으면 안내와 함께 실패한다(먼저 new_sketch로 열어야 한다).
    """
    try:
        sketches = part.MainBody.Sketches
        count = sketches.Count
    except pythoncom.com_error as e:
        raise CatiaError(f"스케치 컬렉션을 열지 못했습니다: {_com_message(e)}")
    if count == 0:
        raise CatiaError("스케치가 없습니다. new_sketch로 먼저 스케치를 여세요.")
    if name:
        try:
            return sketches.Item(name)
        except pythoncom.com_error:
            raise CatiaError(f"스케치 '{name}'를 찾지 못했습니다. active_document로 이름을 확인하세요.")
    return sketches.Item(count)  # 가장 최근에 만든 스케치


def _parse_points(points: str):
    """'[[0,0],[10,5]]'(JSON) 또는 '0,0; 10,5'(세미콜론 구분) 형식을 [(x,y),...]로 파싱."""
    s = (points or "").strip()
    pairs = []
    if s.startswith("["):
        for p in json.loads(s):
            pairs.append((float(p[0]), float(p[1])))
    else:
        for chunk in re.split(r"[;\n]+", s):
            chunk = chunk.strip()
            if not chunk:
                continue
            xy = re.split(r"[,\s]+", chunk)
            if len(xy) < 2:
                raise CatiaError(f"좌표 형식 오류: '{chunk}' (x,y 여야 합니다).")
            pairs.append((float(xy[0]), float(xy[1])))
    if len(pairs) < 2:
        raise CatiaError("스플라인은 점이 최소 2개 필요합니다.")
    return pairs


# ─────────────────────────────── 이름·참조·구속 헬퍼 ───────────────────────────────
# 요소를 우리가 맵으로 들지 않고 CATIA가 든 이름(GeometricElements)을 그대로 조회한다.
# 그래서 대화 기억에 의존하지 않고 list_sketch_geometry로 언제든 현재 상태를 재조회할 수 있다.

# 끝점 참조 접미사: "Line.1.end" → 그 선의 끝점, "Circle.1.center" → 원의 중심점.
_ENDPOINT_SUFFIX = ("start", "end", "center")

# 스케치 구속 종류 → CatConstraintType 정수값.
# ⚠ 이 정수들은 실기 검증 대상이다. 구속 생성 후 반환되는 CATIA 이름(예: "Perpendicularity.1")
# 으로 실제로 맞는 종류가 걸렸는지 바로 확인할 수 있다. 틀리면 이 표의 숫자만 고치면 된다.
_CST_TYPES = {
    "coincidence": 10, "concentric": 11, "tangent": 12,
    "parallel": 13, "perpendicular": 14,
    "horizontal": 15, "vertical": 16,
}

# 치수 구속 종류 → CatConstraintType 정수값 (위와 같은 열거의 치수 계열).
# ⚠ 실기 검증 대상 — 생성된 구속의 CATIA 이름(예: 'Length.1'/'Radius.1')으로 실제 걸린
# 종류를 확인할 수 있다. distance/angle은 두 요소(bi), length/radius는 단일 요소(mono).
_DIM_TYPES = {"distance": 1, "length": 2, "angle": 3, "radius": 5}


def _apply_name(elem, name: str) -> str:
    """요소에 이름을 지정하고(중복/불가 시 자동 이름 유지) 최종 이름을 돌려준다.

    이름을 직접 주면(name="A") 이후 그 이름으로 참조할 수 있어 대화가 auto 이름
    (Line.7 등)을 기억할 필요가 없다.
    """
    name = (name or "").strip()
    if name:
        try:
            elem.Name = name
        except Exception:  # noqa: BLE001 — 중복 등으로 실패하면 자동 이름을 쓴다
            pass
    return _safe(elem, "Name") or "(이름 없음)"


def _pt_coords(pt):
    """Point2D의 (x, y)를 best-effort로 읽는다. 못 읽으면 None."""
    try:
        c = pt.GetCoordinates()
        if isinstance(c, (list, tuple)) and len(c) >= 2:
            return (round(float(c[0]), 4), round(float(c[1]), 4))
    except Exception:  # noqa: BLE001 — 좌표 표시는 부가정보라 실패해도 넘어간다
        pass
    return None


def _describe_geo(elem) -> str:
    """스케치 요소 하나를 'Name [종류] 좌표'로 요약한다(좌표는 best-effort).

    종류는 COM 타입명을 직접 얻기 번거로워 속성 존재로 추정한다: 반지름→원, 끝점→선,
    아니면 점/기타. list_sketch_geometry가 이걸로 목록을 만든다.
    """
    name = _safe(elem, "Name")
    try:  # 원/호: 반지름이 있음
        r = round(float(elem.Radius), 4)
        try:
            c = _pt_coords(elem.CenterPoint)
        except Exception:  # noqa: BLE001
            c = None
        return f"{name} [원] {('중심' + str(c)) if c else ''} r={r}".replace("  ", " ").strip()
    except Exception:  # noqa: BLE001
        pass
    try:  # 선: 시작/끝점이 있음
        sp, ep = _pt_coords(elem.StartPoint), _pt_coords(elem.EndPoint)
        return f"{name} [선] {sp}-{ep}" if (sp and ep) else f"{name} [선]"
    except Exception:  # noqa: BLE001
        pass
    c = _pt_coords(elem)
    return f"{name} [점] {c}" if c is not None else f"{name} [기타]"


def _resolve_ref(part, sketch, token: str):
    """'baseLine' / 'Line.1' / 'baseLine.end' / 'Circle.1.center' 토큰을 Reference로 푼다.

    이름은 sketch.GeometricElements에서 찾고, .start/.end/.center 접미사는 그 요소의
    끝점/중심점으로 해석한다. (별도 맵 없이 CATIA가 든 이름을 그대로 조회.)
    """
    token = (token or "").strip()
    base, suffix = token, ""
    low = token.lower()
    for suf in _ENDPOINT_SUFFIX:
        if low.endswith("." + suf):
            base = token[: -(len(suf) + 1)]
            suffix = suf
            break
    try:
        elem = sketch.GeometricElements.Item(base)
    except pythoncom.com_error:
        raise CatiaError(
            f"스케치에서 요소 '{base}'를 찾지 못했습니다. list_sketch_geometry로 이름을 확인하세요."
        )
    target = elem
    try:
        if suffix == "start":
            target = elem.StartPoint
        elif suffix == "end":
            target = elem.EndPoint
        elif suffix == "center":
            target = elem.CenterPoint
    except pythoncom.com_error:
        raise CatiaError(f"'{base}'에 {suffix} 점이 없습니다(선은 .start/.end, 원은 .center).")
    return part.CreateReferenceFromObject(target)


def _add_constraint(kind: str, a: str, b: str, sketch_name: str) -> str:
    """스케치 구속을 만든다. b가 비면 단일 요소(mono), 있으면 두 요소(bi) 구속.

    반환: 생성된 구속의 CATIA 이름(예: 'Perpendicularity.1') — 실제 걸린 종류 확인용.
    """
    catia = _catia()
    _, part = _active_part(catia)
    sk = _get_sketch(part, sketch_name)
    ctype = _CST_TYPES[kind]
    # 구속은 스케치 편집 상태에서 part.Constraints에 추가한다(V5 자동화 관례).
    sk.OpenEdition()
    try:
        ref1 = _resolve_ref(part, sk, a)
        if b:
            ref2 = _resolve_ref(part, sk, b)
            cst = part.Constraints.AddBiEltCst(ctype, ref1, ref2)
        else:
            cst = part.Constraints.AddMonoEltCst(ctype, ref1)
    finally:
        sk.CloseEdition()
    part.Update()
    return _safe(cst, "Name") or "(이름 없음)"


# ════════════════════════════ 🟢 읽기 (비파괴) ════════════════════════════


@mcp.tool()
@catia_tool
def connection_status() -> str:
    """CATIA 연결 상태와 열린 문서 개수를 확인합니다. (🟢 읽기)

    Returns:
        CATIA 실행 여부, 열린 문서 수, 활성 문서 요약.
    """
    catia = _catia()
    try:
        count = catia.Documents.Count
    except pythoncom.com_error as e:
        return f"CATIA에 연결됐지만 문서 목록을 읽지 못했습니다: {_com_message(e)}"
    active = ""
    if count:
        try:
            active = _doc_label(catia.ActiveDocument)
        except Exception:
            active = "(활성 문서 없음)"
    return (
        "CATIA V5에 연결됨.\n"
        f"열린 문서: {count}개\n"
        f"활성 문서: {active or '(없음)'}"
    )


@mcp.tool()
@catia_tool
def list_open_documents() -> str:
    """지금 CATIA에 열려 있는 모든 문서의 이름·종류·경로를 나열합니다. (🟢 읽기)

    어떤 문서를 다룰지 정할 때 먼저 호출하세요. 여기서 얻은 이름/경로를 다른 도구의
    path 인자로 넘기면 그 문서를 지정할 수 있습니다.
    """
    catia = _catia()
    docs = catia.Documents
    count = docs.Count
    if count == 0:
        return "열려 있는 CATIA 문서가 없습니다. CATIA에서 문서를 여세요."
    active_name = ""
    try:
        active_name = _safe(catia.ActiveDocument, "Name")
    except Exception:
        pass
    out = [f"열린 문서 {count}개:"]
    for i in range(1, count + 1):
        doc = docs.Item(i)
        mark = " ← 활성" if _safe(doc, "Name") == active_name else ""
        out.append(f"  {i}. {_doc_label(doc)}{mark}")
    return "\n".join(out)


@mcp.tool()
@catia_tool
def active_document() -> str:
    """지금 활성화된 문서의 이름·종류·경로·수정 여부를 요약합니다. (🟢 읽기)"""
    catia = _catia()
    doc = _active_document(catia)
    saved = _safe(doc, "Saved")
    lines = [f"활성 문서: {_doc_label(doc)}"]
    if saved:
        lines.append(f"저장 상태: {'저장됨' if saved in ('True', '-1', '1') else '변경 있음(미저장)'}")
    return "\n".join(lines)


@mcp.tool()
@catia_tool
def list_parameters(path: str | None = None, filter: str = "", max_items: int = MAX_ITEMS) -> str:
    """Part/Product의 파라미터를 이름·값·설명과 함께 나열합니다. (🟢 읽기)

    Args:
        path: 대상 문서. 생략하면 활성 문서.
        filter: 이름에 이 문자열이 포함된 파라미터만(대소문자 무시). 비우면 전체.
        max_items: 최대 표시 개수(기본 300).

    Returns:
        `이름 = 값  (설명)` 목록.
    """
    catia = _catia()
    with _document(catia, path) as doc:
        params = _parameter_container(doc)
        total = params.Count
        if total == 0:
            return f"{_doc_label(doc)}\n\n파라미터가 없습니다."
        needle = filter.lower().strip()
        out = [f"{_doc_label(doc)}", f"파라미터 {total}개" + (f" (필터: '{filter}')" if needle else ""), ""]
        shown = 0
        for i in range(1, total + 1):
            if shown >= max_items:
                out.append(f"… (이하 생략, {max_items}개 상한 도달)")
                break
            p = params.Item(i)
            name = _safe(p, "Name")
            if needle and needle not in name.lower():
                continue
            comment = _safe(p, "Comment")
            line = f"- {name} = {_param_value(p)}"
            if comment:
                line += f"  ({comment})"
            out.append(line)
            shown += 1
        if shown == 0:
            out.append("(필터에 맞는 파라미터가 없습니다.)")
        return _truncate("\n".join(out))


@mcp.tool()
@catia_tool
def get_parameter(name: str, path: str | None = None) -> str:
    """파라미터 하나의 값을 이름으로 조회합니다. (🟢 읽기)

    Args:
        name: 파라미터 이름 (list_parameters로 확인한 정확한 이름).
        path: 대상 문서. 생략하면 활성 문서.
    """
    catia = _catia()
    with _document(catia, path) as doc:
        params = _parameter_container(doc)
        try:
            p = params.Item(name)
        except pythoncom.com_error:
            return (
                f"파라미터 '{name}'를 찾지 못했습니다. list_parameters로 정확한 이름을 "
                "확인하세요(대소문자·전체 경로가 정확해야 합니다)."
            )
        comment = _safe(p, "Comment")
        return f"{name} = {_param_value(p)}" + (f"  ({comment})" if comment else "")


@mcp.tool()
@catia_tool
def product_tree(path: str | None = None, max_depth: int = 10, max_nodes: int = MAX_TREE_NODES) -> str:
    """Product(어셈블리)의 제품 구조(BOM 트리)를 들여쓰기로 보여줍니다. (🟢 읽기)

    각 노드에 PartNumber · 명칭(Nomenclature) · 리비전 · 인스턴스 이름을 표시합니다.

    Args:
        path: 대상 문서. 생략하면 활성 문서. (Product가 아니면 안내를 돌려줍니다.)
        max_depth: 최대 탐색 깊이(기본 10).
        max_nodes: 최대 노드 수(기본 500).
    """
    catia = _catia()
    with _document(catia, path) as doc:
        if _doc_type(doc) != "Product":
            return (
                f"{_doc_label(doc)}\n\n이 문서는 Product(어셈블리)가 아니라 제품 트리가 "
                "없습니다. list_parameters나 part_summary를 쓰세요."
            )
        root = doc.Product
        rpn = _safe(root, "PartNumber") or _safe(root, "Name")
        lines = [f"{_doc_label(doc)}", "", f"■ {rpn}"]
        counter = [0]
        _walk_product(root, 1, max_depth, lines, counter, max_nodes)
        if counter[0] == 0:
            lines.append("  (하위 구성요소가 없습니다.)")
        else:
            lines.append("")
            lines.append(f"총 {counter[0]}개 구성요소(표시분).")
        return _truncate("\n".join(lines))


def _walk_product(prod, depth, max_depth, lines, counter, max_nodes):
    """제품 트리를 재귀로 훑어 lines에 채운다. 노드/깊이 상한을 지킨다."""
    try:
        children = prod.Products
        count = children.Count
    except Exception:
        return
    for i in range(1, count + 1):
        if counter[0] >= max_nodes:
            lines.append("  " * depth + "… (이하 생략, 노드 상한 도달)")
            return
        c = children.Item(i)
        counter[0] += 1
        pn = _safe(c, "PartNumber")
        nom = _safe(c, "Nomenclature")
        rev = _safe(c, "Revision")
        inst = _safe(c, "Name")
        label = pn or "(이름 없음)"
        if nom:
            label += f" · {nom}"
        if rev:
            label += f" · rev {rev}"
        if inst and inst != pn:
            label += f"  [{inst}]"
        lines.append("  " * depth + f"- {label}")
        if depth < max_depth:
            _walk_product(c, depth + 1, max_depth, lines, counter, max_nodes)


@mcp.tool()
@catia_tool
def part_summary(path: str | None = None) -> str:
    """Part 문서의 개요(바디 수, 메인 바디, 파라미터 수)를 요약합니다. (🟢 읽기)

    Args:
        path: 대상 문서. 생략하면 활성 문서. (Part가 아니면 안내를 돌려줍니다.)
    """
    catia = _catia()
    with _document(catia, path) as doc:
        if _doc_type(doc) != "Part":
            return f"{_doc_label(doc)}\n\n이 문서는 Part(.CATPart)가 아닙니다."
        part = doc.Part
        lines = [f"{_doc_label(doc)}", ""]
        try:
            lines.append(f"바디 수: {part.Bodies.Count}")
        except Exception:
            pass
        try:
            lines.append(f"메인 바디: {_safe(part.MainBody, 'Name')}")
        except Exception:
            pass
        try:
            lines.append(f"파라미터 수: {part.Parameters.Count}")
        except Exception:
            pass
        return "\n".join(lines)


# ═══════════ 🟡 세션·파트·스케치 (로컬 생성 — 비파괴, confirm 불필요) ═══════════
# CATIA의 Undo로 되돌릴 수 있고 디스크에 쓰지 않으므로 confirm 없이 실행한다.
# (전체 종료 quit_catia만 되돌리기 어려워 🔴 구역에 있다.)


@mcp.tool()
@catia_tool
def launch_catia() -> str:
    """CATIA V5를 실행(또는 이미 떠 있으면 연결)하고 화면에 표시합니다. (🟡)

    다른 도구들은 이미 떠 있는 CATIA에만 붙지만(GetActiveObject), 이 도구는
    Dispatch로 CATIA가 없으면 새로 띄운다. 라이선스 획득·기동에 시간이 걸릴 수 있다.
    """
    app = win32com.client.Dispatch("CATIA.Application")
    try:
        app.Visible = True
    except pythoncom.com_error:
        pass  # 일부 환경은 Visible 설정을 막는다 — 실행 자체는 됨
    try:
        count = app.Documents.Count
    except pythoncom.com_error:
        count = 0
    return f"CATIA 준비됨 (열린 문서 {count}개). new_part로 새 파트를 만들거나 open_document로 파일을 여세요."


@mcp.tool()
@catia_tool
def new_part() -> str:
    """새 Part(.CATPart) 문서를 만듭니다(파트 디자인 시작). (🟡)

    만들어진 파트는 활성 문서가 되어, 이어서 new_sketch·sketch_* 도구로 바로 설계할 수 있다.
    저장 전까지는 메모리에만 있으므로(디스크에 안 씀) save_document로 저장해야 남는다.
    """
    catia = _catia()
    doc = catia.Documents.Add("Part")
    return f"새 Part 생성됨: {_safe(doc, 'Name')}. (활성 문서로 설정됨 — new_sketch로 스케치를 시작하세요.)"


@mcp.tool()
@catia_tool
def open_document(path: str) -> str:
    """저장된 문서(.CATPart/.CATProduct/.CATDrawing)를 열어 세션에 둡니다. (🟡)

    조회만 하는 read 도구들과 달리, 이 도구는 문서를 열어 **닫지 않고 활성 상태로 남긴다**
    (이어서 편집·스케치할 수 있게). 이미 열려 있으면 그 문서를 알린다.

    Args:
        path: 열 파일의 절대경로.
    """
    catia = _catia()
    p = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(p):
        raise CatiaError(f"'{p}' 경로에 파일이 없습니다. 경로를 확인하세요.")
    existing = _find_open_doc(catia, p)
    if existing is not None:
        return f"이미 열려 있습니다: {_doc_label(existing)}"
    doc = catia.Documents.Open(p)
    return f"열림: {_doc_label(doc)}"


@mcp.tool()
@catia_tool
def new_sketch(plane: str = "xy", name: str = "") -> str:
    """활성 파트에 스케치를 새로 엽니다. (🟡)

    기준 평면(xy/yz/zx) 위에 스케치를 만들거나, plane='selection'이면 지금 CATIA
    화면에서 **선택돼 있는 평면/평평한 면** 위에 만든다(임의의 면에 스케치하려면
    CATIA에서 그 면을 먼저 클릭한 뒤 이 도구를 plane='selection'으로 부른다).
    그 외 문자열은 파트 안 평면 요소의 이름으로 해석한다(예: plane_offset으로 만든
    'Plane.1' — 로프트 단면 스케치를 띄운 평면 위에 열 때 쓴다).

    만든 스케치는 이후 sketch_point/line/circle/spline 도구의 기본 대상이 된다.
    여러 스케치를 만들 계획이면 name으로 의미 있는 이름(예: 'pips1', 'topFace')을
    붙여라 — 이후 pad/pocket·그리기 도구의 sketch= 인자로 그 이름을 그대로 넘기면
    엉뚱한 스케치를 물지 않는다.

    평면별 좌표·돌출 방향 (pad/pocket는 스케치 평면의 법선 방향으로 돌출·파냄):
        xy: 스케치 (x,y) → 3D (X,Y),  법선 +Z
        yz: 스케치 (x,y) → 3D (Y,Z),  법선 +X
        zx: 스케치 (x,y) → 3D (Z,X),  법선 +Y
    '세워진' 형상은 보통 xy 평면에 그려 +Z로 pad한다. 튀어나오는 쪽이 원하는
    방향과 반대면 pad/pocket을 reverse=True로 부른다.

    Args:
        plane: 'xy' | 'yz' | 'zx' | 'selection'(현재 선택된 면/평면) | 평면 요소 이름.
        name: 스케치에 붙일 이름(비우면 CATIA 자동, 예: 'Sketch.1'). 여러 스케치를
              다룰 때 의미 있는 이름을 주면 sketch= 로 지목하기 쉽다.

    Returns:
        생성된 스케치 이름(그리기 도구·pad·pocket에 sketch 인자로 그대로 넘길 것).
    """
    catia = _catia()
    doc, part = _active_part(catia)
    key = plane.lower().strip()
    if key in _PLANES:
        target = getattr(part.OriginElements, _PLANES[key])
    elif key in ("selection", "sel", "선택"):
        sel = doc.Selection
        if sel.Count == 0:
            raise CatiaError(
                "CATIA 화면에서 스케치할 평면이나 평평한 면을 먼저 선택한 뒤 "
                "plane='selection'으로 호출하세요."
            )
        target = sel.Item(1).Value
    else:
        try:
            target = _find_part_element(part, plane)
        except CatiaError:
            raise CatiaError(
                "plane은 xy/yz/zx, selection, 또는 파트 안 평면 요소 이름(예: plane_offset"
                "으로 만든 'Plane.1')이어야 합니다. 3D 요소 이름은 list_3d_geometry로 "
                "확인하세요."
            )
    ref = part.CreateReferenceFromObject(target)
    sketch = part.MainBody.Sketches.Add(ref)
    part.Update()
    nm = _apply_name(sketch, name)
    hint = ""
    if key in _PLANE_AXES:
        mapping, normal = _PLANE_AXES[key]
        hint = f" [좌표 {mapping}, 돌출 법선 {normal}]"
    return (
        f"스케치 생성됨: {nm} (평면: {key}){hint}. "
        f"이 이름('{nm}')을 pad/pocket·그리기 도구의 sketch= 인자로 그대로 넘기세요"
        "(비우면 '가장 최근 스케치'가 잡혀 여러 스케치가 있을 때 위험). "
        "sketch_point/line/circle/spline으로 그립니다."
    )


def _draw(sketch_name: str, action):
    """스케치 편집을 열고(OpenEdition) action(factory)을 실행한 뒤 닫고 재생성한다.

    OpenEdition은 2D 요소를 만드는 Factory2D를 돌려준다. 각 그리기 도구는 이걸 통해
    한 요소를 추가하고 CloseEdition으로 마무리한다. (요소별로 열고 닫으므로 에이전트가
    한 번에 하나씩 그려도 안전하다.)
    """
    catia = _catia()
    _, part = _active_part(catia)
    sketch = _get_sketch(part, sketch_name)
    factory = sketch.OpenEdition()
    try:
        result = action(factory)
    finally:
        sketch.CloseEdition()
    part.Update()
    return sketch, result


@mcp.tool()
@catia_tool
def sketch_point(x: float, y: float, sketch: str = "", name: str = "") -> str:
    """스케치에 점을 찍습니다. (🟡)

    좌표는 스케치 평면의 2D 좌표(단위 mm)입니다.

    Args:
        x, y: 점의 2D 좌표(mm).
        sketch: 대상 스케치 이름. 비우면 가장 최근에 만든 스케치.
        name: 이 점에 붙일 이름(비우면 CATIA 자동). 이후 구속에서 이 이름으로 참조한다.

    Returns:
        만들어진 점의 최종 이름(구속·조회에 쓸 수 있음).
    """
    sk, elem = _draw(sketch, lambda f: f.CreatePoint(float(x), float(y)))
    nm = _apply_name(elem, name)
    return f"점 추가: {nm} ({x}, {y}) → 스케치 {_safe(sk, 'Name')}"


@mcp.tool()
@catia_tool
def sketch_line(x1: float, y1: float, x2: float, y2: float, sketch: str = "", name: str = "") -> str:
    """스케치에 선분을 그립니다. (🟡)

    Args:
        x1, y1: 시작점(mm). x2, y2: 끝점(mm).
        sketch: 대상 스케치 이름. 비우면 가장 최근 스케치.
        name: 이 선에 붙일 이름(비우면 CATIA 자동). 구속에서 이 이름(끝점은 '이름.start'/'이름.end')으로 참조한다.

    Returns:
        만들어진 선의 최종 이름.
    """
    sk, elem = _draw(sketch, lambda f: f.CreateLine(float(x1), float(y1), float(x2), float(y2)))
    nm = _apply_name(elem, name)
    return f"선 추가: {nm} ({x1}, {y1})→({x2}, {y2}) → 스케치 {_safe(sk, 'Name')}"


@mcp.tool()
@catia_tool
def sketch_circle(cx: float, cy: float, radius: float, sketch: str = "", name: str = "") -> str:
    """스케치에 원을 그립니다(닫힌 원). (🟡)

    Args:
        cx, cy: 중심(mm). radius: 반지름(mm, 0보다 커야 함).
        sketch: 대상 스케치 이름. 비우면 가장 최근 스케치.
        name: 이 원에 붙일 이름(비우면 CATIA 자동). 중심은 '이름.center'로 참조한다.

    Returns:
        만들어진 원의 최종 이름.
    """
    if float(radius) <= 0:
        raise CatiaError("반지름(radius)은 0보다 커야 합니다.")
    sk, elem = _draw(sketch, lambda f: f.CreateClosedCircle(float(cx), float(cy), float(radius)))
    nm = _apply_name(elem, name)
    return f"원 추가: {nm} 중심({cx}, {cy}) 반지름 {radius} → 스케치 {_safe(sk, 'Name')}"


def _parse_circles(spec: str):
    """'[[15,15,2.5],[7.5,7.5,2.5,"p1"]]'(JSON) 또는 '15,15,2.5; 7.5,7.5,2.5'(구분)를
    [(cx, cy, r, name), ...]로 파싱한다. 네 번째 이름은 선택(없으면 '')."""
    s = (spec or "").strip()
    out = []
    if s.startswith("["):
        for item in json.loads(s):
            if len(item) < 3:
                raise CatiaError(f"원 형식 오류: {item} (cx, cy, r 필요).")
            nm = str(item[3]).strip() if len(item) >= 4 else ""
            out.append((float(item[0]), float(item[1]), float(item[2]), nm))
    else:
        for chunk in re.split(r"[;\n]+", s):
            chunk = chunk.strip()
            if not chunk:
                continue
            parts = re.split(r"[,\s]+", chunk)
            if len(parts) < 3:
                raise CatiaError(f"원 형식 오류: '{chunk}' (cx, cy, r 여야 합니다).")
            out.append((float(parts[0]), float(parts[1]), float(parts[2]), ""))
    if not out:
        raise CatiaError("원을 최소 1개는 지정하세요.")
    return out


@mcp.tool()
@catia_tool
def sketch_circles(circles: str, sketch: str = "", name_prefix: str = "") -> str:
    """여러 원을 한 번에 그립니다(스케치 편집·재생성 1회). (🟡)

    원 여러 개(예: 주사위 눈)를 그릴 때 sketch_circle을 반복 호출하면 원마다 스케치를
    열고 닫고 재생성해 느리고 Undo 단계가 많이 쌓인다. 이 도구는 한 번의 편집 세션에서
    전부 그려 빠르고 Undo도 한 번에 되돌린다.

    Args:
        circles: 원 목록. JSON('[[15,15,2.5],[7.5,7.5,2.5,"p1"]]') 또는
                 '15,15,2.5; 7.5,7.5,2.5' 형식. 각 항목은 cx, cy, r(mm)이고 네 번째로
                 이름(선택). 모든 반지름은 0보다 커야 한다.
        sketch: 대상 스케치 이름. 비우면 가장 최근 스케치.
        name_prefix: 이름을 안 준 원에 붙일 접두어(예: 'pip' → pip1, pip2 …).

    Returns:
        그린 원 개수와 이름 목록.
    """
    specs = _parse_circles(circles)
    if any(r <= 0 for (_cx, _cy, r, _n) in specs):
        raise CatiaError("모든 원의 반지름(r)은 0보다 커야 합니다.")
    catia = _catia()
    _, part = _active_part(catia)
    sk = _get_sketch(part, sketch)
    factory = sk.OpenEdition()
    try:
        elems = [factory.CreateClosedCircle(cx, cy, r) for (cx, cy, r, _n) in specs]
    finally:
        sk.CloseEdition()
    part.Update()
    names = []
    for i, (elem, (_cx, _cy, _r, nm)) in enumerate(zip(elems, specs), start=1):
        want = nm or (f"{name_prefix}{i}" if name_prefix else "")
        names.append(_apply_name(elem, want))
    return f"원 {len(names)}개 추가 → 스케치 {_safe(sk, 'Name')}\n  {', '.join(names)}"


@mcp.tool()
@catia_tool
def sketch_spline(points: str, sketch: str = "", name: str = "") -> str:
    """스케치에 스플라인 곡선을 그립니다(제어점들을 지나는 곡선). (🟡)

    Args:
        points: 제어점 목록. JSON('[[0,0],[10,5],[20,0]]') 또는 '0,0; 10,5; 20,0' 형식.
                최소 2개(mm 단위).
        sketch: 대상 스케치 이름. 비우면 가장 최근 스케치.
        name: 이 스플라인에 붙일 이름(비우면 CATIA 자동).

    Returns:
        만들어진 스플라인의 최종 이름.
    """
    pts = _parse_points(points)

    def build(f):
        # 각 좌표에 Point2D를 만들고, 그 배열로 스플라인을 만든다.
        # ⚠ 실기 검증 대상: Factory2D.CreateSpline이 Point2D 배열을 직접 받는지,
        # pywin32가 파이썬 리스트를 SAFEARRAY로 마샬링하는지 개발 PC에 CATIA가 없어
        # 미검증이다. (3D의 spline_3d는 AddNewSpline()+AddPoint(ref)로 요소별 추가하지만
        # 2D 팩토리엔 대응 API가 없어 배열 전달 방식을 쓴다.) 실기에서 실패하면 이 지점을
        # 점 생성 후 개별 추가하는 방식으로 바꿀 것.
        p2ds = [f.CreatePoint(px, py) for (px, py) in pts]
        return f.CreateSpline(p2ds)

    sk, elem = _draw(sketch, build)
    nm = _apply_name(elem, name)
    return f"스플라인 추가: {nm} 점 {len(pts)}개 → 스케치 {_safe(sk, 'Name')}"


@mcp.tool()
@catia_tool
def sketch_arc(
    cx: float, cy: float, radius: float, start_angle: float, end_angle: float,
    sketch: str = "", name: str = "",
) -> str:
    """스케치에 원호(호)를 그립니다. (🟡)

    각도는 도(°) 단위이고 x축 기준 반시계 방향으로 잰다. start_angle에서 end_angle까지
    반시계로 호를 그린다(예: 0→90 = 오른쪽에서 위쪽으로 가는 1/4 호).

    Args:
        cx, cy: 중심(mm). radius: 반지름(mm, 0보다 커야 함).
        start_angle, end_angle: 시작/끝 각도(도).
        sketch: 대상 스케치 이름. 비우면 가장 최근 스케치.
        name: 이 호에 붙일 이름(비우면 CATIA 자동). 중심은 '이름.center'로 참조한다.

    Returns:
        만들어진 호의 최종 이름.
    """
    if float(radius) <= 0:
        raise CatiaError("반지름(radius)은 0보다 커야 합니다.")
    sk, elem = _draw(
        sketch,
        lambda f: f.CreateCircle(
            float(cx), float(cy), float(radius),
            math.radians(float(start_angle)), math.radians(float(end_angle)),
        ),
    )
    nm = _apply_name(elem, name)
    return f"호 추가: {nm} 중심({cx}, {cy}) r={radius} {start_angle}°→{end_angle}° → 스케치 {_safe(sk, 'Name')}"


@mcp.tool()
@catia_tool
def sketch_rectangle(x1: float, y1: float, x2: float, y2: float, sketch: str = "", name: str = "") -> str:
    """스케치에 직사각형(선 4개)을 그립니다. (🟡)

    대각선 반대 방향의 두 꼭짓점으로 지정한다. 네 선의 끝점이 좌표로 맞닿아 있어
    닫힌 프로파일로 pad/pocket에 바로 쓸 수 있다(구속은 걸지 않는다 — 필요하면
    sketch_coincidence 등으로 추가).

    Args:
        x1, y1: 한 꼭짓점(mm). x2, y2: 반대 꼭짓점(mm).
        sketch: 대상 스케치 이름. 비우면 가장 최근 스케치.
        name: 이름 접두어. 주면 네 선이 '이름.bottom/right/top/left'로 명명된다.

    Returns:
        만들어진 네 선의 이름.
    """
    if float(x1) == float(x2) or float(y1) == float(y2):
        raise CatiaError("직사각형은 가로/세로 길이가 0이면 안 됩니다. 두 꼭짓점을 대각으로 주세요.")
    catia = _catia()
    _, part = _active_part(catia)
    sk = _get_sketch(part, sketch)
    factory = sk.OpenEdition()
    try:
        sides = [
            ("bottom", factory.CreateLine(float(x1), float(y1), float(x2), float(y1))),
            ("right", factory.CreateLine(float(x2), float(y1), float(x2), float(y2))),
            ("top", factory.CreateLine(float(x2), float(y2), float(x1), float(y2))),
            ("left", factory.CreateLine(float(x1), float(y2), float(x1), float(y1))),
        ]
    finally:
        sk.CloseEdition()
    part.Update()
    names = [_apply_name(line, f"{name}.{side}" if name else "") for side, line in sides]
    return (
        f"직사각형 추가: ({x1}, {y1})–({x2}, {y2}) → 스케치 {_safe(sk, 'Name')}\n"
        f"  선 4개: {', '.join(names)}"
    )


@mcp.tool()
@catia_tool
def list_sketch_geometry(sketch: str = "") -> str:
    """스케치의 기하요소(점·선·원 등)를 이름·종류·좌표로 나열합니다. (🟢 읽기)

    구속을 걸거나 이어서 그리기 전에 '지금 어떤 요소가 어떤 이름인지'를 확인하는 용도.
    기억에 의존하지 말고 이 도구로 현재 상태를 조회하세요(이름은 CATIA가 관리하므로
    저장·재열기 후에도 유지됩니다).

    Args:
        sketch: 대상 스케치 이름. 비우면 가장 최근 스케치.
    """
    catia = _catia()
    _, part = _active_part(catia)
    sk = _get_sketch(part, sketch)
    try:
        elems = sk.GeometricElements
        count = elems.Count
    except pythoncom.com_error as e:
        return f"기하요소를 읽지 못했습니다: {_com_message(e)}"
    if count == 0:
        return f"스케치 '{_safe(sk, 'Name')}'에 기하요소가 없습니다."
    lines = [f"스케치 '{_safe(sk, 'Name')}' — 요소 {count}개:"]
    for i in range(1, count + 1):
        lines.append("  - " + _describe_geo(elems.Item(i)))
    return _truncate("\n".join(lines))


# ─────────── 스케치 구속 (비파괴 — Undo 가능, confirm 불필요) ───────────
# 두 요소(bi): coincidence(일치)·perpendicular(직교)·parallel(평행)·tangent·concentric.
# 단일 요소(mono): horizontal(수평)·vertical(수직). 요소는 이름으로 지목하고, 선의
# 끝점은 '이름.start'/'이름.end', 원의 중심은 '이름.center'로 참조한다.


@mcp.tool()
@catia_tool
def sketch_coincidence(a: str, b: str, sketch: str = "") -> str:
    """두 요소(주로 두 점)를 일치(coincidence)시킵니다. (🟡)

    선의 끝점끼리 이으려면 '선이름.end', '선이름.start'로 지목합니다.

    Args:
        a, b: 일치시킬 두 요소/점. 예: 'A.end', 'B.start'.
        sketch: 대상 스케치. 비우면 가장 최근 스케치.

    Returns:
        생성된 구속의 CATIA 이름(예: 'Coincidence.1') — 실제 걸린 종류 확인용.
    """
    nm = _add_constraint("coincidence", a, b, sketch)
    return f"일치 구속 생성: {nm}  ({a} ↔ {b})"


@mcp.tool()
@catia_tool
def sketch_perpendicular(a: str, b: str, sketch: str = "") -> str:
    """두 선을 직교(perpendicular)시킵니다. (🟡)

    Args:
        a, b: 직교시킬 두 선 이름. 예: 'A', 'B'.
        sketch: 대상 스케치. 비우면 가장 최근 스케치.

    Returns:
        생성된 구속의 CATIA 이름(예: 'Perpendicularity.1').
    """
    nm = _add_constraint("perpendicular", a, b, sketch)
    return f"직교 구속 생성: {nm}  ({a} ⊥ {b})"


@mcp.tool()
@catia_tool
def sketch_parallel(a: str, b: str, sketch: str = "") -> str:
    """두 선을 평행(parallel)하게 합니다. (🟡)

    Args:
        a, b: 평행시킬 두 선 이름.
        sketch: 대상 스케치. 비우면 가장 최근 스케치.
    """
    nm = _add_constraint("parallel", a, b, sketch)
    return f"평행 구속 생성: {nm}  ({a} ∥ {b})"


@mcp.tool()
@catia_tool
def sketch_horizontal(line: str, sketch: str = "") -> str:
    """선을 수평(horizontal)으로 구속합니다. (🟡)

    Args:
        line: 대상 선 이름.
        sketch: 대상 스케치. 비우면 가장 최근 스케치.
    """
    nm = _add_constraint("horizontal", line, "", sketch)
    return f"수평 구속 생성: {nm}  ({line})"


@mcp.tool()
@catia_tool
def sketch_vertical(line: str, sketch: str = "") -> str:
    """선을 수직(vertical)으로 구속합니다. (🟡)

    Args:
        line: 대상 선 이름.
        sketch: 대상 스케치. 비우면 가장 최근 스케치.
    """
    nm = _add_constraint("vertical", line, "", sketch)
    return f"수직 구속 생성: {nm}  ({line})"


@mcp.tool()
@catia_tool
def sketch_dimension(kind: str, a: str, b: str = "", value: float | None = None, sketch: str = "") -> str:
    """스케치에 치수 구속을 걸고(선택) 값을 지정합니다 — 파라메트릭 설계의 핵심. (🟡)

    종류(kind):
        length   — 선 하나의 길이 (a만, mm)
        radius   — 원/호 하나의 반지름 (a만, mm)
        distance — 두 요소 사이 거리 (a와 b, mm). 점끼리는 '이름.start'/'이름.end'로 지목.
        angle    — 두 선 사이 각도 (a와 b, 도)

    value를 주면 그 값으로 설정하고 형상을 재계산한다. 생략하면 현재 값으로 구속만 건다.
    이후 list_parameters/set_parameter로도 이 치수를 찾아 바꿀 수 있다.

    Args:
        kind: length | radius | distance | angle.
        a: 첫 요소 이름. b: 둘째 요소 이름(distance/angle만 필수).
        value: 설정할 치수 값(mm 또는 도). 생략하면 현재 값 유지.
        sketch: 대상 스케치. 비우면 가장 최근 스케치.

    Returns:
        생성된 구속의 CATIA 이름과 값(예: 'Length.1 = 50mm') — 실제 걸린 종류 확인용.
    """
    k = (kind or "").lower().strip()
    if k not in _DIM_TYPES:
        raise CatiaError(f"kind는 {'/'.join(_DIM_TYPES)} 중 하나여야 합니다 (받은 값: '{kind}').")
    if k in ("distance", "angle") and not b.strip():
        raise CatiaError(f"{k} 치수는 두 요소가 필요합니다 — b 인자에 둘째 요소 이름을 주세요.")
    if k in ("length", "radius") and b.strip():
        raise CatiaError(f"{k} 치수는 단일 요소 구속입니다 — b 인자를 비우세요.")
    catia = _catia()
    _, part = _active_part(catia)
    sk = _get_sketch(part, sketch)
    sk.OpenEdition()
    try:
        ref1 = _resolve_ref(part, sk, a)
        if b:
            cst = part.Constraints.AddBiEltCst(_DIM_TYPES[k], ref1, _resolve_ref(part, sk, b))
        else:
            cst = part.Constraints.AddMonoEltCst(_DIM_TYPES[k], ref1)
        if value is not None:
            cst.Dimension.Value = float(value)  # 길이는 mm, 각도는 도(°)
    finally:
        sk.CloseEdition()
    part.Update()
    nm = _safe(cst, "Name") or "(이름 없음)"
    try:
        cur = cst.Dimension.Value
    except Exception:  # noqa: BLE001 — 값 표시는 부가정보
        cur = value
    unit = "°" if k == "angle" else "mm"
    target = f"{a} ↔ {b}" if b else a
    return f"치수 구속 생성: {nm} = {cur}{unit}  ({k}: {target})"


# ═══════ 🟡 3D 와이어프레임·서피스·솔리드 피처 (로컬 생성 — 비파괴, confirm 불필요) ═══════
# 와이어프레임(점·선·스플라인·평면)과 로프트 서피스는 HybridShapeFactory로 만들어
# 기하학적 세트(HybridBody)에 담고, 솔리드 피처(pad/pocket/shaft/groove/fillet)는
# ShapeFactory로 MainBody에 쌓는다. 모두 CATIA Undo로 되돌릴 수 있고 디스크에 쓰지
# 않으므로 스케치 도구와 같은 🟡 등급으로 confirm 없이 실행한다. 좌표 단위는 mm.


def _hsf(part):
    """HybridShapeFactory(3D 와이어프레임·서피스 팩토리)를 돌려준다."""
    try:
        return part.HybridShapeFactory
    except pythoncom.com_error as e:
        raise CatiaError(
            f"3D 와이어프레임 팩토리를 열지 못했습니다: {_com_message(e)} "
            "(GSD 계열 라이선스가 없는 환경일 수 있습니다.)"
        )


def _shape_factory(part):
    """ShapeFactory(솔리드 피처 팩토리)를 돌려준다."""
    try:
        return part.ShapeFactory
    except pythoncom.com_error as e:
        raise CatiaError(f"솔리드 피처 팩토리를 열지 못했습니다: {_com_message(e)}")


def _geoset(part, name: str = ""):
    """3D 요소를 담을 기하학적 세트(HybridBody)를 확보한다.

    name 지정 → 그 이름의 세트(없으면 만들어 이름 지정). name 비움 → 첫 세트,
    하나도 없으면 새로 만든다. (3D 와이어프레임·서피스는 바디가 아니라 세트에
    담는 게 V5 관례다.)
    """
    hbs = part.HybridBodies
    if name:
        try:
            return hbs.Item(name)
        except pythoncom.com_error:
            hb = hbs.Add()
            try:
                hb.Name = name
            except Exception:  # noqa: BLE001 — 이름 지정 실패 시 자동 이름 유지
                pass
            return hb
    if hbs.Count > 0:
        return hbs.Item(1)
    return hbs.Add()


def _iter_hybrid_shapes(part):
    """모든 기하학적 세트(중첩 포함)의 (세트, 3D 요소)를 순회한다."""

    def walk(hbs):
        for i in range(1, hbs.Count + 1):
            hb = hbs.Item(i)
            try:
                shapes = hb.HybridShapes
                for j in range(1, shapes.Count + 1):
                    yield hb, shapes.Item(j)
            except Exception:  # noqa: BLE001 — 이 세트만 건너뛴다
                pass
            try:
                yield from walk(hb.HybridBodies)
            except Exception:  # noqa: BLE001
                pass

    try:
        yield from walk(part.HybridBodies)
    except Exception:  # noqa: BLE001
        return


def _find_part_element(part, token: str):
    """이름으로 파트 안의 요소를 찾는다: 기준 평면(xy/yz/zx) → 스케치 → 3D 요소 순.

    로프트 단면·가이드, 스케치 기준 평면 등 '파트 어딘가의 요소'를 이름 하나로 지목할
    때 쓴다. (스케치 안 2D 요소는 _resolve_ref 담당 — 여긴 파트 수준.)
    """
    t = (token or "").strip()
    if t.lower() in _PLANES:
        return getattr(part.OriginElements, _PLANES[t.lower()])
    try:
        bodies = part.Bodies
        for i in range(1, bodies.Count + 1):
            try:
                return bodies.Item(i).Sketches.Item(t)
            except Exception:  # noqa: BLE001 — 이 바디에 없으면 다음 바디
                continue
    except Exception:  # noqa: BLE001
        pass
    for _hb, shape in _iter_hybrid_shapes(part):
        if _safe(shape, "Name") == t:
            return shape
    raise CatiaError(
        f"파트에서 '{t}' 요소를 찾지 못했습니다. 스케치는 그 이름 그대로, 3D 요소는 "
        "list_3d_geometry로 확인한 이름으로 지목하세요."
    )


def _part_ref(part, token: str):
    """파트 수준 요소 이름을 Reference로 푼다(로프트 단면·가이드·기준 평면용)."""
    return part.CreateReferenceFromObject(_find_part_element(part, token))


def _parse_points3d(points: str):
    """'[[0,0,0],[10,5,3]]'(JSON) 또는 '0,0,0; 10,5,3'(세미콜론 구분)을 [(x,y,z),...]로 파싱."""
    s = (points or "").strip()
    triples = []
    if s.startswith("["):
        for p in json.loads(s):
            if len(p) < 3:
                raise CatiaError(f"좌표 형식 오류: {p} (x,y,z 세 값이어야 합니다).")
            triples.append((float(p[0]), float(p[1]), float(p[2])))
    else:
        for chunk in re.split(r"[;\n]+", s):
            chunk = chunk.strip()
            if not chunk:
                continue
            xyz = re.split(r"[,\s]+", chunk)
            if len(xyz) < 3:
                raise CatiaError(f"좌표 형식 오류: '{chunk}' (x,y,z 여야 합니다).")
            triples.append((float(xyz[0]), float(xyz[1]), float(xyz[2])))
    return triples


def _parse_names(names: str) -> list[str]:
    """'["Sketch.1","Sketch.2"]'(JSON) 또는 'Sketch.1, Sketch.2'(구분자)를 이름 목록으로."""
    s = (names or "").strip()
    if not s:
        return []
    if s.startswith("["):
        return [str(x).strip() for x in json.loads(s) if str(x).strip()]
    return [t.strip() for t in re.split(r"[;,\n]+", s) if t.strip()]


def _describe_shape3d(shape) -> str:
    """3D 요소 하나를 'Name (좌표)'로 요약한다(점 좌표는 best-effort)."""
    name = _safe(shape, "Name") or "(이름 없음)"
    try:  # 점: X/Y/Z 길이 파라미터가 있음
        return f"{name} ({round(shape.X.Value, 3)}, {round(shape.Y.Value, 3)}, {round(shape.Z.Value, 3)})"
    except Exception:  # noqa: BLE001 — 점이 아니면 이름만
        return name


@mcp.tool()
@catia_tool
def point_3d(x: float, y: float, z: float, name: str = "", geoset: str = "") -> str:
    """3D 공간 좌표에 점을 만듭니다(기하학적 세트에 저장). (🟡)

    Args:
        x, y, z: 3D 좌표(mm).
        name: 붙일 이름(비우면 CATIA 자동). 이후 다른 도구에서 이 이름으로 지목한다.
        geoset: 담을 기하학적 세트 이름. 비우면 첫 세트(없으면 새로 만듦).

    Returns:
        만들어진 점의 최종 이름.
    """
    catia = _catia()
    _, part = _active_part(catia)
    hb = _geoset(part, geoset)
    pt = _hsf(part).AddNewPointCoord(float(x), float(y), float(z))
    hb.AppendHybridShape(pt)
    nm = _apply_name(pt, name)
    part.Update()
    return f"3D 점 추가: {nm} ({x}, {y}, {z}) → 세트 {_safe(hb, 'Name')}"


@mcp.tool()
@catia_tool
def line_3d(
    x1: float, y1: float, z1: float, x2: float, y2: float, z2: float,
    name: str = "", geoset: str = "",
) -> str:
    """3D 공간에서 두 좌표를 잇는 직선을 만듭니다. (🟡)

    끝점 두 개를 datum 점으로 만들어 잇는다(점들도 같은 세트에 함께 남는다).

    Args:
        x1, y1, z1: 시작점(mm). x2, y2, z2: 끝점(mm).
        name: 직선에 붙일 이름(비우면 CATIA 자동).
        geoset: 담을 기하학적 세트 이름. 비우면 첫 세트(없으면 새로 만듦).

    Returns:
        만들어진 직선의 최종 이름.
    """
    catia = _catia()
    _, part = _active_part(catia)
    hsf = _hsf(part)
    hb = _geoset(part, geoset)
    p1 = hsf.AddNewPointCoord(float(x1), float(y1), float(z1))
    p2 = hsf.AddNewPointCoord(float(x2), float(y2), float(z2))
    hb.AppendHybridShape(p1)
    hb.AppendHybridShape(p2)
    line = hsf.AddNewLinePtPt(
        part.CreateReferenceFromObject(p1), part.CreateReferenceFromObject(p2)
    )
    hb.AppendHybridShape(line)
    nm = _apply_name(line, name)
    part.Update()
    return f"3D 직선 추가: {nm} ({x1}, {y1}, {z1})→({x2}, {y2}, {z2}) → 세트 {_safe(hb, 'Name')}"


@mcp.tool()
@catia_tool
def spline_3d(points: str, name: str = "", geoset: str = "") -> str:
    """3D 제어점들을 지나는 스플라인 곡선을 만듭니다. (🟡)

    각 좌표에 datum 점을 만들고 그 점들을 지나는 곡선을 만든다(점들도 세트에 남는다).
    로프트의 가이드 곡선으로 쓸 수 있다.

    Args:
        points: 제어점 목록. JSON('[[0,0,0],[10,5,3],[20,0,10]]') 또는
                '0,0,0; 10,5,3; 20,0,10' 형식. 최소 2개(mm 단위).
        name: 스플라인에 붙일 이름(비우면 CATIA 자동).
        geoset: 담을 기하학적 세트 이름. 비우면 첫 세트(없으면 새로 만듦).

    Returns:
        만들어진 스플라인의 최종 이름.
    """
    pts = _parse_points3d(points)
    if len(pts) < 2:
        raise CatiaError("스플라인은 점이 최소 2개 필요합니다.")
    catia = _catia()
    _, part = _active_part(catia)
    hsf = _hsf(part)
    hb = _geoset(part, geoset)
    spline = hsf.AddNewSpline()
    for (px, py, pz) in pts:
        pt = hsf.AddNewPointCoord(px, py, pz)
        hb.AppendHybridShape(pt)
        spline.AddPoint(part.CreateReferenceFromObject(pt))
    hb.AppendHybridShape(spline)
    nm = _apply_name(spline, name)
    part.Update()
    return f"3D 스플라인 추가: {nm} 점 {len(pts)}개 → 세트 {_safe(hb, 'Name')}"


@mcp.tool()
@catia_tool
def plane_offset(base: str, offset: float, reverse: bool = False, name: str = "", geoset: str = "") -> str:
    """기존 평면에서 지정 거리만큼 띄운 평면을 만듭니다. (🟡)

    로프트 단면 스케치들을 서로 다른 높이에 배치하는 준비 단계: 이 도구로 평면을
    만들고 new_sketch(plane='그 이름')로 그 위에 스케치를 연다.

    Args:
        base: 기준 평면 — 'xy'/'yz'/'zx' 또는 기존 평면 요소 이름.
        offset: 띄울 거리(mm).
        reverse: True면 반대 방향으로 띄운다.
        name: 평면에 붙일 이름(비우면 CATIA 자동, 예: 'Plane.1').
        geoset: 담을 기하학적 세트 이름. 비우면 첫 세트(없으면 새로 만듦).

    Returns:
        만들어진 평면의 최종 이름(new_sketch의 plane 인자로 넘길 수 있음).
    """
    catia = _catia()
    _, part = _active_part(catia)
    hb = _geoset(part, geoset)
    pl = _hsf(part).AddNewPlaneOffset(_part_ref(part, base), float(offset), bool(reverse))
    hb.AppendHybridShape(pl)
    nm = _apply_name(pl, name)
    part.Update()
    return f"평면 추가: {nm} ({base}에서 {offset}mm{' 반대방향' if reverse else ''}) → 세트 {_safe(hb, 'Name')}"


@mcp.tool()
@catia_tool
def create_loft(sections: str, guides: str = "", solid: bool = False, name: str = "", geoset: str = "") -> str:
    """단면(스케치/곡선)들을 순서대로 이어 로프트를 만듭니다. (🟡)

    solid=False(기본)면 서피스 로프트(기하학적 세트에 저장), True면 솔리드 로프트
    (MainBody에 재료로 추가 — 단면이 모두 닫힌 프로파일이어야 한다).

    전형적 흐름: plane_offset으로 평면 여러 장 → 각 평면에 new_sketch + sketch_circle
    등으로 단면 → create_loft(sections='Sketch.2, Sketch.3, Sketch.4').

    Args:
        sections: 단면 이름 목록(잇는 순서대로, 최소 2개). 'Sketch.1, Sketch.2' 또는 JSON.
        guides: 가이드 곡선 이름 목록(선택). spline_3d/line_3d 등으로 만든 3D 곡선.
        solid: True면 솔리드 로프트(ShapeFactory), False면 서피스 로프트.
        name: 로프트에 붙일 이름(비우면 CATIA 자동).
        geoset: (서피스일 때) 담을 기하학적 세트 이름.

    Returns:
        만들어진 로프트의 최종 이름.
    """
    secs = _parse_names(sections)
    if len(secs) < 2:
        raise CatiaError("로프트는 단면이 최소 2개 필요합니다.")
    gds = _parse_names(guides)
    catia = _catia()
    _, part = _active_part(catia)
    if solid:
        part.InWorkObject = part.MainBody
        loft = _shape_factory(part).AddNewLoft()
    else:
        loft = _hsf(part).AddNewLoft()
    # ⚠ 실기 검증 대상: 단면 방향(1)과 커플링 점 생략(빈 Reference)은 V5 매크로 관례를 따랐다.
    nothing = part.CreateReferenceFromName("")
    for s in secs:
        loft.AddSectionToLoft(_part_ref(part, s), 1, nothing)
    for g in gds:
        loft.AddGuideToLoft(_part_ref(part, g))
    if not solid:
        _geoset(part, geoset).AppendHybridShape(loft)
    nm = _apply_name(loft, name)
    part.Update()
    kind = "솔리드" if solid else "서피스"
    tail = f", 가이드 {len(gds)}개" if gds else ""
    return f"{kind} 로프트 생성: {nm} (단면 {len(secs)}개{tail})"


@mcp.tool()
@catia_tool
def list_3d_geometry() -> str:
    """기하학적 세트의 3D 요소와 바디의 솔리드 피처를 이름으로 나열합니다. (🟢 읽기)

    로프트 단면·가이드를 지목하거나 이어서 작업하기 전에 '지금 무엇이 어떤 이름인지'
    확인하는 용도. 기억에 의존하지 말고 이 도구로 현재 상태를 조회하세요.
    """
    catia = _catia()
    _, part = _active_part(catia)
    shape_lines = []
    for hb, shape in _iter_hybrid_shapes(part):
        shape_lines.append(f"  - [{_safe(hb, 'Name')}] {_describe_shape3d(shape)}")
    out = (
        [f"3D 요소 {len(shape_lines)}개:"] + shape_lines
        if shape_lines
        else ["3D 요소 없음 (point_3d/line_3d/spline_3d/plane_offset으로 만듭니다)."]
    )
    try:
        bodies = part.Bodies
        for i in range(1, bodies.Count + 1):
            body = bodies.Item(i)
            try:
                shapes = body.Shapes
                if shapes.Count == 0:
                    continue
                out.append(f"바디 '{_safe(body, 'Name')}' 피처 {shapes.Count}개:")
                for j in range(1, shapes.Count + 1):
                    out.append(f"  - {_safe(shapes.Item(j), 'Name')}")
            except Exception:  # noqa: BLE001 — 이 바디만 건너뛴다
                continue
    except Exception:  # noqa: BLE001
        pass
    return _truncate("\n".join(out))


@mcp.tool()
@catia_tool
def list_sketches() -> str:
    """활성 파트의 모든 스케치를 이름·소속 바디·요소 수로 나열합니다. (🟢 읽기)

    여러 스케치를 그린 뒤 pad/pocket에 어느 것을 넘길지 헷갈릴 때, 기억에 의존하지
    말고 이 도구로 지금 어떤 스케치가 있는지 확인하세요. (list_3d_geometry는 3D
    요소·솔리드 피처만 보여주고 스케치는 이 도구가 담당합니다.) 여기서 얻은 이름을
    pad/pocket·그리기 도구의 sketch= 인자로 그대로 넘기면 엉뚱한 스케치를 물지 않습니다.
    특정 스케치 안에 무엇이 있는지는 list_sketch_geometry로 봅니다.
    """
    catia = _catia()
    _, part = _active_part(catia)
    out = []
    try:
        bodies = part.Bodies
        for i in range(1, bodies.Count + 1):
            body = bodies.Item(i)
            try:
                sketches = body.Sketches
            except Exception:  # noqa: BLE001 — 스케치 컬렉션이 없는 바디면 건너뛴다
                continue
            for j in range(1, sketches.Count + 1):
                sk = sketches.Item(j)
                try:
                    ecount = str(sk.GeometricElements.Count)
                except Exception:  # noqa: BLE001 — 요소 수는 부가정보
                    ecount = "?"
                out.append(f"  - {_safe(sk, 'Name')}  (바디 '{_safe(body, 'Name')}', 요소 {ecount}개)")
    except Exception as e:  # noqa: BLE001 — 전체 실패도 안내 문자열로 돌려준다
        return f"스케치 목록을 읽지 못했습니다: {type(e).__name__}: {e}"
    if not out:
        return "스케치가 없습니다. new_sketch로 먼저 스케치를 여세요."
    return _truncate(f"스케치 {len(out)}개:\n" + "\n".join(out))


@mcp.tool()
@catia_tool
def get_selection() -> str:
    """지금 CATIA 화면에서 사용자가 선택해 둔 항목들을 나열합니다. (🟢 읽기)

    "선택한 것을 ~해줘" 같은 요청을 받으면 먼저 이 도구로 무엇이 선택돼 있는지
    확인한다. 이름이 나오는 항목은 delete_element/set_color/set_visibility 등
    이름 기반 도구로 바로 지목할 수 있고, 이름이 없는 모서리/면은
    fillet·chamfer_selected_edges/shell_selected_faces/new_sketch(plane='selection')
    처럼 선택을 직접 소비하는 도구로 이어서 쓴다.
    """
    catia = _catia()
    doc = _active_document(catia)
    sel = doc.Selection
    n = int(sel.Count)
    if n == 0:
        return (
            "선택된 항목이 없습니다. CATIA 화면에서 (Ctrl 클릭으로 여러 개) 선택한 뒤 "
            "다시 호출하세요."
        )
    lines = [f"선택된 항목 {n}개 (문서: {_safe(doc, 'Name')}):"]
    for i in range(1, n + 1):
        try:
            item = sel.Item2(i)
        except Exception:  # noqa: BLE001 — 이 항목만 건너뛴다
            lines.append(f"  [{i}] (항목을 읽지 못함)")
            continue
        kind = _safe(item, "Type") or "종류 미상"
        try:
            value_name = _safe(item.Value, "Name")
        except Exception:  # noqa: BLE001 — 모서리/면 등은 Name이 없을 수 있다
            value_name = ""
        try:
            leaf = _safe(item.LeafProduct, "Name")  # 어셈블리에서 소속 컴포넌트
        except Exception:  # noqa: BLE001 — Part 단독 문서면 없다
            leaf = ""
        desc = f"  [{i}] {kind}" + (f": {value_name}" if value_name else " (이름 없음 — 화면 선택 소비 도구로 사용)")
        if leaf and leaf != value_name:
            desc += f"  [컴포넌트: {leaf}]"
        lines.append(desc)
    return _truncate("\n".join(lines))


# ─────────── 면(BRep) 조회·선택 — LLM이 좌표/법선으로 면을 직접 고르게 한다 ───────────
# CATIA 면은 이름 없는 위상 참조라, 이름으로 못 잡고 Selection.Search로 훑는다.
# list_faces가 각 면의 면적·법선·중심을 인덱스와 함께 보여주면, LLM이 "법선 +Z인
# 면=윗면"처럼 판단해 select_face로 그 면을 잡는다. 잡은 면은 기존 선택 소비 도구
# (new_sketch plane='selection' / shell_selected_faces)가 그대로 쓴다.
# ⚠ 실기 검증 대상 전반: 검색 쿼리 문자열, GetPlane/GetCOG의 out-배열 형식.


def _out_array(n: int):
    """SPAWorkbench가 out-배열(byref)로 채우는 double 배열 VARIANT를 만든다."""
    return win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8 | pythoncom.VT_BYREF, [0.0] * n
    )


# ⚠ 실기 검증 대상: 면 전체를 잡는 Search 쿼리. 버전/로케일 편차가 있어 후보를
# 순서대로 시도하고, 하나라도 면을 찾으면 그걸 쓴다(우아한 저하).
_FACE_QUERIES = ("CATPrtSearch.Face,all", "Topology.CGMFace,all", "Type=Face,all")


def _search_faces(sel) -> int:
    """활성 파트의 면들을 현재 선택에 채우고 개수를 돌려준다(0이면 실패)."""
    for q in _FACE_QUERIES:
        try:
            sel.Clear()
            sel.Search(q)
            if sel.Count > 0:
                return int(sel.Count)
        except Exception:  # noqa: BLE001 — 이 쿼리가 안 되면 다음 후보
            continue
    return 0


def _measure_face(spa, ref) -> dict:
    """면 하나의 면적·무게중심·(평면이면)법선을 측정해 dict로 돌려준다(각각 best-effort)."""
    info: dict = {"area": None, "center": None, "normal": None}
    meas = spa.GetMeasurable(ref)
    try:
        info["area"] = round(float(meas.Area), 3)
    except Exception:  # noqa: BLE001 — 면적을 못 얻어도 나머지는 시도
        pass
    try:
        cog = _out_array(3)
        meas.GetCOG(cog)
        info["center"] = tuple(round(float(v), 2) for v in cog.value[:3])
    except Exception:  # noqa: BLE001
        pass
    try:  # 평면 면만: GetPlane = 원점(3)+평면내 방향2개(3+3). 법선 = 두 방향의 외적.
        pl = _out_array(9)
        meas.GetPlane(pl)
        v = [float(x) for x in pl.value]
        ax, ay, az = v[3], v[4], v[5]
        bx, by, bz = v[6], v[7], v[8]
        nx, ny, nz = ay * bz - az * by, az * bx - ax * bz, ax * by - ay * bx
        mag = (nx * nx + ny * ny + nz * nz) ** 0.5
        if mag:
            info["normal"] = (round(nx / mag, 3), round(ny / mag, 3), round(nz / mag, 3))
    except Exception:  # noqa: BLE001 — 곡면이면 평면이 없어 실패한다(정상)
        pass
    return info


@mcp.tool()
@catia_tool
def list_faces() -> str:
    """활성 파트 솔리드의 면들을 인덱스·면적·법선·중심으로 나열합니다. (🟢 읽기)

    LLM이 '윗면(법선 +Z)'처럼 원하는 면을 이름 없이도 직접 고를 수 있게, 면을 전부
    찾아 각 면의 면적·(평면이면)법선·무게중심을 보여준다. 여기 나온 [인덱스]를
    select_face에 넘겨 그 면을 잡은 뒤, new_sketch(plane='selection')로 스케치하거나
    shell_selected_faces로 뚫는다.

    ⚠ 인덱스는 '지금 이 순간'의 검색 순서다. 형상을 편집하면(pad/pocket/삭제 등)
    순서가 바뀔 수 있으니, list_faces로 확인한 직후 곧바로 select_face를 부르세요.
    """
    catia = _catia()
    doc, _part = _active_part(catia)
    sel = doc.Selection
    n = _search_faces(sel)
    if n == 0:
        sel.Clear()
        return (
            "면을 찾지 못했습니다. 솔리드(pad 등)가 있는지 확인하세요. "
            "(⚠ 면 검색 쿼리가 이 CATIA 버전과 맞지 않을 수도 있습니다 — _FACE_QUERIES 확인.)"
        )
    refs = [sel.Item2(i).Reference for i in range(1, n + 1)]
    spa = doc.GetWorkbench("SPAWorkbench")
    lines = [f"면 {n}개 (현재 검색 순서 — 편집하면 인덱스가 바뀝니다):"]
    for idx, ref in enumerate(refs):
        try:
            info = _measure_face(spa, ref)
            desc = f"  [{idx}] 면적 {info['area']}"
            desc += f", 법선 {info['normal']}" if info["normal"] else ", 법선 (곡면)"
            if info["center"]:
                desc += f", 중심 {info['center']}"
            lines.append(desc)
        except Exception as e:  # noqa: BLE001 — 이 면만 건너뛴다
            lines.append(f"  [{idx}] (측정 실패: {type(e).__name__})")
    sel.Clear()  # 조회 뒤 선택을 비워 사용자 화면을 방해하지 않는다
    return _truncate("\n".join(lines))


@mcp.tool()
@catia_tool
def select_face(index: int, add: bool = False) -> str:
    """list_faces의 [인덱스] 면을 CATIA 현재 선택으로 잡습니다. (🟡)

    잡은 뒤 new_sketch(plane='selection', name=...)로 그 면에 스케치하거나,
    shell_selected_faces로 뚫습니다. 여러 면을 함께 잡으려면 add=True로 이어 부릅니다.

    ⚠ list_faces를 부른 직후 바로 사용하세요(편집하면 인덱스가 바뀝니다).

    Args:
        index: list_faces에서 확인한 면 인덱스(0부터).
        add: True면 기존 선택에 더한다(여러 면 동시 선택). False면 이 면만.

    Returns:
        선택된 면의 법선·중심 요약과 현재 선택 개수.
    """
    catia = _catia()
    doc, _part = _active_part(catia)
    sel = doc.Selection
    # 면 검색은 선택을 덮어쓰므로, add=True면 이미 잡아 둔 참조를 먼저 보존한다.
    keep = [sel.Item2(i).Reference for i in range(1, sel.Count + 1)] if add else []
    n = _search_faces(sel)
    if n == 0:
        sel.Clear()
        return "면을 찾지 못했습니다. list_faces로 먼저 확인하세요."
    if not 0 <= index < n:
        sel.Clear()
        return f"인덱스 {index}가 범위를 벗어났습니다(면 {n}개, 0~{n - 1}). list_faces로 다시 확인하세요."
    target = sel.Item2(index + 1).Reference
    try:
        info = _measure_face(doc.GetWorkbench("SPAWorkbench"), target)
    except Exception:  # noqa: BLE001 — 요약 정보는 부가라 실패해도 선택은 진행
        info = {"normal": None, "center": None}
    sel.Clear()
    for r in keep:
        try:
            sel.Add(r)
        except Exception:  # noqa: BLE001 — 보존 실패한 참조는 건너뛴다
            pass
    sel.Add(target)
    tail = (f" 법선 {info['normal']}" if info.get("normal") else " (곡면)")
    if info.get("center"):
        tail += f", 중심 {info['center']}"
    return (
        f"면 [{index}] 선택됨{tail}. (현재 선택 {sel.Count}개)\n"
        "이제 new_sketch(plane='selection', name=...)로 이 면에 스케치하거나 "
        "shell_selected_faces로 뚫으세요."
    )


# ─────────── 모서리(BRep) 조회·선택 — 필렛/모따기를 클릭 없이 ───────────
# 면과 같은 원리(이름 없는 위상 참조). list_edges가 각 모서리의 길이·중점을 인덱스와
# 함께 보여주면, LLM이 '중점 (15,30,30)인 모서리=윗-뒤 모서리'처럼 판단해 select_edge로
# 잡는다. 잡은 모서리는 fillet_selected_edges/chamfer_selected_edges가 그대로 쓴다.
# 상자의 모서리는 길이가 다 같을 수 있어 중점(GetCOG)이 주요 식별자다.


# ⚠ 실기 검증 대상: 모서리 전체를 잡는 Search 쿼리(버전/로케일 편차 → 후보 순차 시도).
_EDGE_QUERIES = ("CATPrtSearch.Edge,all", "Topology.CGMEdge,all", "Type=Edge,all")


def _search_edges(sel) -> int:
    """활성 파트의 모서리들을 현재 선택에 채우고 개수를 돌려준다(0이면 실패)."""
    for q in _EDGE_QUERIES:
        try:
            sel.Clear()
            sel.Search(q)
            if sel.Count > 0:
                return int(sel.Count)
        except Exception:  # noqa: BLE001 — 이 쿼리가 안 되면 다음 후보
            continue
    return 0


def _measure_edge(spa, ref) -> dict:
    """모서리 하나의 길이·중점을 측정해 dict로 돌려준다(각각 best-effort)."""
    info: dict = {"length": None, "center": None}
    meas = spa.GetMeasurable(ref)
    try:
        info["length"] = round(float(meas.Length), 3)
    except Exception:  # noqa: BLE001
        pass
    try:
        cog = _out_array(3)
        meas.GetCOG(cog)
        info["center"] = tuple(round(float(v), 2) for v in cog.value[:3])
    except Exception:  # noqa: BLE001
        pass
    return info


@mcp.tool()
@catia_tool
def list_edges() -> str:
    """활성 파트 솔리드의 모서리들을 인덱스·길이·중점으로 나열합니다. (🟢 읽기)

    LLM이 특정 모서리를 이름 없이도 고를 수 있게, 모서리를 전부 찾아 각 모서리의
    길이와 중점(무게중심)을 보여준다. 여기 나온 [인덱스]를 select_edge에 넘겨 그
    모서리를 잡은 뒤, fillet_selected_edges(라운드)/chamfer_selected_edges(모따기)를
    부른다. (상자처럼 길이가 같은 모서리가 많으면 중점으로 구분한다.)

    ⚠ 인덱스는 '지금 이 순간'의 검색 순서다. 형상을 편집하면 순서가 바뀌므로,
    list_edges로 확인한 직후 곧바로 select_edge를 부르세요.
    """
    catia = _catia()
    doc, _part = _active_part(catia)
    sel = doc.Selection
    n = _search_edges(sel)
    if n == 0:
        sel.Clear()
        return (
            "모서리를 찾지 못했습니다. 솔리드(pad 등)가 있는지 확인하세요. "
            "(⚠ 모서리 검색 쿼리가 이 CATIA 버전과 맞지 않을 수도 있습니다 — _EDGE_QUERIES 확인.)"
        )
    refs = [sel.Item2(i).Reference for i in range(1, n + 1)]
    spa = doc.GetWorkbench("SPAWorkbench")
    lines = [f"모서리 {n}개 (현재 검색 순서 — 편집하면 인덱스가 바뀝니다):"]
    for idx, ref in enumerate(refs):
        try:
            info = _measure_edge(spa, ref)
            desc = f"  [{idx}] 길이 {info['length']}"
            if info["center"]:
                desc += f", 중점 {info['center']}"
            lines.append(desc)
        except Exception as e:  # noqa: BLE001 — 이 모서리만 건너뛴다
            lines.append(f"  [{idx}] (측정 실패: {type(e).__name__})")
    sel.Clear()  # 조회 뒤 선택을 비워 사용자 화면을 방해하지 않는다
    return _truncate("\n".join(lines))


@mcp.tool()
@catia_tool
def select_edge(index: int, add: bool = False) -> str:
    """list_edges의 [인덱스] 모서리를 CATIA 현재 선택으로 잡습니다. (🟡)

    잡은 뒤 fillet_selected_edges(radius=...)로 라운드를 주거나 chamfer_selected_edges로
    모따기합니다. 여러 모서리를 함께 잡으려면 add=True로 이어 부릅니다.

    ⚠ list_edges를 부른 직후 바로 사용하세요(편집하면 인덱스가 바뀝니다).

    Args:
        index: list_edges에서 확인한 모서리 인덱스(0부터).
        add: True면 기존 선택에 더한다(여러 모서리 동시 선택). False면 이 모서리만.

    Returns:
        선택된 모서리의 길이·중점 요약과 현재 선택 개수.
    """
    catia = _catia()
    doc, _part = _active_part(catia)
    sel = doc.Selection
    # 모서리 검색은 선택을 덮어쓰므로, add=True면 이미 잡아 둔 참조를 먼저 보존한다.
    keep = [sel.Item2(i).Reference for i in range(1, sel.Count + 1)] if add else []
    n = _search_edges(sel)
    if n == 0:
        sel.Clear()
        return "모서리를 찾지 못했습니다. list_edges로 먼저 확인하세요."
    if not 0 <= index < n:
        sel.Clear()
        return f"인덱스 {index}가 범위를 벗어났습니다(모서리 {n}개, 0~{n - 1}). list_edges로 다시 확인하세요."
    target = sel.Item2(index + 1).Reference
    try:
        info = _measure_edge(doc.GetWorkbench("SPAWorkbench"), target)
    except Exception:  # noqa: BLE001 — 요약 정보는 부가라 실패해도 선택은 진행
        info = {"length": None, "center": None}
    sel.Clear()
    for r in keep:
        try:
            sel.Add(r)
        except Exception:  # noqa: BLE001 — 보존 실패한 참조는 건너뛴다
            pass
    sel.Add(target)
    tail = (f" 길이 {info['length']}" if info.get("length") is not None else "")
    if info.get("center"):
        tail += f", 중점 {info['center']}"
    return (
        f"모서리 [{index}] 선택됨{tail}. (현재 선택 {sel.Count}개)\n"
        "이제 fillet_selected_edges(radius=...)로 라운드를 주거나 "
        "chamfer_selected_edges(length=...)로 모따기하세요."
    )


# ─────────── 스케치 → 솔리드 피처 (비파괴 — Undo 가능, confirm 불필요) ───────────
# 스케치 프로파일로 MainBody에 재료를 만들거나 파낸다. pad(돌출)·pocket(파냄)·
# shaft(회전 돌출)·groove(회전 파냄)·fillet_selected_edges(모서리 라운드).
# pocket/groove/fillet은 기존 솔리드가 있어야 의미가 있다.
# pad/pocket은 스케치 평면의 법선을 따라 방향이 정해진다 — reverse로 뒤집는다.


def _apply_reverse(feat, reverse: bool) -> str:
    """reverse=True면 Pad/Pocket(둘 다 Prism)의 재료 방향을 뒤집고 결과 노트를 돌려준다.

    ⚠ 실기 검증 대상 — DirectionOrientation의 정수값이 개발 PC에 CATIA가 없어
    미검증이다. CATIA 관례상 스케치 평면 법선 기준 0↔1로 토글되므로(정방향/역방향)
    현재값을 읽어 뒤집는다. 속성명·값이 다르면 이 함수만 고치면 된다. 적용에
    실패해도 예외로 죽지 않고(우아한 저하) 노트로 알린다 — 피처 자체는 정상 생성된다.
    """
    if not reverse:
        return ""
    try:
        cur = int(feat.DirectionOrientation)
        feat.DirectionOrientation = 1 - cur  # 0↔1 토글 (법선 반대편으로)
        return " 방향반전(reverse)"
    except Exception:  # noqa: BLE001 — 이 버전/속성에서 안 되면 노트만 남긴다
        return " ⚠방향반전 미적용 — DirectionOrientation API를 실기에서 확인하세요"


@mcp.tool()
@catia_tool
def pad(length: float, sketch: str = "", second_length: float = 0, mirrored: bool = False,
        reverse: bool = False, name: str = "") -> str:
    """스케치 프로파일을 스케치 평면의 법선 방향으로 돌출(Pad)해 솔리드를 만듭니다. (🟡)

    돌출 방향은 스케치가 얹힌 평면의 법선이다(xy→+Z, yz→+X, zx→+Y). 튀어나오는
    쪽이 원하는 방향과 반대면 reverse=True로 뒤집는다. 여러 스케치를 그렸다면
    어느 것을 돌출할지 sketch= 인자로 이름을 반드시 명시하라 — 비우면 '가장 최근
    스케치'가 잡혀 엉뚱한 프로파일을 물 수 있다.

    Args:
        length: 돌출 길이(mm).
        sketch: 프로파일 스케치 이름. 비우면 가장 최근 스케치. 닫힌 프로파일이어야 한다.
        second_length: 반대 방향으로도 돌출할 길이(mm). 0이면 한쪽만.
        mirrored: True면 스케치 평면 기준 양쪽 대칭 돌출(second_length 무시).
        reverse: True면 돌출 방향(법선)을 반대로 뒤집는다.
        name: 피처에 붙일 이름(비우면 CATIA 자동).

    Returns:
        만들어진 Pad 피처의 최종 이름.
    """
    catia = _catia()
    _, part = _active_part(catia)
    sk = _get_sketch(part, sketch)
    part.InWorkObject = part.MainBody
    feat = _shape_factory(part).AddNewPad(sk, float(length))
    if mirrored:
        feat.IsSymmetric = True
    elif float(second_length):
        feat.SecondLimit.Dimension.Value = float(second_length)
    rev_note = _apply_reverse(feat, reverse)
    nm = _apply_name(feat, name)
    part.Update()
    extra = " 대칭" if mirrored else (f" +반대쪽 {second_length}mm" if float(second_length) else "")
    return f"Pad 생성: {nm} (스케치 {_safe(sk, 'Name')}, 길이 {length}mm{extra}){rev_note}"


@mcp.tool()
@catia_tool
def pocket(depth: float, sketch: str = "", mirrored: bool = False, reverse: bool = False, name: str = "") -> str:
    """스케치 프로파일 모양으로 스케치 평면의 법선을 따라 솔리드를 파냅니다(Pocket). (🟡)

    기존 솔리드(Pad 등)가 있어야 파낼 재료가 있다. 파고드는 쪽이 반대면 reverse=True로
    뒤집는다. 여러 스케치를 그렸다면 파낼 스케치를 sketch= 인자로 반드시 명시하라 —
    비우면 '가장 최근 스케치'가 잡혀 엉뚱한 구멍을 팔 수 있다.

    Args:
        depth: 파낼 깊이(mm).
        sketch: 프로파일 스케치 이름. 비우면 가장 최근 스케치.
        mirrored: True면 스케치 평면 기준 양쪽 대칭으로 파낸다.
        reverse: True면 파내는 방향(법선)을 반대로 뒤집는다.
        name: 피처에 붙일 이름(비우면 CATIA 자동).

    Returns:
        만들어진 Pocket 피처의 최종 이름.
    """
    catia = _catia()
    _, part = _active_part(catia)
    sk = _get_sketch(part, sketch)
    part.InWorkObject = part.MainBody
    feat = _shape_factory(part).AddNewPocket(sk, float(depth))
    if mirrored:
        feat.IsSymmetric = True
    rev_note = _apply_reverse(feat, reverse)
    nm = _apply_name(feat, name)
    part.Update()
    return f"Pocket 생성: {nm} (스케치 {_safe(sk, 'Name')}, 깊이 {depth}mm{' 대칭' if mirrored else ''}{rev_note})"


def _revolve(kind: str, sketch: str, angle: float, second_angle: float, axis: str, name: str) -> str:
    """shaft(회전 돌출)/groove(회전 파냄) 공통 구현. 반환: 결과 요약 문자열.

    axis가 있으면 스케치 안 그 선을 회전축(CenterLine)으로 지정한다. 없으면 스케치에
    이미 축이 정의돼 있어야 하며, 없을 경우 CATIA 오류가 안내로 반환된다.
    """
    catia = _catia()
    _, part = _active_part(catia)
    sk = _get_sketch(part, sketch)
    if axis:
        try:
            axis_elem = sk.GeometricElements.Item(axis)
        except pythoncom.com_error:
            raise CatiaError(
                f"스케치에서 축으로 쓸 '{axis}' 요소를 찾지 못했습니다. "
                "list_sketch_geometry로 이름을 확인하세요."
            )
        sk.CenterLine = axis_elem  # 이 선을 스케치의 회전축으로 지정
    part.InWorkObject = part.MainBody
    sf = _shape_factory(part)
    feat = sf.AddNewShaft(sk) if kind == "shaft" else sf.AddNewGroove(sk)
    feat.FirstAngle.Value = float(angle)
    feat.SecondAngle.Value = float(second_angle)
    nm = _apply_name(feat, name)
    part.Update()
    label = "Shaft(회전 돌출)" if kind == "shaft" else "Groove(회전 파냄)"
    return f"{label} 생성: {nm} (스케치 {_safe(sk, 'Name')}, {angle}°)"


@mcp.tool()
@catia_tool
def shaft(sketch: str = "", angle: float = 360, second_angle: float = 0, axis: str = "", name: str = "") -> str:
    """스케치 프로파일을 축 둘레로 회전시켜 솔리드를 만듭니다(Shaft/Revolve). (🟡)

    프로파일은 축의 한쪽에만 있어야 한다(축을 가로지르면 CATIA 오류).

    Args:
        sketch: 프로파일 스케치 이름. 비우면 가장 최근 스케치.
        angle: 회전 각도(도, 기본 360 = 온전한 회전체).
        second_angle: 반대 방향 회전 각도(도, 기본 0).
        axis: 회전축으로 쓸 스케치 안 선 이름(예: sketch_line으로 그린 'axisLine').
              비우면 스케치에 이미 정의된 축을 쓴다(없으면 오류 안내).
        name: 피처에 붙일 이름(비우면 CATIA 자동).

    Returns:
        만들어진 Shaft 피처의 최종 이름.
    """
    return _revolve("shaft", sketch, angle, second_angle, axis, name)


@mcp.tool()
@catia_tool
def groove(sketch: str = "", angle: float = 360, second_angle: float = 0, axis: str = "", name: str = "") -> str:
    """스케치 프로파일을 축 둘레로 회전시켜 솔리드를 파냅니다(Groove). (🟡)

    기존 솔리드가 있어야 파낼 재료가 있다. 인자는 shaft와 동일.

    Args:
        sketch: 프로파일 스케치 이름. 비우면 가장 최근 스케치.
        angle: 회전 각도(도, 기본 360).
        second_angle: 반대 방향 회전 각도(도, 기본 0).
        axis: 회전축으로 쓸 스케치 안 선 이름. 비우면 스케치에 정의된 축 사용.
        name: 피처에 붙일 이름(비우면 CATIA 자동).

    Returns:
        만들어진 Groove 피처의 최종 이름.
    """
    return _revolve("groove", sketch, angle, second_angle, axis, name)


@mcp.tool()
@catia_tool
def fillet_selected_edges(radius: float, name: str = "") -> str:
    """지금 CATIA 화면에서 선택된 모서리(들)에 라운드(필렛)를 줍니다. (🟡)

    모서리는 자동화로 이름 지목이 어려워 화면 선택을 쓴다(new_sketch의 selection과
    같은 방식): CATIA에서 모서리를 (Ctrl 클릭으로 여러 개) 선택한 뒤 호출한다.

    Args:
        radius: 필렛 반지름(mm, 0보다 커야 함).
        name: 피처에 붙일 이름(비우면 CATIA 자동).

    Returns:
        만들어진 필렛 피처의 최종 이름.
    """
    if float(radius) <= 0:
        raise CatiaError("반지름(radius)은 0보다 커야 합니다.")
    catia = _catia()
    doc, part = _active_part(catia)
    refs = _selected_refs(doc, "필렛을 줄 모서리")
    part.InWorkObject = part.MainBody
    # ⚠ 실기 검증 대상: 전파 모드 1 = 접선 연속(tangency) 전파.
    feat = _shape_factory(part).AddNewSolidEdgeFilletWithConstantRadius(refs[0], 1, float(radius))
    for r in refs[1:]:
        feat.AddObjectToFillet(r)
    nm = _apply_name(feat, name)
    part.Update()
    return f"필렛 생성: {nm} (모서리 {len(refs)}개, R{radius}mm)"


def _selected_refs(doc, what: str) -> list:
    """지금 CATIA 화면에서 선택된 요소들의 Reference 목록. 비어 있으면 안내와 함께 실패."""
    sel = doc.Selection
    if sel.Count == 0:
        raise CatiaError(f"CATIA 화면에서 {what}을(를) 먼저 선택한 뒤 다시 호출하세요.")
    return [sel.Item2(i).Reference for i in range(1, sel.Count + 1)]


@mcp.tool()
@catia_tool
def chamfer_selected_edges(length: float, angle: float = 45, name: str = "") -> str:
    """지금 CATIA 화면에서 선택된 모서리(들)를 모따기(챔퍼)합니다. (🟡)

    fillet_selected_edges와 같은 방식: CATIA에서 모서리를 (Ctrl 클릭으로 여러 개)
    선택한 뒤 호출한다.

    Args:
        length: 모따기 길이(mm, 0보다 커야 함).
        angle: 모따기 각도(도, 기본 45).
        name: 피처에 붙일 이름(비우면 CATIA 자동).

    Returns:
        만들어진 챔퍼 피처의 최종 이름.
    """
    if float(length) <= 0:
        raise CatiaError("모따기 길이(length)는 0보다 커야 합니다.")
    catia = _catia()
    doc, part = _active_part(catia)
    refs = _selected_refs(doc, "모따기할 모서리")
    part.InWorkObject = part.MainBody
    # ⚠ 실기 검증 대상: 전파 1=접선 연속, 모드 0=길이+각도, 방향 0=기본.
    feat = _shape_factory(part).AddNewChamfer(refs[0], 1, 0, 0, float(length), float(angle))
    for r in refs[1:]:
        feat.AddObjectToChamfer(r)
    nm = _apply_name(feat, name)
    part.Update()
    return f"모따기 생성: {nm} (모서리 {len(refs)}개, {length}mm × {angle}°)"


@mcp.tool()
@catia_tool
def shell_selected_faces(thickness: float, outer_thickness: float = 0, name: str = "") -> str:
    """솔리드의 속을 비웁니다(쉘). 지금 CATIA 화면에서 선택된 면(들)이 뚫립니다. (🟡)

    CATIA에서 제거할(열어 둘) 면을 선택한 뒤 호출한다. 남는 벽 두께를 지정한다.

    Args:
        thickness: 안쪽으로 남길 벽 두께(mm, 0보다 커야 함).
        outer_thickness: 바깥쪽으로 더할 두께(mm, 기본 0).
        name: 피처에 붙일 이름(비우면 CATIA 자동).

    Returns:
        만들어진 쉘 피처의 최종 이름.
    """
    if float(thickness) <= 0:
        raise CatiaError("벽 두께(thickness)는 0보다 커야 합니다.")
    catia = _catia()
    doc, part = _active_part(catia)
    refs = _selected_refs(doc, "제거할(열어 둘) 면")
    part.InWorkObject = part.MainBody
    feat = _shape_factory(part).AddNewShell(refs[0], float(thickness), float(outer_thickness))
    for r in refs[1:]:
        feat.AddFaceToRemove(r)
    nm = _apply_name(feat, name)
    part.Update()
    return f"쉘 생성: {nm} (면 {len(refs)}개 제거, 두께 {thickness}mm)"


@mcp.tool()
@catia_tool
def mirror_body(plane: str, name: str = "") -> str:
    """현재 솔리드(MainBody)를 평면 기준으로 대칭 복제합니다(Mirror). (🟡)

    Args:
        plane: 대칭 기준 평면 — 'xy'/'yz'/'zx' 또는 평면 요소 이름(plane_offset으로 만든 것).
        name: 피처에 붙일 이름(비우면 CATIA 자동).

    Returns:
        만들어진 미러 피처의 최종 이름.
    """
    catia = _catia()
    _, part = _active_part(catia)
    part.InWorkObject = part.MainBody
    feat = _shape_factory(part).AddNewMirror(_part_ref(part, plane))
    nm = _apply_name(feat, name)
    part.Update()
    return f"미러 생성: {nm} (기준 평면: {plane})"


# ─────────── 🟡 어셈블리 (Product — 로컬 생성·이동, 비파괴) ───────────
# Product 문서에 부품/하위 제품을 넣고 배치한다. 디스크에 쓰지 않고 Undo 가능.
# 조회는 product_tree(🟢), 저장은 save_document(🔴)가 담당한다.


@mcp.tool()
@catia_tool
def new_product() -> str:
    """새 Product(.CATProduct, 어셈블리) 문서를 만듭니다. (🟡)

    만들어진 제품은 활성 문서가 되어 add_component로 부품을 넣을 수 있다.
    """
    catia = _catia()
    doc = catia.Documents.Add("Product")
    return (
        f"새 Product 생성됨: {_safe(doc, 'Name')}. "
        "(활성 문서로 설정됨 — add_component로 부품 파일을 추가하세요.)"
    )


@mcp.tool()
@catia_tool
def add_component(file_path: str) -> str:
    """활성 Product에 부품/제품 파일(.CATPart/.CATProduct)을 구성요소로 추가합니다. (🟡)

    Args:
        file_path: 추가할 파일의 절대경로.

    Returns:
        추가된 구성요소의 인스턴스 이름(move_component에서 이 이름으로 지목).
    """
    catia = _catia()
    doc = _active_document(catia)
    if _doc_type(doc) != "Product":
        raise CatiaError(
            "활성 문서가 Product(.CATProduct)가 아닙니다. new_product로 만들거나 CATIA에서 "
            "어셈블리를 활성화한 뒤 다시 시도하세요."
        )
    p = os.path.abspath(os.path.expanduser(file_path))
    if not os.path.exists(p):
        raise CatiaError(f"'{p}' 경로에 파일이 없습니다. 경로를 확인하세요.")
    prods = doc.Product.Products
    before = prods.Count
    try:
        prods.AddComponentsFromFiles([p], "All")
    except (pythoncom.com_error, TypeError):
        # ⚠ 실기 검증 대상: 일부 pywin32/CATIA 조합은 파이썬 리스트를 SAFEARRAY로
        # 받지 못한다. 명시적 VARIANT(BSTR 배열)로 한 번 더 시도한다.
        arr = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR, [p])
        prods.AddComponentsFromFiles(arr, "All")
    added = [_safe(prods.Item(i), "Name") for i in range(before + 1, prods.Count + 1)]
    doc.Product.Update()
    return f"구성요소 추가됨: {', '.join(added) or '(이름 확인 실패)'}  ← {os.path.basename(p)}"


@mcp.tool()
@catia_tool
def move_component(instance: str, dx: float = 0, dy: float = 0, dz: float = 0) -> str:
    """활성 Product의 구성요소를 지정 거리만큼 평행 이동합니다. (🟡)

    회전은 지원하지 않는다(위치 행렬의 이동 성분만 바꾼다). 인스턴스 이름은
    product_tree 또는 add_component의 반환값으로 확인한다.

    Args:
        instance: 구성요소 인스턴스 이름(예: 'Part1.1').
        dx, dy, dz: 이동량(mm).

    Returns:
        이동 후 절대 위치.
    """
    catia = _catia()
    doc = _active_document(catia)
    if _doc_type(doc) != "Product":
        raise CatiaError("활성 문서가 Product(.CATProduct)가 아닙니다.")
    try:
        item = doc.Product.Products.Item(instance)
    except pythoncom.com_error:
        raise CatiaError(
            f"구성요소 '{instance}'를 찾지 못했습니다. product_tree로 인스턴스 이름을 확인하세요."
        )
    pos = item.Position
    try:
        # ⚠ 실기 검증 대상: GetComponents의 out-배열이 pywin32에서 반환값으로 나오는지.
        comps = list(pos.GetComponents())
    except Exception as e:  # noqa: BLE001
        raise CatiaError(
            f"위치 행렬을 읽지 못했습니다({type(e).__name__}). 이 환경의 pywin32에서는 "
            "구성요소 이동이 지원되지 않을 수 있습니다 — CATIA에서 직접 이동하세요."
        )
    if len(comps) < 12:
        raise CatiaError(f"위치 행렬 형식이 예상과 다릅니다(길이 {len(comps)}).")
    comps[9] += float(dx)
    comps[10] += float(dy)
    comps[11] += float(dz)
    pos.SetComponents(comps)
    doc.Product.Update()
    return (
        f"이동 완료: {instance} (+{dx}, +{dy}, +{dz})\n"
        f"  새 위치: ({round(comps[9], 3)}, {round(comps[10], 3)}, {round(comps[11], 3)})"
    )


# ─────────── 측정(🟢)·뷰·유틸(🟡) ───────────


@mcp.tool()
@catia_tool
def measure_element(name: str) -> str:
    """파트 요소(스케치/피처/3D 요소)의 길이·면적·부피를 측정합니다. (🟢 읽기)

    SPAWorkbench 측정기를 쓴다. 요소 종류에 따라 의미 있는 값만 나온다(곡선→길이,
    면/서피스→면적, 솔리드→부피). ⚠ 단위는 CATIA 문서 설정을 따른다(보통 mm 계열) —
    값 크기가 이상하면 단위 환산을 의심할 것.

    Args:
        name: 측정할 요소 이름(list_3d_geometry/list_sketch_geometry로 확인).
    """
    t = (name or "").strip()
    if not t:
        raise CatiaError("측정할 요소 이름(name)을 지정하세요.")
    catia = _catia()
    doc, part = _active_part(catia)
    target, where = _find_part_object(part, t)
    ref = part.CreateReferenceFromObject(target)
    try:
        meas = doc.GetWorkbench("SPAWorkbench").GetMeasurable(ref)
    except pythoncom.com_error as e:
        raise CatiaError(f"측정기를 열지 못했습니다: {_com_message(e)}")
    lines = [f"측정: {_safe(target, 'Name') or t} ({where})"]
    found = False
    for attr, label in (("Length", "길이"), ("Area", "면적"), ("Volume", "부피")):
        try:
            v = float(getattr(meas, attr))
        except Exception:  # noqa: BLE001 — 이 요소에 없는 측정값이면 건너뛴다
            continue
        if v:
            lines.append(f"  {label}: {round(v, 4)}")
            found = True
    if not found:
        lines.append("  (이 요소에서 길이/면적/부피를 얻지 못했습니다.)")
    return "\n".join(lines)


@mcp.tool()
@catia_tool
def measure_body(body: str = "") -> str:
    """바디(솔리드)의 부피·면적·질량·밀도·무게중심을 측정합니다. (🟢 읽기)

    질량·밀도는 파트에 재질(material)이 적용돼 있어야 의미 있는 값이 나온다
    (기본 밀도 1000kg/m³로 계산될 수 있음). ⚠ 단위는 CATIA 설정을 따른다.
    ⚠ 질량 계산에 쓰는 Inertias.Add는 관성 엔티티를 스펙 트리에 추가하므로, 읽은
    직후 그 엔티티를 제거해 문서를 원상태로 되돌린다(아래 주석 참고).

    Args:
        body: 바디 이름. 비우면 메인 바디(PartBody).
    """
    catia = _catia()
    doc, part = _active_part(catia)
    if body:
        try:
            b = part.Bodies.Item(body)
        except pythoncom.com_error:
            raise CatiaError(f"바디 '{body}'를 찾지 못했습니다. part_summary로 확인하세요.")
    else:
        b = part.MainBody
    spa = doc.GetWorkbench("SPAWorkbench")
    lines = [f"바디 측정: {_safe(b, 'Name')}"]
    try:
        meas = spa.GetMeasurable(part.CreateReferenceFromObject(b))
        for attr, label in (("Volume", "부피"), ("Area", "표면적")):
            try:
                lines.append(f"  {label}: {round(float(getattr(meas, attr)), 4)}")
            except Exception:  # noqa: BLE001
                pass
    except Exception:  # noqa: BLE001 — 측정기 실패 시 관성 값만이라도 시도한다
        pass
    inertia = None
    try:
        inertia = spa.Inertias.Add(b)
        lines.append(f"  질량: {round(float(inertia.Mass), 6)} kg")
        lines.append(f"  밀도: {round(float(inertia.Density), 3)} kg/m³")
        try:
            # COG는 out-배열 인자라 pywin32에서 VARIANT byref로 받아야 한다.
            cog = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8 | pythoncom.VT_BYREF, [0.0, 0.0, 0.0]
            )
            inertia.GetCOGPosition(cog)
            x, y, z = (round(float(v), 4) for v in cog.value)
            lines.append(f"  무게중심: ({x}, {y}, {z})")
        except Exception:  # noqa: BLE001 — COG는 부가정보, 실패해도 넘어간다
            pass
    except Exception:  # noqa: BLE001
        pass
    finally:
        # Inertias.Add는 관성 엔티티를 문서 스펙 트리에 남겨(문서가 '변경됨' 상태가
        # 됨) 🟢 읽기 계약을 깬다. 방금 추가한 마지막 관성을 컬렉션에서 제거해 반복
        # 측정 때 트리에 쌓이지 않게 한다. ⚠ 실기 검증 대상: Inertias.Remove의 존재·
        # 시그니처(인덱스 1-based 가정)가 미검증이다. 없으면 조용히 남겨둔다(측정값은
        # 이미 얻었다). 솔리드를 지울 위험이 있는 Selection.Delete는 의도적으로 안 쓴다.
        if inertia is not None:
            try:
                spa.Inertias.Remove(spa.Inertias.Count)
            except Exception:  # noqa: BLE001 — 제거 API가 없으면 남겨둔다
                pass
    if len(lines) == 1:
        lines.append("  (측정값을 얻지 못했습니다 — 솔리드가 비어 있는지 확인하세요.)")
    return "\n".join(lines)


@mcp.tool()
@catia_tool
def capture_view(out_path: str, fit: bool = True) -> str:
    """현재 3D 뷰를 PNG 이미지로 캡처합니다. (🟡)

    모델을 만들거나 바꾼 뒤 결과를 사용자에게 보여줄 때 쓴다. 새 파일만 만들고
    **기존 파일은 절대 덮어쓰지 않으므로**(있으면 오류로 물러섬) 비파괴로 취급한다.

    Args:
        out_path: 출력 PNG 경로(예: C:/tmp/view.png). 이미 존재하면 실패한다.
        fit: True면 캡처 전에 모델 전체가 보이도록 화면을 맞춘다(Reframe).
    """
    p = os.path.abspath(os.path.expanduser(out_path))
    if not p.lower().endswith(".png"):
        p += ".png"
    if os.path.exists(p):
        raise CatiaError(f"'{p}' 파일이 이미 있습니다. 덮어쓰지 않으니 다른 경로를 지정하세요.")
    out_dir = os.path.dirname(p)
    if out_dir and not os.path.isdir(out_dir):
        try:
            os.makedirs(out_dir, exist_ok=True)  # 없는 상위 폴더까지 만든다
        except OSError as e:
            raise CatiaError(f"출력 폴더를 만들지 못했습니다: {out_dir} ({e})")
    catia = _catia()
    try:
        viewer = catia.ActiveWindow.ActiveViewer
    except pythoncom.com_error as e:
        raise CatiaError(f"활성 뷰어를 얻지 못했습니다: {_com_message(e)}")
    if fit:
        try:
            viewer.Reframe()
        except Exception:  # noqa: BLE001 — 화면 맞춤 실패해도 캡처는 진행
            pass
    viewer.CaptureToFile(4, p)  # ⚠ 실기 검증 대상: 4 = PNG (CatCaptureFormat)
    return f"캡처 완료: {p}"


@mcp.tool()
@catia_tool
def undo(count: int = 1) -> str:
    """CATIA의 실행 취소(Undo)를 지정 횟수만큼 수행합니다. (🟡)

    🟡 도구들이 만든 요소를 잘못 만들었을 때 삭제 대신 되돌리는 용도.
    ⚠ 실기 검증 대상: COM에 Undo API가 없어 StartCommand('Undo')로 UI 커맨드를
    보낸다 — 환경에 따라 동작하지 않을 수 있다(그 경우 CATIA에서 Ctrl+Z).

    Args:
        count: 되돌릴 횟수(1~20).
    """
    n = max(1, min(int(count), 20))
    catia = _catia()
    for _ in range(n):
        catia.StartCommand("Undo")
    return f"실행 취소 {n}회 요청함. list_sketch_geometry/list_3d_geometry로 현재 상태를 확인하세요."


@mcp.tool()
@catia_tool
def redo(count: int = 1) -> str:
    """CATIA의 다시 실행(Redo)을 지정 횟수만큼 수행합니다. (🟡)

    undo로 되돌린 작업을 다시 적용하는 용도. undo와 짝을 이룬다.
    ⚠ 실기 검증 대상: COM에 Redo API가 없어 StartCommand('Redo')로 UI 커맨드를
    보낸다 — 환경에 따라 동작하지 않을 수 있다(그 경우 CATIA에서 Ctrl+Y).

    Args:
        count: 다시 실행할 횟수(1~20).
    """
    n = max(1, min(int(count), 20))
    catia = _catia()
    for _ in range(n):
        catia.StartCommand("Redo")
    return f"다시 실행 {n}회 요청함. list_sketch_geometry/list_3d_geometry로 현재 상태를 확인하세요."


@mcp.tool()
@catia_tool
def set_visibility(name: str, visible: bool) -> str:
    """파트 요소(스케치/피처/3D 요소/세트)를 화면에서 숨기거나 다시 보이게 합니다. (🟡)

    삭제가 아니라 표시 전환이다(모델은 그대로). 보조 형상(평면·와이어프레임)을
    정리해 화면을 깔끔하게 할 때 쓴다.

    Args:
        name: 대상 요소 이름.
        visible: True면 보이기, False면 숨기기.
    """
    catia = _catia()
    doc, part = _active_part(catia)
    target, where = _find_part_object(part, name)
    sel = doc.Selection
    sel.Clear()
    sel.Add(target)
    sel.VisProperties.SetShow(0 if visible else 1)  # 0=Show, 1=NoShow
    sel.Clear()
    label = _safe(target, "Name") or name
    return f"{'표시' if visible else '숨김'}: {label} ({where})"


@mcp.tool()
@catia_tool
def set_color(name: str, r: int, g: int, b: int) -> str:
    """파트 요소(피처/3D 요소/바디 등)의 색을 바꿉니다. (🟡)

    Args:
        name: 대상 요소 이름.
        r, g, b: 색상(0~255).
    """
    for v, lab in ((r, "r"), (g, "g"), (b, "b")):
        if not 0 <= int(v) <= 255:
            raise CatiaError(f"{lab} 값은 0~255 사이여야 합니다 (받은 값: {v}).")
    catia = _catia()
    doc, part = _active_part(catia)
    target, where = _find_part_object(part, name)
    sel = doc.Selection
    sel.Clear()
    sel.Add(target)
    sel.VisProperties.SetRealColor(int(r), int(g), int(b), 1)
    sel.Clear()
    label = _safe(target, "Name") or name
    return f"색 변경: {label} ({where}) → RGB({r}, {g}, {b})"


# ═══════════ 🔴 수정·저장·내보내기·종료 (되돌리기 어려움 — confirm 게이팅) ═══════════
# 이 구역의 도구는 모델을 바꾸거나 디스크에 파일을 쓰거나 CATIA를 종료합니다. 모두
# confirm=True를 받아야 실행되며, confirm 없이 부르면 무엇을 할지 프리뷰만 돌려줍니다.


@mcp.tool()
@catia_tool
def quit_catia(confirm: bool = False) -> str:
    """🔴 CATIA를 완전히 종료합니다(열린 모든 문서가 닫힘). (confirm=True 필요)

    저장하지 않은 문서가 있으면 그 변경 내용이 사라질 수 있습니다. confirm 없이
    부르면 열린 문서와 저장 안 된 문서를 요약한 프리뷰만 돌려줍니다.

    Args:
        confirm: 실제 종료하려면 True. 없으면 프리뷰만.
    """
    catia = _catia()
    docs = catia.Documents
    count = docs.Count
    unsaved = []
    for i in range(1, count + 1):
        d = docs.Item(i)
        saved = _safe(d, "Saved")
        if saved and saved not in ("True", "-1", "1"):
            unsaved.append(_safe(d, "Name"))
    if not confirm:
        details = [f"열린 문서: {count}개"]
        details.append("저장 안 됨: " + (", ".join(unsaved) if unsaved else "없음"))
        return _preview("CATIA 완전 종료(모든 문서 닫힘)", details, "(quit_catia confirm=true)")
    catia.Quit()
    return f"CATIA를 종료했습니다. (닫힌 문서 {count}개)"


def _find_in_sketch(sk, token: str):
    """스케치 안에서 삭제 대상(2D 요소 또는 구속)을 찾아 (객체, 위치 설명)으로 돌려준다."""
    label = _safe(sk, "Name")
    try:
        return sk.GeometricElements.Item(token), f"스케치 '{label}'의 2D 요소"
    except pythoncom.com_error:
        pass
    try:  # 구속(sketch_coincidence 등이 만든 'Coincidence.1' 같은 이름)
        return sk.Constraints.Item(token), f"스케치 '{label}'의 구속"
    except Exception:  # noqa: BLE001
        pass
    raise CatiaError(
        f"스케치 '{label}'에서 '{token}'를 찾지 못했습니다. list_sketch_geometry로 "
        "이름을 확인하세요."
    )


def _find_part_object(part, token: str):
    """파트 수준에서 요소를 찾아 (객체, 위치 설명)으로 돌려준다.

    delete_element·measure_element·set_visibility·set_color가 함께 쓴다.
    탐색 순서: 스케치 → 솔리드 피처(Pad/Pocket/Shaft/Groove/필렛/로프트) → 3D 요소
    → 기하학적 세트. (스케치 안 2D 요소는 _find_in_sketch 담당.)
    """
    try:
        bodies = part.Bodies
        for i in range(1, bodies.Count + 1):
            body = bodies.Item(i)
            label = _safe(body, "Name")
            try:
                return body.Sketches.Item(token), f"바디 '{label}'의 스케치"
            except Exception:  # noqa: BLE001 — 이 바디에 없으면 다음 후보
                pass
            try:
                return body.Shapes.Item(token), f"바디 '{label}'의 솔리드 피처"
            except Exception:  # noqa: BLE001
                pass
    except Exception:  # noqa: BLE001
        pass
    for hb, shape in _iter_hybrid_shapes(part):
        if _safe(shape, "Name") == token:
            return shape, f"기하학적 세트 '{_safe(hb, 'Name')}'의 3D 요소"
    try:
        return part.HybridBodies.Item(token), "기하학적 세트(안의 모든 요소 포함)"
    except Exception:  # noqa: BLE001
        pass
    raise CatiaError(
        f"파트에서 '{token}'를 찾지 못했습니다. 이름은 list_3d_geometry(3D 요소·피처)/"
        "list_sketch_geometry(스케치 안)로 확인하세요. (스케치 안 2D 요소는 파트 수준 "
        "탐색에 잡히지 않습니다 — 해당 도구의 sketch 인자를 쓰세요.)"
    )


@mcp.tool()
@catia_tool
def delete_element(name: str, sketch: str = "", confirm: bool = False) -> str:
    """🔴 파트의 요소(스케치·솔리드 피처·3D 요소·기하학적 세트) 또는 스케치 안
    2D 요소·구속을 이름으로 삭제합니다. (confirm=True 필요)

    CATIA Undo로 되돌릴 수는 있지만, 지운 요소를 참조하는 다른 요소(스케치를 쓰는
    Pad, 끝점에 걸린 구속 등)가 함께 깨지거나 지워질 수 있어 confirm 게이트를 둔다.
    confirm 없이 부르면 무엇이 어디에서 지워질지 찾은 결과만 보여준다.

    Args:
        name: 삭제할 요소 이름. sketch를 비우면 파트 수준(스케치/피처/3D 요소/세트)
              에서, sketch를 주면 그 스케치 안(2D 요소/구속)에서 찾는다.
        sketch: 스케치 안 요소를 지울 때 그 스케치 이름. 비우면 파트 수준 탐색.
                (다른 도구와 달리 여기서 비움은 '최근 스케치'가 아니라 '파트 수준'이다.)
        confirm: 실제 삭제하려면 True. 없으면 프리뷰만.
    """
    t = (name or "").strip()
    if not t:
        raise CatiaError("삭제할 요소 이름(name)을 지정하세요.")
    catia = _catia()
    doc, part = _active_part(catia)
    sk = _get_sketch(part, sketch) if sketch else None
    target, where = _find_in_sketch(sk, t) if sk is not None else _find_part_object(part, t)
    label = _safe(target, "Name") or t
    if not confirm:
        return _preview(
            "요소 삭제",
            [
                f"대상: {label}",
                f"위치: {where}",
                "주의: 이 요소를 참조하는 다른 요소(피처·구속)가 함께 깨지거나 지워질 수 있습니다.",
            ],
            "(delete_element ... confirm=true)",
        )
    sel = doc.Selection
    if sk is not None:
        sk.OpenEdition()  # 스케치 내부 요소는 편집 상태에서 지운다(_draw와 같은 관례)
        try:
            sel.Clear()
            sel.Add(target)
            sel.Delete()
        finally:
            sk.CloseEdition()
    else:
        sel.Clear()
        sel.Add(target)
        sel.Delete()
    _update(doc)  # 삭제로 참조가 깨져 재생성이 실패해도 삭제 자체는 이미 반영됨
    return f"삭제 완료: {label} ({where}). CATIA의 Undo로 되돌릴 수 있습니다."


@mcp.tool()
@catia_tool
def set_parameter(name: str, value: str, path: str | None = None, confirm: bool = False) -> str:
    """🔴 파라미터 값을 바꾸고 형상을 재생성합니다. (confirm=True 필요)

    숫자로 해석되면 실수(float)로, 아니면 문자열로 설정합니다. 길이 파라미터의 값은
    CATIA의 현재 단위계를 따릅니다(보통 mm) — 단위가 헷갈리면 먼저 get_parameter로
    현재 표기를 확인하세요.

    Args:
        name: 바꿀 파라미터 이름(정확히).
        value: 설정할 값. 숫자 문자열이면 수치로, 아니면 문자열로 넣습니다.
        path: 대상 문서. 생략하면 활성 문서.
        confirm: 실제 변경하려면 True. 없으면 변경 전/후 프리뷰만.

    Returns:
        변경 결과, 또는 (confirm 없을 때) 프리뷰.
    """
    catia = _catia()
    with _document(catia, path) as doc:
        params = _parameter_container(doc)
        try:
            p = params.Item(name)
        except pythoncom.com_error:
            return f"파라미터 '{name}'를 찾지 못했습니다. list_parameters로 이름을 확인하세요."
        old = _param_value(p)
        if not confirm:
            return _preview(
                "파라미터 변경",
                [f"문서: {_doc_label(doc)}", f"파라미터: {name}", f"현재값: {old}", f"바꿀값: {value}"],
                "(set_parameter ... confirm=true)",
            )
        # 값 강제 변환: 불리언 문자열 → bool, 숫자 → float, 그 외 → 문자열 그대로.
        # (불리언 파라미터에 float("true")를 넣던 기존 동작 보완.) 파라미터 타입과
        # 안 맞으면 com_error가 나므로, 마지막에 사유를 붙여 안내한다.
        low = value.strip().lower()
        try:
            if low in ("true", "false"):
                p.Value = (low == "true")
            else:
                p.Value = float(value)
        except (ValueError, pythoncom.com_error):
            try:
                p.Value = value  # 숫자/불리언이 아니면 문자열 파라미터로 시도
            except pythoncom.com_error as e:
                return (
                    f"파라미터 '{name}' 값을 '{value}'로 설정하지 못했습니다: {_com_message(e)}. "
                    "타입이 맞지 않을 수 있습니다 — get_parameter로 현재 표기를 확인하세요."
                )
        _update(doc)
        return f"변경 완료.\n  문서: {_doc_label(doc)}\n  {name}: {old} → {_param_value(p)}"


@mcp.tool()
@catia_tool
def save_document(path: str | None = None, confirm: bool = False) -> str:
    """🔴 문서를 현재 경로에 저장합니다(덮어쓰기). (confirm=True 필요)

    Args:
        path: 대상 문서. 생략하면 활성 문서. (아직 저장된 적 없는 새 문서는 경로가
              없어 저장할 수 없습니다 — CATIA에서 먼저 이름을 지정해 저장하세요.)
        confirm: 실제 저장하려면 True. 없으면 프리뷰만.
    """
    catia = _catia()
    # 저장 대상은 '이미 열려 있는' 문서여야 의미가 있다 — 우리가 열어 저장 후 닫는 것은
    # 사용자 의도와 어긋나므로, path가 열려 있지 않으면 _document가 열어버리기 전에 막지
    # 않고 그대로 저장을 지원한다(사용자가 경로를 명시했다면 그 파일을 저장하려는 의도).
    with _document(catia, path) as doc:
        full = _full_name(doc)
        if not full or not os.path.isabs(full):
            return (
                "이 문서는 아직 디스크에 저장된 적이 없어 경로가 없습니다. CATIA에서 "
                "먼저 '다른 이름으로 저장'해 경로를 정하세요."
            )
        if not confirm:
            return _preview(
                "문서 저장(덮어쓰기)",
                [f"문서: {_doc_label(doc)}", f"경로: {full}"],
                "(save_document ... confirm=true)",
            )
        doc.Save()
        return f"저장 완료.\n  경로: {full}"


@mcp.tool()
@catia_tool
def export_document(out_path: str, format: str, path: str | None = None, confirm: bool = False) -> str:
    """🔴 문서를 다른 형식으로 내보냅니다(디스크에 새 파일 생성). (confirm=True 필요)

    CATIA의 ExportData를 씁니다. 지원 형식은 설치된 번역기/라이선스에 따라 다릅니다.

    Args:
        out_path: 출력 파일 경로. 확장자는 붙여도 되고 없어도 됩니다(CATIA가 형식에 맞춰
                  붙입니다). 예: C:/tmp/part.stp
        format: 출력 형식 — stp/step, igs/iges, stl, wrl, cgr, model, 3dxml 중 하나.
        path: 대상 문서. 생략하면 활성 문서.
        confirm: 실제 내보내려면 True. 없으면 프리뷰만.
    """
    fmt = format.lower().strip()
    if fmt not in EXPORT_FORMATS:
        return (
            f"지원하지 않는 형식 '{format}'. 사용 가능: {', '.join(sorted(set(EXPORT_FORMATS)))}. "
            "(실제 지원 여부는 설치된 CATIA 번역기 라이선스에 따라 다릅니다.)"
        )
    # ExportData는 확장자 없는 기본 경로를 받고 형식에 맞춰 확장자를 붙인다.
    base, _ext = os.path.splitext(os.path.abspath(os.path.expanduser(out_path)))
    catia = _catia()
    with _document(catia, path) as doc:
        if not confirm:
            return _preview(
                f"{EXPORT_FORMATS[fmt]}로 내보내기",
                [f"문서: {_doc_label(doc)}", f"형식: {fmt} ({EXPORT_FORMATS[fmt]})", f"출력: {base}.{fmt}"],
                "(export_document ... confirm=true)",
            )
        doc.ExportData(base, fmt)
        return f"내보내기 완료.\n  형식: {EXPORT_FORMATS[fmt]}\n  출력: {base}.{fmt}"


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="CATIA V5 문서 읽기·수정 MCP 서버")
    parser.add_argument(
        "--transport",
        choices=["stdio", "http", "sse"],
        default=os.getenv("CATIA_MCP_TRANSPORT", "stdio"),
        help=(
            "stdio(기본): Claude 등 로컬 클라이언트가 프로세스를 직접 실행해 붙는다. "
            "http: n8n 등 네트워크 클라이언트가 URL로 접속한다. "
            "sse: 구버전 n8n MCP 노드가 Streamable HTTP를 못 쓸 때 사용한다."
        ),
    )
    parser.add_argument("--host", default=os.getenv("CATIA_MCP_HOST", "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.getenv("CATIA_MCP_PORT", "8089")))
    args = parser.parse_args()

    if not COM_AVAILABLE:
        print(f"경고: pywin32를 불러올 수 없습니다 ({COM_IMPORT_ERROR}).", file=sys.stderr)
        print("서버는 실행되지만 모든 도구가 안내 메시지만 반환합니다.", file=sys.stderr)

    if args.transport in ("http", "sse"):
        # 네트워크 접속용. COM 특성상, 반드시 사용자가 로그인한 그 세션에서 실행해야
        # 열린 CATIA가 보인다(서비스나 다른 세션에서는 안 보임).
        path = "/mcp/" if args.transport == "http" else "/sse/"
        url = f"http://{args.host}:{args.port}{path}"
        print(f"CATIA MCP 서버 시작 ({args.transport}) — {url}", file=sys.stderr)
        mcp.run(transport=args.transport, host=args.host, port=args.port)
    else:
        # stdio: stdout이 MCP 프로토콜 채널이므로 로그는 stderr로 보낸다.
        print("CATIA MCP 서버 시작 (stdio)", file=sys.stderr)
        mcp.run(transport="stdio")
