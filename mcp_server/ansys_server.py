"""ansys_server.py

ANSYS MAPDL 열해석(정상상태·과도)을 PyMAPDL(gRPC)로 조종하는 MCP 서버입니다.
CATIA 서버와 달리 COM이 아니므로 사용자 로그인 세션 제약이 없고, 원격 MAPDL
인스턴스에도 붙을 수 있습니다. 대신 **ANSYS MAPDL 설치 + 라이선스**가 필요합니다
(PyMAPDL은 클라이언트일 뿐, 솔버는 ANSYS 본체가 담당).

폐쇄망 반입 체크리스트:
    - ansys-mapdl-core 가 사내 미러에 있는지 확인 (grpcio·numpy 등 의존성이 많다).
      없으면 이 서버는 뜨되 모든 도구가 안내만 반환한다 (우아한 저하).
    - 회사 PC에 ANSYS MAPDL 설치 + 라이선스. 도구 호출 중 라이선스를 체크아웃한다.
    - MAPDL 기동은 수 초~수십 초 — launch_ansys로 한 번 띄워 상주시키는 구조다.

세션 모델:
    서버가 MAPDL 세션 하나를 전역으로 잡아둔다. launch_ansys(로컬 기동) 또는
    connect_ansys(이미 떠 있는 인스턴스 접속) 후 나머지 도구를 쓴다. 도구 호출은
    락으로 직렬화한다 (MAPDL은 한 번에 한 명령).

열해석 워크플로 (instructions에도 안내):
    launch_ansys → create_block/cylinder → set_thermal_element → define_thermal_material
    → mesh_model → apply_temperature/convection/heat_flux/heat_generation
    → solve_steady 또는 solve_transient → result_temperature/heat_flux/at_point
    → capture_plot(온도 컨투어 PNG)

안전 등급 (office/outlook/catia와 같은 3티어):
    🟢 읽기 — ansys_status, mesh_info, list_areas, get_parameter, result_*
    🟡 메모리 작업 — 세션 기동/접속, 형상·재료·메시·하중, solve, run_apdl,
       capture_plot(새 PNG 파일 하나 생성 — 기존 파일은 절대 덮어쓰지 않음),
       open_gui(현재 모델을 MAPDL GUI 창으로 직접 봄 — 블로킹, 로컬 세션 전용)
    🔴 파괴 — clear_model(모델 전체 삭제), shutdown_ansys(현재 세션 종료),
       kill_local_ansys(이 PC의 고아 MAPDL 전체 정리).
       confirm=True 없이는 실행되지 않는다.

정리(고아 방지): 서버가 정상 종료되면 MAPDL 세션을 자동으로 닫는다(atexit). 강제
종료로 MAPDL이 안 꺼지고 남으면 `python ansys_server.py --cleanup`(또는 세션 중
kill_local_ansys 도구)으로 이 PC의 로컬 gRPC 인스턴스를 정리한다.

⚠ 실기 검증 대상: 개발 PC에 ANSYS가 없어 APDL 커맨드 시퀀스와 *GET 조회는
실제 MAPDL에 대고 검증하지 못했다. 표시된 곳을 실기에서 확인할 것.

사용:
    python ansys_server.py                    # stdio (기본)
    python ansys_server.py --transport http   # n8n 등 네트워크용, :8091
"""

from __future__ import annotations

import argparse
import atexit
import os
import shutil
import sys
import threading
import time
from functools import wraps

from fastmcp import FastMCP

# PyMAPDL — 없으면 서버는 뜨고 모든 도구가 안내만 반환한다 (우아한 저하).
try:
    from ansys.mapdl import core as pymapdl

    MAPDL_AVAILABLE = True
    MAPDL_IMPORT_ERROR = ""
except Exception as e:  # noqa: BLE001 — grpcio 등 하위 의존성 실패 포함
    pymapdl = None  # type: ignore[assignment]
    MAPDL_AVAILABLE = False
    MAPDL_IMPORT_ERROR = str(e)

mcp = FastMCP(
    name="ansys",
    instructions=(
        "ANSYS MAPDL 열해석 MCP 서버입니다. 순서: ① launch_ansys(또는 connect_ansys)로 "
        "세션을 열고 ② create_block/create_cylinder로 형상 ③ set_thermal_element로 열해석 "
        "요소 ④ define_thermal_material로 열물성(전도율 등) ⑤ mesh_model로 메시 "
        "⑥ apply_temperature(고정 온도)/apply_convection(대류)/apply_heat_flux(열유속)/"
        "apply_heat_generation(발열)으로 경계조건 — 면 번호는 list_areas로 확인 "
        "⑦ solve_steady(정상상태) 또는 solve_transient(과도) ⑧ result_temperature/"
        "result_heat_flux/temperature_at_point로 결과 확인, capture_plot으로 컨투어 저장. "
        "solve는 모델 크기에 따라 오래 걸릴 수 있습니다. 모델 삭제(clear_model)와 "
        "종료(shutdown_ansys)는 confirm=True가 필요합니다."
    ),
)

# ─────────────────────────────── 상태/헬퍼 ───────────────────────────────

_mapdl = None  # 전역 MAPDL 세션 (launch/connect가 채운다)
_lock = threading.Lock()  # MAPDL은 한 번에 한 명령 — 도구 호출을 직렬화

MAX_CHARS = 4000


class AnsysError(Exception):
    """도구가 사용자에게 그대로 돌려줄 안내 메시지를 담은 예외."""


def ansys_tool(fn):
    """예외를 안내 문자열로 바꾸고, 세션 접근을 락으로 직렬화한다."""

    @wraps(fn)
    def wrapper(*args, **kwargs):
        with _lock:
            try:
                return fn(*args, **kwargs)
            except AnsysError as e:
                return str(e)
            except Exception as e:  # noqa: BLE001 — 도구는 항상 문자열을 돌려준다
                return f"작업에 실패했습니다: {type(e).__name__}: {e}"

    return wrapper


def _truncate(text: str, limit: int = MAX_CHARS) -> str:
    text = text or ""
    return text if len(text) <= limit else text[:limit] + "\n…(생략)"


def _require_pymapdl() -> None:
    if not MAPDL_AVAILABLE:
        raise AnsysError(
            "ansys-mapdl-core 패키지가 없어 ANSYS를 조종할 수 없습니다"
            f"({MAPDL_IMPORT_ERROR}). 사내 미러에서 pip install ansys-mapdl-core 후 "
            "다시 시도하세요. (서버 자체는 정상 동작 중입니다.)"
        )


def _session():
    """살아 있는 MAPDL 세션을 돌려준다. 없으면 안내와 함께 실패."""
    _require_pymapdl()
    if _mapdl is None:
        raise AnsysError(
            "MAPDL 세션이 없습니다. launch_ansys(로컬 기동) 또는 "
            "connect_ansys(기존 인스턴스 접속)를 먼저 호출하세요."
        )
    return _mapdl


def _cleanup_session() -> None:
    """서버가 종료될 때 MAPDL 자식 프로세스를 정리한다 — 고아(orphan) 방지.

    launch_mapdl로 띄운 MAPDL은 이 서버가 죽어도 자동으로 꺼지지 않는다. 서버가
    정상 종료되는 경로(Ctrl+C·클라이언트 연결 종료·mcp.run 반환)에서 이 함수가 세션을
    닫아 준다. 프로세스 강제 종료(kill -9 등)는 어떤 코드로도 잡을 수 없다 — 그때는
    --cleanup 이나 kill_local_ansys로 정리한다. 종료 중이라 락은 잡지 않고(교착 방지),
    어떤 예외든 삼킨다.
    """
    global _mapdl
    m, _mapdl = _mapdl, None
    if m is not None:
        try:
            m.exit()
        except Exception:  # noqa: BLE001 — 종료 경로라 무엇이든 무시
            pass


def _param(mapdl, name: str) -> float:
    """MAPDL 스칼라 파라미터를 float로 읽는다. (*GET 결과 회수용)"""
    try:
        return float(mapdl.parameters[name])
    except Exception as e:  # noqa: BLE001
        raise AnsysError(f"MAPDL 파라미터 '{name}'를 읽지 못했습니다: {e}") from e


def _area_spec(area: int) -> str:
    """면 지정 문자열 — 0이면 ALL(모든 면), 아니면 그 번호."""
    return "ALL" if int(area) == 0 else str(int(area))


# ════════════════════════════ 🟢 상태 조회 ════════════════════════════


@mcp.tool()
@ansys_tool
def ansys_status() -> str:
    """PyMAPDL 패키지·MAPDL 세션 상태를 확인합니다. (🟢 읽기)

    무엇이 안 될 때 가장 먼저 호출하세요.
    """
    lines = []
    if not MAPDL_AVAILABLE:
        lines.append(f"ansys-mapdl-core: 없음 ({MAPDL_IMPORT_ERROR})")
        lines.append("→ 사내 미러에서 pip install ansys-mapdl-core 후 서버를 재시작하세요.")
        return "\n".join(lines)
    lines.append("ansys-mapdl-core: 설치됨")
    if _mapdl is None:
        lines.append("MAPDL 세션: 없음 — launch_ansys 또는 connect_ansys로 시작하세요.")
        return "\n".join(lines)
    try:
        ver = _mapdl.version
        lines.append(f"MAPDL 세션: 연결됨 (버전 {ver})")
    except Exception as e:  # noqa: BLE001 — 죽은 세션
        lines.append(f"MAPDL 세션: 응답 없음 ({type(e).__name__}: {e}) — 재기동이 필요할 수 있습니다.")
        return "\n".join(lines)
    try:
        lines.append(f"작업 폴더: {_mapdl.directory}")
    except Exception:  # noqa: BLE001
        pass
    return "\n".join(lines)


# ════════════════════════════ 🟡 세션 관리 ════════════════════════════


@mcp.tool()
@ansys_tool
def launch_ansys(exec_file: str = "", run_location: str = "", jobname: str = "thermal") -> str:
    """로컬 PC의 ANSYS MAPDL을 새로 띄워 세션을 엽니다. (🟡)

    기동에 수 초~수십 초 걸리고 라이선스를 체크아웃합니다. 이미 세션이 있으면
    그대로 씁니다 (재기동하려면 shutdown_ansys 후 다시).

    Args:
        exec_file: ansys 실행 파일 경로. 비우면 설치 위치를 자동 탐색.
        run_location: 작업 폴더. 비우면 임시 폴더.
        jobname: MAPDL 잡 이름 (기본 thermal).
    """
    global _mapdl
    _require_pymapdl()
    if _mapdl is not None:
        return "이미 MAPDL 세션이 있습니다. 그대로 사용하세요 (재기동은 shutdown_ansys 후)."
    kwargs: dict = {"override": True, "loglevel": "ERROR", "jobname": jobname}
    if exec_file:
        kwargs["exec_file"] = exec_file
    if run_location:
        os.makedirs(run_location, exist_ok=True)
        kwargs["run_location"] = run_location
    start = time.time()
    _mapdl = pymapdl.launch_mapdl(**kwargs)
    return (
        f"MAPDL 기동 완료 ({time.time() - start:.1f}초, 버전 {_mapdl.version}). "
        "다음: create_block/create_cylinder로 형상을 만드세요."
    )


@mcp.tool()
@ansys_tool
def connect_ansys(ip: str = "127.0.0.1", port: int = 50052) -> str:
    """이미 떠 있는 MAPDL 인스턴스(gRPC)에 접속합니다. (🟡)

    다른 PC/서버에서 `ansys ... -grpc` 로 띄워 둔 인스턴스에 붙을 때 사용.

    Args:
        ip: MAPDL gRPC 서버 주소.
        port: gRPC 포트 (MAPDL 기본 50052).
    """
    global _mapdl
    _require_pymapdl()
    if _mapdl is not None:
        return "이미 MAPDL 세션이 있습니다. 바꾸려면 shutdown_ansys(로컬 기동일 때) 후 다시 접속하세요."
    # ⚠ 실기 검증 대상: 버전에 따라 connect_to_mapdl 유무가 다르다 — Mapdl 클래스로 폴백.
    connect = getattr(pymapdl, "connect_to_mapdl", None)
    if connect is not None:
        _mapdl = connect(ip=ip, port=int(port))
    else:
        _mapdl = pymapdl.Mapdl(ip=ip, port=int(port))
    return f"MAPDL 접속 완료 ({ip}:{port}, 버전 {_mapdl.version})."


@mcp.tool()
@ansys_tool
def open_gui(include_result: bool = True) -> str:
    """🟡 현재 모델을 ANSYS MAPDL GUI 창으로 띄워 직접 봅니다 (보기 전용).

    PyMAPDL이 현재 DB를 저장한 뒤 진짜 MAPDL GUI를 엽니다. capture_plot이 정지 PNG를
    주는 것과 달리, 3D로 회전·확대하며 살펴볼 수 있습니다.

    ⚠ 제약:
    - **블로킹** — GUI 창을 닫을 때까지 이 호출이 반환되지 않고, 그동안 다른 도구
      호출은 락에서 대기합니다. 잠깐 확인할 때만 쓰세요.
    - **로컬 데스크톱 전용** — MAPDL이 실제로 도는 그 PC의 로그인 세션 화면에만 뜹니다.
      원격(connect_ansys) 인스턴스나 창을 못 띄우는 세션이면 열리지 않습니다.
    - **보기 전용** — GUI에서 만진 변경이 세션으로 되돌아온다는 보장은 없습니다.

    Args:
        include_result: 결과 파일(.rst)도 함께 불러올지 (기본 True — solve 후 결과 확인용).
                        아직 solve 전이면 False로 두세요.
    """
    mapdl = _session()
    mapdl.open_gui(include_result=include_result)
    return "MAPDL GUI 창을 닫아 세션으로 복귀했습니다."


@mcp.tool()
@ansys_tool
def shutdown_ansys(confirm: bool = False) -> str:
    """🔴 MAPDL 세션(프로세스)을 종료합니다. (confirm=True 필요)

    저장하지 않은 모델은 사라집니다. confirm 없이 부르면 프리뷰만 돌려줍니다.
    """
    global _mapdl
    mapdl = _session()
    if not confirm:
        return (
            "⚠️ 승인 필요 — 아직 실행하지 않았습니다: MAPDL 세션 종료\n"
            "저장하지 않은 모델·결과는 사라집니다. 진행하려면 confirm=true로 다시 호출하세요."
        )
    try:
        mapdl.exit()
    finally:
        _mapdl = None
    return "MAPDL 세션을 종료했습니다."


@mcp.tool()
@ansys_tool
def kill_local_ansys(confirm: bool = False) -> str:
    """🔴 이 PC에 떠 있는 로컬 MAPDL(gRPC) 인스턴스를 모두 강제 종료합니다. (confirm=True 필요)

    서버가 비정상 종료돼 MAPDL이 고아(orphan)로 남았을 때 청소용입니다. shutdown_ansys가
    '현재 세션 하나'만 닫는 것과 달리, 이 도구는 이 PC에서 PyMAPDL이 띄운 **모든** 로컬
    gRPC 인스턴스를 닫습니다(현재 세션 포함). 대화형으로 직접 실행한 ANSYS GUI는 대상이
    아닙니다 — gRPC 인스턴스만 닫힙니다. 서버 밖에서 돌리려면 `--cleanup` CLI를 쓰세요.
    """
    global _mapdl
    _require_pymapdl()
    if not confirm:
        return (
            "⚠️ 승인 필요 — 아직 실행하지 않았습니다: 로컬 MAPDL(gRPC) 인스턴스 전체 종료\n"
            "이 PC의 PyMAPDL gRPC 인스턴스가 모두 닫힙니다(현재 세션 포함). "
            "진행하려면 confirm=true로 다시 호출하세요."
        )
    if _mapdl is not None:
        try:
            _mapdl.exit()
        except Exception:  # noqa: BLE001 — 이미 죽었을 수도
            pass
        _mapdl = None
    pymapdl.close_all_local_instances()
    return "이 PC의 로컬 MAPDL(gRPC) 인스턴스를 정리했습니다."


@mcp.tool()
@ansys_tool
def clear_model(confirm: bool = False) -> str:
    """🔴 현재 모델(형상·메시·하중·결과)을 전부 지웁니다. (confirm=True 필요)

    세션은 유지되고 빈 모델에서 다시 시작합니다. confirm 없이 부르면 프리뷰만.
    """
    mapdl = _session()
    if not confirm:
        return (
            "⚠️ 승인 필요 — 아직 실행하지 않았습니다: 모델 전체 삭제(/CLEAR)\n"
            "형상·메시·경계조건·결과가 모두 사라집니다. 진행하려면 confirm=true로 다시 호출하세요."
        )
    mapdl.clear()
    return "모델을 비웠습니다. create_block 등으로 새로 시작하세요."


# ════════════════════════ 🟡 형상·요소·재료·메시 ════════════════════════


@mcp.tool()
@ansys_tool
def create_block(x1: float, y1: float, z1: float, x0: float = 0, y0: float = 0, z0: float = 0) -> str:
    """직육면체 볼륨을 만듭니다. (🟡) 좌표 단위는 모델 단위(보통 m 권장 — SI 일관).

    Args:
        x1, y1, z1: 반대쪽 꼭짓점 좌표.
        x0, y0, z0: 시작 꼭짓점 좌표 (기본 원점).
    """
    mapdl = _session()
    mapdl.run("/PREP7")
    mapdl.run(f"BLOCK,{x0},{x1},{y0},{y1},{z0},{z1}")
    return (
        f"블록 생성: ({x0},{y0},{z0}) ~ ({x1},{y1},{z1}). "
        "면 번호는 list_areas로 확인하세요."
    )


@mcp.tool()
@ansys_tool
def create_cylinder(radius: float, z0: float = 0, z1: float = 1) -> str:
    """Z축 방향 원기둥 볼륨을 만듭니다. (🟡)

    Args:
        radius: 반지름.
        z0, z1: 밑면/윗면 Z 좌표.
    """
    if float(radius) <= 0:
        raise AnsysError("반지름(radius)은 0보다 커야 합니다.")
    mapdl = _session()
    mapdl.run("/PREP7")
    mapdl.run(f"CYL4,0,0,{radius},,,,{float(z1) - float(z0)}")
    if float(z0) != 0:
        mapdl.run(f"VGEN,,ALL,,,,,{z0},,,1")  # ⚠ 실기 검증 대상: Z 이동
    return f"원기둥 생성: R{radius}, Z {z0}~{z1}. 면 번호는 list_areas로 확인하세요."


@mcp.tool()
@ansys_tool
def set_thermal_element(kind: str = "solid70") -> str:
    """열해석 요소 타입을 지정합니다. (🟡)

    Args:
        kind: 'solid70'(3D 8절점, 기본) | 'solid90'(3D 20절점, 고정밀) |
              'plane55'(2D 4절점 — 2D 단면 해석용).
    """
    types = {"solid70": "SOLID70", "solid90": "SOLID90", "plane55": "PLANE55"}
    key = kind.strip().lower()
    if key not in types:
        raise AnsysError(f"kind는 {'/'.join(types)} 중 하나여야 합니다 (받은 값: {kind}).")
    mapdl = _session()
    mapdl.run("/PREP7")
    mapdl.run(f"ET,1,{types[key]}")
    return f"요소 타입 지정: {types[key]}"


@mcp.tool()
@ansys_tool
def define_thermal_material(conductivity: float, density: float = 0,
                            specific_heat: float = 0, material_id: int = 1) -> str:
    """열물성 재료를 정의합니다. (🟡)

    정상상태 해석은 열전도율만 있어도 되고, 과도(시간) 해석은 밀도·비열도 필요합니다.
    단위는 모델 단위와 일관되게 (SI: W/m·K, kg/m³, J/kg·K).

    Args:
        conductivity: 열전도율 KXX (W/m·K).
        density: 밀도 (kg/m³) — 과도 해석 시 필수.
        specific_heat: 비열 (J/kg·K) — 과도 해석 시 필수.
        material_id: 재료 번호 (기본 1).
    """
    if float(conductivity) <= 0:
        raise AnsysError("열전도율(conductivity)은 0보다 커야 합니다.")
    mapdl = _session()
    mapdl.run("/PREP7")
    mid = int(material_id)
    mapdl.run(f"MP,KXX,{mid},{conductivity}")
    parts = [f"KXX={conductivity}"]
    if float(density) > 0:
        mapdl.run(f"MP,DENS,{mid},{density}")
        parts.append(f"DENS={density}")
    if float(specific_heat) > 0:
        mapdl.run(f"MP,C,{mid},{specific_heat}")
        parts.append(f"C={specific_heat}")
    note = "" if (float(density) > 0 and float(specific_heat) > 0) else \
        " (과도 해석을 하려면 density·specific_heat도 지정 필요)"
    return f"재료 {mid} 정의: {', '.join(parts)}{note}"


@mcp.tool()
@ansys_tool
def mesh_model(element_size: float) -> str:
    """모델 전체를 메시합니다. (🟡) 볼륨이 있으면 VMESH, 없으면 AMESH(2D).

    Args:
        element_size: 요소 목표 크기(모델 단위). 작을수록 정밀·느림.
    """
    if float(element_size) <= 0:
        raise AnsysError("요소 크기(element_size)는 0보다 커야 합니다.")
    mapdl = _session()
    mapdl.run("/PREP7")
    mapdl.run(f"ESIZE,{element_size}")
    # ⚠ 실기 검증 대상: *GET 볼륨 개수로 3D/2D 분기.
    mapdl.run("*GET,mcp_nvol,VOLU,0,COUNT")
    nvol = _param(mapdl, "mcp_nvol")
    if nvol > 0:
        mapdl.run("VMESH,ALL")
    else:
        mapdl.run("AMESH,ALL")
    mapdl.run("*GET,mcp_nnode,NODE,0,COUNT")
    mapdl.run("*GET,mcp_nelem,ELEM,0,COUNT")
    return (
        f"메시 완료: 노드 {int(_param(mapdl, 'mcp_nnode'))}개, "
        f"요소 {int(_param(mapdl, 'mcp_nelem'))}개 (크기 {element_size})"
    )


@mcp.tool()
@ansys_tool
def mesh_info() -> str:
    """현재 모델 통계 — 볼륨/면/노드/요소 수. (🟢 읽기)"""
    mapdl = _session()
    out = []
    # ⚠ 실기 검증 대상: *GET COUNT 시퀀스.
    for label, ent in (("볼륨", "VOLU"), ("면", "AREA"), ("노드", "NODE"), ("요소", "ELEM")):
        mapdl.run(f"*GET,mcp_cnt,{ent},0,COUNT")
        out.append(f"{label} {int(_param(mapdl, 'mcp_cnt'))}개")
    return " / ".join(out)


@mcp.tool()
@ansys_tool
def list_areas() -> str:
    """면(Area) 목록을 번호와 함께 나열합니다. (🟢 읽기)

    apply_temperature/apply_convection 등에서 면 번호를 지정하기 전에 확인하는 용도.
    블록의 경우 보통 6개 면이 1~6번입니다. 위치 확인이 어려우면 capture_plot(kind='areas')로
    번호가 표시된 그림을 떠서 보세요.
    """
    mapdl = _session()
    out = mapdl.run("ALIST,ALL")  # ⚠ 실기 검증 대상: 표 형식 텍스트 반환
    return _truncate(str(out))


# ════════════════════════ 🟡 경계조건·하중 ════════════════════════


@mcp.tool()
@ansys_tool
def apply_temperature(value: float, area: int = 0) -> str:
    """면에 고정 온도(디리클레 경계)를 겁니다. (🟡)

    Args:
        value: 온도 값(모델 단위 — 섭씨든 켈빈이든 일관되게).
        area: 면 번호 (list_areas로 확인). 0이면 모든 면.
    """
    mapdl = _session()
    mapdl.run("/PREP7")
    mapdl.run(f"DA,{_area_spec(area)},TEMP,{value}")
    return f"고정 온도 적용: 면 {_area_spec(area)} = {value}"


@mcp.tool()
@ansys_tool
def apply_convection(film_coefficient: float, bulk_temperature: float, area: int = 0) -> str:
    """면에 대류 경계조건을 겁니다. (🟡)

    Args:
        film_coefficient: 대류 열전달계수 h (W/m²·K).
        bulk_temperature: 주변 유체 온도.
        area: 면 번호 (0이면 모든 면).
    """
    mapdl = _session()
    mapdl.run("/PREP7")
    mapdl.run(f"SFA,{_area_spec(area)},1,CONV,{film_coefficient},{bulk_temperature}")
    return f"대류 적용: 면 {_area_spec(area)}, h={film_coefficient}, T∞={bulk_temperature}"


@mcp.tool()
@ansys_tool
def apply_heat_flux(flux: float, area: int = 0) -> str:
    """면에 열유속(단위면적당 입열)을 겁니다. (🟡)

    Args:
        flux: 열유속 (W/m²). 양수 = 유입.
        area: 면 번호 (0이면 모든 면).
    """
    mapdl = _session()
    mapdl.run("/PREP7")
    mapdl.run(f"SFA,{_area_spec(area)},1,HFLUX,{flux}")
    return f"열유속 적용: 면 {_area_spec(area)} = {flux}"


@mcp.tool()
@ansys_tool
def apply_heat_generation(rate: float, volume: int = 0) -> str:
    """볼륨에 체적 발열(단위부피당 발열량)을 겁니다. (🟡)

    Args:
        rate: 발열률 (W/m³).
        volume: 볼륨 번호. 0이면 모든 볼륨.
    """
    mapdl = _session()
    mapdl.run("/PREP7")
    spec = "ALL" if int(volume) == 0 else str(int(volume))
    mapdl.run(f"BFV,{spec},HGEN,{rate}")
    return f"발열 적용: 볼륨 {spec} = {rate}"


@mcp.tool()
@ansys_tool
def set_initial_temperature(value: float) -> str:
    """전체 초기(균일) 온도를 지정합니다 — 과도 해석의 시작 온도. (🟡)

    Args:
        value: 초기 온도.
    """
    mapdl = _session()
    mapdl.run(f"TUNIF,{value}")
    return f"초기 균일 온도: {value}"


# ════════════════════════ 🟡 해석 실행 ════════════════════════


@mcp.tool()
@ansys_tool
def solve_steady() -> str:
    """정상상태 열해석을 실행합니다. (🟡 — 모델 크기에 따라 오래 걸릴 수 있음)"""
    mapdl = _session()
    mapdl.run("FINISH")
    mapdl.run("/SOLU")
    mapdl.run("ANTYPE,STATIC")
    start = time.time()
    mapdl.run("SOLVE")
    mapdl.run("FINISH")
    return (
        f"정상상태 해석 완료 ({time.time() - start:.1f}초). "
        "result_temperature로 결과를 확인하세요."
    )


@mcp.tool()
@ansys_tool
def solve_transient(end_time: float, time_step: float) -> str:
    """과도(시간 이력) 열해석을 실행합니다. (🟡 — 오래 걸릴 수 있음)

    시작 온도는 set_initial_temperature로 미리 지정하세요. 재료에 밀도·비열이
    정의돼 있어야 합니다 (define_thermal_material).

    Args:
        end_time: 해석 종료 시각(초).
        time_step: 초기 시간 증분(초). 자동 시간 스텝(AUTOTS)이 조정한다.
    """
    if float(end_time) <= 0 or float(time_step) <= 0:
        raise AnsysError("end_time과 time_step은 0보다 커야 합니다.")
    mapdl = _session()
    mapdl.run("FINISH")
    mapdl.run("/SOLU")
    mapdl.run("ANTYPE,TRANS")
    mapdl.run(f"TIME,{end_time}")
    mapdl.run(f"DELTIM,{time_step}")
    mapdl.run("AUTOTS,ON")
    mapdl.run("OUTRES,ALL,ALL")  # 모든 시간 스텝 결과 저장
    start = time.time()
    mapdl.run("SOLVE")
    mapdl.run("FINISH")
    return (
        f"과도 해석 완료: 0~{end_time}초 ({time.time() - start:.1f}초 소요). "
        "result_temperature(time=...)로 특정 시각 결과를 확인하세요."
    )


# ════════════════════════ 🟢 결과 조회 ════════════════════════


def _set_result(mapdl, at_time: float) -> str:
    """POST1 진입 + 결과 세트 선택. at_time<=0이면 마지막 세트."""
    mapdl.run("/POST1")
    if float(at_time) > 0:
        mapdl.run(f"SET,,,,,{at_time}")  # ⚠ 실기 검증 대상: 시각 기준 세트 선택
        return f"(t={at_time}s)"
    mapdl.run("SET,LAST")
    return "(마지막 결과 세트)"


@mcp.tool()
@ansys_tool
def result_temperature(time: float = 0) -> str:
    """온도 결과 요약 — 최고/최저 온도와 해당 노드. (🟢 읽기)

    Args:
        time: 과도 해석에서 조회할 시각(초). 0이면 마지막 결과.
    """
    mapdl = _session()
    label = _set_result(mapdl, time)
    # ⚠ 실기 검증 대상: NSORT 후 *GET SORT MAX/MIN.
    mapdl.run("NSORT,TEMP")
    mapdl.run("*GET,mcp_tmax,SORT,0,MAX")
    mapdl.run("*GET,mcp_tmin,SORT,0,MIN")
    tmax = _param(mapdl, "mcp_tmax")
    tmin = _param(mapdl, "mcp_tmin")
    return f"온도 {label}: 최고 {tmax:.4g} / 최저 {tmin:.4g}"


@mcp.tool()
@ansys_tool
def result_heat_flux(time: float = 0) -> str:
    """열유속 크기(TFSUM) 최대값. (🟢 읽기)

    Args:
        time: 과도 해석에서 조회할 시각(초). 0이면 마지막 결과.
    """
    mapdl = _session()
    label = _set_result(mapdl, time)
    # ⚠ 실기 검증 대상: TF(열유속) 벡터 합 정렬.
    mapdl.run("NSORT,TF,SUM")
    mapdl.run("*GET,mcp_fmax,SORT,0,MAX")
    return f"열유속 크기 최대 {label}: {_param(mapdl, 'mcp_fmax'):.4g}"


@mcp.tool()
@ansys_tool
def temperature_at_point(x: float, y: float, z: float = 0, time: float = 0) -> str:
    """지정 좌표에서 가장 가까운 노드의 온도. (🟢 읽기)

    Args:
        x, y, z: 조회할 좌표(모델 단위).
        time: 과도 해석에서 조회할 시각(초). 0이면 마지막 결과.
    """
    mapdl = _session()
    label = _set_result(mapdl, time)
    # ⚠ 실기 검증 대상: NODE() get-함수로 최근접 노드 → 노드 온도 조회.
    mapdl.run(f"mcp_nid=NODE({x},{y},{z})")
    mapdl.run("*GET,mcp_t,NODE,mcp_nid,TEMP")
    nid = int(_param(mapdl, "mcp_nid"))
    return f"({x}, {y}, {z}) 근처 노드 {nid}의 온도 {label}: {_param(mapdl, 'mcp_t'):.4g}"


@mcp.tool()
@ansys_tool
def get_parameter(name: str) -> str:
    """MAPDL 스칼라 파라미터 값을 조회합니다 (run_apdl로 만든 값 회수용). (🟢 읽기)

    Args:
        name: 파라미터 이름.
    """
    mapdl = _session()
    return f"{name} = {_param(mapdl, name)}"


@mcp.tool()
@ansys_tool
def capture_plot(path: str, kind: str = "temp") -> str:
    """결과/모델 그림을 PNG로 저장합니다. (🟡 — 새 파일 하나 생성, 덮어쓰기 거부)

    MAPDL 자체 렌더러(/SHOW,PNG)를 쓰므로 pyvista 같은 추가 패키지가 필요 없다
    (폐쇄망 대응).

    Args:
        path: 저장할 PNG 경로. 이미 있으면 실패한다(덮어쓰지 않음).
        kind: 'temp'(온도 컨투어) | 'flux'(열유속 크기) | 'mesh'(요소) | 'areas'(면 번호).
    """
    p = os.path.abspath(os.path.expanduser(path))
    if not p.lower().endswith(".png"):
        p += ".png"
    if os.path.exists(p):
        raise AnsysError(f"'{p}' 파일이 이미 있습니다. 다른 이름을 지정하세요 (덮어쓰지 않습니다).")
    plots = {"temp": "PLNSOL,TEMP", "flux": "PLNSOL,TF,SUM", "mesh": "EPLOT", "areas": "APLOT"}
    key = kind.strip().lower()
    if key not in plots:
        raise AnsysError(f"kind는 {'/'.join(plots)} 중 하나여야 합니다 (받은 값: {kind}).")
    mapdl = _session()
    if key in ("temp", "flux"):
        mapdl.run("/POST1")
        mapdl.run("SET,LAST")
    if key == "areas":
        mapdl.run("/PNUM,AREA,1")  # 면 번호 표시
    # ⚠ 실기 검증 대상: /SHOW,PNG 덤프 — 작업 폴더에 jobname###.png가 생긴다.
    workdir = str(mapdl.directory)
    before = set(f for f in os.listdir(workdir) if f.lower().endswith(".png"))
    mapdl.run("/SHOW,PNG")
    mapdl.run(plots[key])
    mapdl.run("/SHOW,CLOSE")
    new = [f for f in os.listdir(workdir) if f.lower().endswith(".png") and f not in before]
    if not new:
        raise AnsysError("PNG가 생성되지 않았습니다. 결과가 있는지(solve 후인지) 확인하세요.")
    newest = max((os.path.join(workdir, f) for f in new), key=os.path.getmtime)
    os.makedirs(os.path.dirname(p) or ".", exist_ok=True)
    shutil.copyfile(newest, p)
    return f"그림 저장: {p} ({kind})"


# ════════════════════════ 🟡 만능 탈출구 ════════════════════════


@mcp.tool()
@ansys_tool
def run_apdl(command: str) -> str:
    """APDL 명령 한 줄을 그대로 실행하고 출력을 돌려줍니다. (🟡 — 전용 도구가 없을 때만)

    /CLEAR·/EXIT처럼 파괴적인 명령은 여기서 막는다 — clear_model/shutdown_ansys
    (confirm 게이트)를 쓸 것.

    Args:
        command: APDL 명령 (예: 'NLIST,1,10' 또는 'mypar=3.14').
    """
    cmd = (command or "").strip()
    if not cmd:
        raise AnsysError("command가 비어 있습니다.")
    upper = cmd.upper().lstrip()
    for banned, alt in (("/CLE", "clear_model"), ("/EXI", "shutdown_ansys"), ("EXIT", "shutdown_ansys")):
        if upper.startswith(banned):
            raise AnsysError(f"파괴적 명령은 run_apdl로 실행할 수 없습니다 — {alt}(confirm 게이트)를 쓰세요.")
    mapdl = _session()
    out = mapdl.run(cmd)
    return _truncate(str(out) if out is not None else "(출력 없음)")


# ─────────────────────────────── CLI / 서버 기동 ───────────────────────────────


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="ANSYS MAPDL 열해석 MCP 서버 (PyMAPDL)")
    parser.add_argument(
        "--transport", choices=["stdio", "http", "sse"],
        default=os.getenv("ANSYS_MCP_TRANSPORT", "stdio"),
        help="stdio(기본): 로컬 클라이언트가 직접 실행. http/sse: n8n 등 네트워크 접속.",
    )
    parser.add_argument("--host", default=os.getenv("ANSYS_MCP_HOST", "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.getenv("ANSYS_MCP_PORT", "8091")))
    parser.add_argument(
        "--cleanup", action="store_true",
        help="이 PC의 고아 MAPDL(gRPC) 인스턴스를 모두 정리하고 종료 (서버는 띄우지 않음). "
             "서버가 비정상 종료돼 MAPDL이 안 꺼지고 남았을 때 사용.",
    )
    args = parser.parse_args()

    if args.cleanup:
        # 서버를 띄우지 않고, 떠 있는 로컬 MAPDL 인스턴스만 정리한다.
        if not MAPDL_AVAILABLE:
            print(f"[오류] ansys-mapdl-core 없음({MAPDL_IMPORT_ERROR}) — 정리할 수 없습니다.",
                  file=sys.stderr)
            sys.exit(1)
        pymapdl.close_all_local_instances()
        print("이 PC의 로컬 MAPDL(gRPC) 인스턴스를 정리했습니다.")
        sys.exit(0)

    if not MAPDL_AVAILABLE:
        print(f"[주의] ansys-mapdl-core 없음({MAPDL_IMPORT_ERROR}) — 도구가 안내만 반환합니다.",
              file=sys.stderr)

    # 서버가 종료되는 어떤 정상 경로(Ctrl+C·EOF·mcp.run 반환)에서도 MAPDL을 정리한다.
    # (강제 종료는 잡지 못한다 — 그때는 --cleanup으로.)
    atexit.register(_cleanup_session)
    try:
        if args.transport in ("http", "sse"):
            path = "/mcp/" if args.transport == "http" else "/sse/"
            print(f"ANSYS MCP 서버 시작 ({args.transport}) — http://{args.host}:{args.port}{path}",
                  file=sys.stderr)
            mcp.run(transport=args.transport, host=args.host, port=args.port)
        else:
            # stdio: stdout은 MCP 프로토콜 채널 — 로그는 stderr로.
            print("ANSYS MCP 서버 시작 (stdio)", file=sys.stderr)
            mcp.run(transport="stdio")
    finally:
        _cleanup_session()
