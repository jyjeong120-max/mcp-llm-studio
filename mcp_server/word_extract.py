"""word_extract.py

Word COM으로 **임의의 파일을 열어 본문 텍스트를 뽑는** 공용 헬퍼. pdf_server(PDF를 Word가
자동 변환)와 text_server(DRM 텍스트/코드 파일)가 함께 쓴다.

왜 헬퍼로 두는가:
    FastMCP 인스턴스가 없는 **순수 헬퍼**라 import해도 서버 부작용(다른 서버의 mcp 객체·
    도구가 딸려옴)이 없다. 그래서 서버끼리 안전하게 공유할 수 있다 — pdf_server가
    office_server를 통째로 import하지 않으려던 이유(자립)를 헬퍼로 지키면서 중복을 없앤다.

DRM 원리:
    사내 보안프로그램은 인증된 SW(Word.exe)에만 실시간 복호화를 해준다. Python open()·cmd
    copy는 화이트리스트 밖이라 암호화 바이트만 읽힌다. 그래서 파일을 Word로 열어
    doc.Content.Text를 읽어야 평문이 나온다(office_server가 DRM Word 문서를 읽는 원리와 동일).

hang-safe:
    Word의 '변환/인코딩 확인' 대화상자는 DisplayAlerts=0으로 억제되지 않아(개발 PC 재현)
    백그라운드 인스턴스를 무한 대기시킬 수 있다. 그래서 (1) Open을 데몬 스레드에서 돌리고
    (2) 워치독이 그 대화상자를 자동 확인하며 (3) 타임아웃 초과 시 **우리가 띄운 Word PID만**
    골라 강제 종료한다 — 막혀도 MCP 서버가 영영 얼지 않는다. 사용자 Word는 건드리지 않는다.

⚠ 실기 검증 대상: DRM이 Word.exe에 해당 확장자(.pdf/.txt 등) 복호화를 허용하는지,
    대화상자 자동 확인이 실기에서 실제로 통하는지 (각 서버의 --probe로 확인).
"""

from __future__ import annotations

import subprocess
import threading
import time

# Word COM (pywin32) — 없으면 이 헬퍼 전체가 비활성(호출 시 None + 사유). Windows가 아니거나
# pywin32 미설치면 여기서 죽지 않고 COM_AVAILABLE=False로 우아하게 저하한다.
try:
    import pythoncom
    import win32com.client

    COM_AVAILABLE = True
    COM_IMPORT_ERROR = ""
except ImportError as e:  # Windows가 아니거나 pywin32 미설치
    COM_AVAILABLE = False
    COM_IMPORT_ERROR = str(e)

# 기본 하드 타임아웃(초). 호출자가 필요에 맞게 넘긴다(큰 PDF 변환은 오래 걸림).
DEFAULT_WORD_TIMEOUT = 90.0


def clean_word_text(raw: str) -> str:
    """Word Content.Text의 제어문자를 사람이 읽을 수 있게 정리한다.

    Word는 단락을 \\r, 페이지 나눔을 \\x0c, 셀/특수 끝표시를 \\x07 등으로 표현한다.
    """
    if not raw:
        return ""
    text = raw.replace("\r\n", "\n").replace("\r", "\n").replace("\x0c", "\n\n")
    # 표 셀 끝표시(\x07)와 기타 제어문자 제거 (줄바꿈/탭은 남긴다)
    text = "".join(ch for ch in text if ch >= " " or ch in "\n\t")
    return text.strip()


def winword_pids() -> set[int]:
    """지금 실행 중인 WINWORD.EXE 프로세스 ID 집합. (psutil 없이 ctypes/psapi)

    우리가 백그라운드로 띄운 Word만 골라 종료하기 위해, 생성 전후의 차집합을 쓴다.
    """
    try:
        import ctypes
        from ctypes import wintypes
    except Exception:  # noqa: BLE001 — Windows가 아니면 여기 올 일이 없다
        return set()
    try:
        psapi = ctypes.WinDLL("Psapi.dll")
        kernel = ctypes.WinDLL("kernel32.dll")
        # 64비트에서 핸들이 잘리지 않도록 반환/인자 타입을 명시한다.
        kernel.OpenProcess.restype = wintypes.HANDLE
        kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        psapi.GetModuleBaseNameW.argtypes = [
            wintypes.HANDLE, wintypes.HMODULE, wintypes.LPWSTR, wintypes.DWORD,
        ]
        arr = (wintypes.DWORD * 4096)()
        needed = wintypes.DWORD()
        if not psapi.EnumProcesses(ctypes.byref(arr), ctypes.sizeof(arr), ctypes.byref(needed)):
            return set()
        count = needed.value // ctypes.sizeof(wintypes.DWORD)
        pids: set[int] = set()
        # GetModuleBaseNameW는 PROCESS_VM_READ가 필요하다(LIMITED 핸들로는 실패).
        PROCESS_QUERY_INFORMATION = 0x0400
        PROCESS_VM_READ = 0x0010
        for i in range(count):
            pid = int(arr[i])
            if pid == 0:
                continue
            h = kernel.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid)
            if not h:
                continue
            try:
                buf = ctypes.create_unicode_buffer(260)
                if psapi.GetModuleBaseNameW(h, None, buf, 260) and buf.value.upper() == "WINWORD.EXE":
                    pids.add(pid)
            finally:
                kernel.CloseHandle(h)
        return pids
    except Exception:  # noqa: BLE001
        return set()


def kill_pids(pids: set[int]) -> None:
    """지정 PID들을 강제 종료한다(우리가 띄운 Word 정리용). 사용자 문서는 대상이 아니다."""
    for pid in pids:
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/F", "/T"],
                           capture_output=True, timeout=10)
        except Exception:  # noqa: BLE001
            pass


def dismiss_word_dialogs(pids: set[int]) -> None:
    """지정 Word PID가 띄운 모달 대화상자(예: 'PDF를 편집 가능한 문서로 변환', 인코딩 확인)를
    자동으로 확인(기본 버튼/Enter)해 Open이 진행되도록 한다. best-effort — 실패해도 무해."""
    if not pids:
        return
    try:
        import win32api
        import win32con
        import win32gui
        import win32process
    except Exception:  # noqa: BLE001 — pywin32 UI 모듈이 없으면 그냥 건너뛴다(타임아웃이 처리)
        return

    def _cb(hwnd, _):
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
            if wpid in pids and win32gui.GetClassName(hwnd) in ("#32770", "NUIDialog"):
                try:
                    win32gui.SetForegroundWindow(hwnd)
                except Exception:  # noqa: BLE001
                    pass
                # 기본 버튼(OK/확인) 실행: WM_COMMAND IDOK 와 Enter 키 둘 다 시도
                win32gui.PostMessage(hwnd, win32con.WM_COMMAND, win32con.IDOK, 0)
                win32api.keybd_event(0x0D, 0, 0, 0)
                win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception:  # noqa: BLE001
            pass
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:  # noqa: BLE001
        pass


def extract_via_word(path: str, reasons: list[str],
                     timeout: float = DEFAULT_WORD_TIMEOUT) -> str | None:
    """Word를 백그라운드로 띄워 파일을 열고 본문 텍스트를 뽑는다. 성공 시 텍스트, 못 읽으면 None.

    Word가 DRM 인증 앱이면 복호화된 내용이 읽힌다. PDF는 Word가 자동 변환하고, 텍스트/코드
    파일은 그대로 열린다. 못 읽거나 비어 있으면 None을 돌려주고 사유를 reasons에 담는다.

    ⚠ 변환/인코딩 확인창이 DisplayAlerts로 억제되지 않으므로: (1) Open을 데몬 스레드에서
    돌리고 (2) 워치독이 대화상자를 자동 확인하며 (3) 타임아웃 시 우리가 띄운 Word만 종료한다.
    """
    if not COM_AVAILABLE:
        reasons.append(f"word_com: pywin32 없음({COM_IMPORT_ERROR})")
        return None

    before = winword_pids()
    result: dict[str, str] = {}

    def _opener():
        # 데몬 스레드 — 자체 아파트먼트에서 COM을 초기화하고 app도 이 스레드에서 만든다
        # (COM 객체를 스레드 간에 넘기지 않는다). 메인 스레드는 PID/창만 다룬다.
        pythoncom.CoInitialize()
        app = None
        doc = None
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            # FileName, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
            doc = app.Documents.Open(path, False, True, False)
            result["text"] = clean_word_text(doc.Content.Text or "")
        except Exception as e:  # noqa: BLE001 — pythoncom.com_error 포함
            result["err"] = f"{type(e).__name__}: {e}"
        finally:
            try:
                if doc is not None:
                    doc.Close(SaveChanges=0)
            except Exception:
                pass
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    th = threading.Thread(target=_opener, daemon=True)
    th.start()

    # 대기하면서 우리가 띄운 Word의 대화상자를 계속 닫아 준다.
    ours: set[int] = set()
    deadline = time.time() + timeout
    while th.is_alive() and time.time() < deadline:
        if not ours:
            ours = winword_pids() - before
        dismiss_word_dialogs(ours)
        time.sleep(0.3)

    if th.is_alive():
        # 타임아웃 — 우리가 띄운 Word만 강제 종료(스레드는 데몬이라 함께 정리된다).
        kill_pids(winword_pids() - before)
        reasons.append(
            f"word_com: {int(timeout)}초 내 완료되지 않아 중단했습니다"
            "(Word의 변환/인코딩 대화상자에 막혔을 수 있음 — 실기에서 --probe로 확인)."
        )
        return None

    if "text" in result:
        text = result["text"]
        if text.strip():
            return text
        reasons.append("word_com: 열었으나 텍스트가 비어 있음(빈 파일/이미지일 수 있음)")
        return None

    # 오류로 끝난 경로 — 혹시 남은 우리 Word가 있으면 정리.
    kill_pids(winword_pids() - before)
    reasons.append(f"word_com: Word로 열지 못함({result.get('err', '알 수 없는 오류')})")
    return None
