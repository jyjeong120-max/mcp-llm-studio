"""pdf_server.py

DRM(사내 보안프로그램)으로 암호화된 PDF의 **텍스트를 읽어오는** MCP 서버입니다.
읽기 전용(🟢) — 파일을 수정하지 않습니다.

왜 이렇게 만드는가:
    사내 보안프로그램이 파일을 잠그고 암호화한 뒤 **인증된 SW에만 실시간 복호화**를
    해준다. Python으로 파일을 직접 열면(open) 암호화된 바이트만 읽히고, cmd copy도
    마찬가지다(둘 다 화이트리스트 밖 프로세스). 따라서 DRM을 '우회'할 수는 없고,
    **DRM이 복호화를 허용하는 인증 앱을 통과(through)**해야 한다. office_server가
    Word COM으로 DRM Word 문서를 읽는 것과 같은 원리다.

    ⚠ notepad/type으로 읽어 재구성하는 접근은 불가능하다: (1) 일반 프로세스에는
    복호화를 안 해주고, (2) PDF는 바이너리라 텍스트로 읽으면 손상된다.

추출 백엔드 (순서대로 시도 = 우아한 저하 체인):
    1. direct       — 원본 앞부분이 '%PDF'면(=DRM 미적용) pypdf로 바로 추출.
                      암호화돼 있으면(매직 불일치) 조용히 다음 백엔드로 넘어간다.
    2. word_com     — Word를 백그라운드로 띄워 PDF를 열고(Word가 자동 변환)
                      본문 텍스트를 뽑는다. Word가 DRM 인증 앱이면 복호화된 내용이
                      읽힌다. 레이아웃/표는 뭉개지지만 텍스트 추출엔 충분.
    3. reader_print — (실험적) 인증 확정된 Acrobat Reader로 'Microsoft Print to PDF'
                      인쇄를 걸어 평문 PDF를 만들고 pypdf로 추출. 조용한 파일 출력이
                      환경마다 달라 신뢰도가 낮다 — 아래 ⚠ 참고.

    어느 백엔드도 성공하지 못하면 예외가 아니라 **안내 문자열**을 돌려준다.

폐쇄망 반입 체크리스트:
    - pypdf 가 사내 미러에 있는지 확인 (requirements.txt에 pypdf>=4.0). 없으면
      direct/reader_print 백엔드가 비활성화되고 word_com만 남는다 (우아한 저하).
    - Word COM 백엔드는 office_server와 같은 제약을 상속한다: **사용자 로그인 세션**에서
      실행, Windows + Office(Word) 설치 필요. pywin32(win32com)가 없으면 word_com도
      비활성.
    - DRM이 Word.exe(프로세스)에 PDF 복호화까지 허용하는지는 **실기 확인 대상**이다.
      확장자 단위로 스코프를 나누는 DRM이면 Word가 .docx는 되어도 .pdf는 막힐 수 있다.
      → 서버 없이 `python pdf_server.py --probe <PDF경로>` 로 각 백엔드를 진단할 것.

⚠ 실기 검증 대상:
    - Word가 DRM PDF를 실제로 복호화해 여는지 (--probe로 확인).
    - Word의 'PDF를 편집 가능한 문서로 변환' 확인창 처리: DisplayAlerts=0으로는 억제되지
      않아(개발 PC에서 재현됨) word_com은 (1) Open을 데몬 스레드에서 돌리고 (2) 워치독이
      그 대화상자를 자동 확인하며 (3) WORD_TIMEOUT(기본 90초) 초과 시 우리가 띄운 Word만
      강제 종료한다 — 그래서 막혀도 서버가 얼지 않는다. 실기에서 대화상자 자동 확인이
      실제로 통하는지는 --probe로 확인할 것(안 통하면 타임아웃으로 물러선다).
    - reader_print의 조용한 파일 출력. 'Microsoft Print to PDF'는 보통 저장 대화상자를
      띄우므로 자동 무인 인쇄가 안 될 수 있다(타임아웃으로 물러섬). 그 경우의 수동
      대안: Reader에서 파일 → 인쇄 → Microsoft Print to PDF로 평문 PDF를 만든 뒤,
      그 평문 PDF를 read_pdf_text에 넘기면 direct 백엔드로 바로 읽힌다.

사용:
    python pdf_server.py                     # MCP 서버 (stdio 기본)
    python pdf_server.py --transport http    # n8n 등 네트워크용, :8092
    python pdf_server.py --probe C:\a.pdf     # 각 백엔드 진단 (서버를 띄우지 않음)

llm_studio 장착 (코드 변경 불필요): 데이터 폴더의 mcp_servers.json에 등록하면
도구가 pdf__read_pdf_text 등으로 모델에 노출된다.
    {"mcpServers": {"pdf": {"command": "python",
                            "args": ["C:\\경로\\mcp_server\\pdf_server.py"]}}}
"""

from __future__ import annotations

import argparse
import glob
import os
import subprocess
import sys
import tempfile
from functools import wraps

from fastmcp import FastMCP

# pypdf — 없으면 direct/reader_print 백엔드가 비활성화된다 (word_com만 남음).
# import에서 죽지 않고, 도구 호출 시 안내로 저하한다.
try:
    from pypdf import PdfReader

    PYPDF_AVAILABLE = True
    PYPDF_IMPORT_ERROR = ""
except Exception as e:  # noqa: BLE001 — 하위 의존성 실패 포함
    PdfReader = None  # type: ignore[assignment]
    PYPDF_AVAILABLE = False
    PYPDF_IMPORT_ERROR = str(e)

# Word COM (pywin32) — 없으면 word_com 백엔드가 비활성화된다.
# office_server의 방식을 '차용'하되 import 결합은 만들지 않는다(자립형): office_server를
# import하면 그쪽 FastMCP 인스턴스와 전체 도구가 부작용으로 딸려오고, office import가
# 실패하면 여기 Word 백엔드까지 죽는다. ansys_server처럼 자체 스텁으로 독립시킨다.
try:
    import pythoncom  # pdf_tool 데코레이터의 스레드별 CoInitialize에 쓴다

    COM_AVAILABLE = True
    COM_IMPORT_ERROR = ""
except ImportError as e:  # Windows가 아니거나 pywin32 미설치
    COM_AVAILABLE = False
    COM_IMPORT_ERROR = str(e)

# Word COM 추출은 공용 헬퍼로 뺐다(text_server와 공유, 중복 0). 순수 헬퍼라 서버 부작용 없음.
from word_extract import extract_via_word  # noqa: E402

mcp = FastMCP(
    name="pdf",
    instructions=(
        "DRM으로 암호화된 사내 PDF의 텍스트를 읽어오는 읽기 전용 MCP 서버입니다. "
        "PDF 내용에 관한 질문을 받으면 read_pdf_text(path=...)로 본문을 읽어 근거로 "
        "답하세요. 문서가 너무 길면 pages 인자로 범위를 좁히세요(예: '1-5'). PDF가 "
        "안 읽히거나 어떤 추출 경로가 되는지 궁금하면 pdf_status(path=...)로 먼저 "
        "진단하세요. 이 서버는 파일을 수정하지 않습니다."
    ),
)

# 출력이 컨텍스트를 통째로 삼키지 않도록 하는 기본 상한. 도구 인자로 조정할 수 있다.
MAX_CHARS = 20000

# word_com 백엔드 하드 타임아웃(초). Word가 PDF 변환 대화상자에 멈추더라도 이 시간이
# 지나면 우리가 띄운 Word 프로세스를 강제 종료하고 물러선다 — MCP 서버가 영영 얼지
# 않도록 하는 안전장치다. 큰 PDF 변환은 오래 걸릴 수 있어 넉넉히 잡는다.
WORD_TIMEOUT = float(os.getenv("PDF_WORD_TIMEOUT", "90"))


class PdfError(Exception):
    """도구가 사용자에게 그대로 돌려줄 안내 메시지를 담은 예외."""


def pdf_tool(fn):
    """예외를 안내 문자열로 바꾸고, COM 스레드 초기화를 감싸는 도구 데코레이터.

    FastMCP는 동기 도구를 워커 스레드에서 실행한다. COM은 스레드마다 CoInitialize가
    필요하므로, pywin32가 있으면 매 호출마다 초기화/해제한다(word_com 백엔드용).
    pypdf 전용 경로에는 무해하다.
    """

    @wraps(fn)
    def wrapper(*args, **kwargs):
        if COM_AVAILABLE:
            pythoncom.CoInitialize()
        try:
            return fn(*args, **kwargs)
        except PdfError as e:
            return str(e)
        except Exception as e:  # noqa: BLE001 — 도구는 항상 문자열을 돌려준다
            return f"작업에 실패했습니다: {type(e).__name__}: {e}"
        finally:
            if COM_AVAILABLE:
                pythoncom.CoUninitialize()

    return wrapper


# ─────────────────────────────── 공용 헬퍼 ───────────────────────────────


def _abspath(path: str) -> str:
    """경로를 정규화하고 존재를 확인한다."""
    if not path:
        raise PdfError("path 인자에 PDF 파일 경로를 지정하세요.")
    p = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(p):
        raise PdfError(f"'{p}' 경로에 파일이 없습니다. 경로를 확인하세요.")
    if os.path.isdir(p):
        raise PdfError(f"'{p}'는 폴더입니다. PDF 파일 경로를 지정하세요.")
    return p


def _resolve_pages(spec: str, n_pages: int) -> list[int]:
    """'1-5', '3', '1,3,5-7' 같은 1-기반 페이지 지정을 0-기반 인덱스 리스트로 바꾼다.

    빈 문자열이면 전체 페이지. 범위를 벗어난 번호는 조용히 버린다.
    """
    if not spec or not spec.strip():
        return list(range(n_pages))
    idx: list[int] = []
    for part in spec.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            a, _, b = part.partition("-")
            try:
                lo, hi = int(a), int(b)
            except ValueError:
                raise PdfError(f"페이지 지정 '{part}'을(를) 이해하지 못했습니다. 예: '1-5'")
            for one in range(lo, hi + 1):
                if 1 <= one <= n_pages:
                    idx.append(one - 1)
        else:
            try:
                one = int(part)
            except ValueError:
                raise PdfError(f"페이지 지정 '{part}'을(를) 이해하지 못했습니다. 예: '3'")
            if 1 <= one <= n_pages:
                idx.append(one - 1)
    # 중복 제거하되 요청 순서 유지
    seen: set[int] = set()
    out: list[int] = []
    for i in idx:
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out


def _truncate(text: str, limit: int) -> str:
    text = text or ""
    if limit <= 0 or len(text) <= limit:
        return text
    return text[:limit] + f"\n…(생략 — 전체 {len(text)}자 중 앞 {limit}자. pages로 범위를 좁히세요)"


# ─────────────────────────────── 추출 백엔드 ───────────────────────────────
#
# 각 백엔드는 성공 시 텍스트(str)를, "이 백엔드로는 못 읽음(다음으로)"이면 None을
# 돌려준다. 치명적이지 않은 실패 사유는 reasons 리스트에 담아 안내에 활용한다.


def _extract_direct(path: str, pages: str, reasons: list[str]) -> str | None:
    """DRM이 안 걸린 평문 PDF를 pypdf로 바로 읽는다. 암호화면 None(다음 백엔드)."""
    if not PYPDF_AVAILABLE:
        reasons.append(f"direct: pypdf 없음({PYPDF_IMPORT_ERROR})")
        return None
    # 앞부분에 '%PDF' 서명이 없으면 DRM 암호화 등으로 봉인된 상태 — pypdf가 못 읽는다.
    try:
        with open(path, "rb") as f:
            head = f.read(1024)
    except OSError as e:
        reasons.append(f"direct: 파일 읽기 실패({e})")
        return None
    if b"%PDF" not in head[:8]:
        reasons.append("direct: '%PDF' 서명 없음 — DRM 암호화로 보임(인증 앱 경로로 시도)")
        return None
    try:
        reader = PdfReader(path)
        if reader.is_encrypted:
            # 표준 PDF 암호(DRM과 별개) — 빈 암호로 열리는 경우만 처리한다.
            try:
                reader.decrypt("")
            except Exception:  # noqa: BLE001
                reasons.append("direct: 표준 암호로 보호된 PDF(열기 암호 필요)")
                return None
        n = len(reader.pages)
        want = _resolve_pages(pages, n)
        parts = [reader.pages[i].extract_text() or "" for i in want]
        text = "\n\n".join(p.strip() for p in parts if p.strip())
        if not text.strip():
            reasons.append("direct: 추출된 텍스트 없음(스캔 이미지 PDF일 수 있음)")
            return None
        return text
    except Exception as e:  # noqa: BLE001
        reasons.append(f"direct: pypdf 추출 실패({type(e).__name__}: {e})")
        return None


def _find_acrobat() -> str | None:
    """설치된 Acrobat/Reader 실행 파일 경로를 찾는다. 없으면 None."""
    candidates: list[str] = []
    for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        if not base:
            continue
        candidates += glob.glob(os.path.join(base, "Adobe", "*", "Reader", "AcroRd32.exe"))
        candidates += glob.glob(os.path.join(base, "Adobe", "*", "Acrobat", "Acrobat.exe"))
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def _extract_reader_print(path: str, pages: str, reasons: list[str]) -> str | None:
    """(실험적) Acrobat Reader로 'Microsoft Print to PDF' 인쇄 → 평문 PDF → pypdf.

    ⚠ 'Microsoft Print to PDF'는 보통 저장 대화상자를 띄워 무인 출력이 안 될 수 있다.
    그 경우 타임아웃으로 물러선다(멈추지 않음). 신뢰할 수 있는 대안은 docstring 참고.
    """
    if not PYPDF_AVAILABLE:
        reasons.append(f"reader_print: pypdf 없음({PYPDF_IMPORT_ERROR})")
        return None
    exe = _find_acrobat()
    if not exe:
        reasons.append("reader_print: Acrobat/Reader 실행 파일을 찾지 못함")
        return None
    tmp_dir = tempfile.mkdtemp(prefix="pdf_print_")
    out = os.path.join(tmp_dir, "out.pdf")
    proc = None
    try:
        # /t <파일> <프린터> — 지정 프린터로 인쇄 후 종료. 출력 파일명을 못 지정하므로
        # 조용한 파일 출력은 환경(프린터 포트 설정)에 의존한다.
        proc = subprocess.Popen([exe, "/t", path, "Microsoft Print to PDF"])
        try:
            proc.wait(timeout=60)
        except subprocess.TimeoutExpired:
            reasons.append("reader_print: 인쇄가 60초 내 끝나지 않음(저장 대화상자 가능) — 중단")
            proc.kill()
            return None
        if os.path.exists(out) and os.path.getsize(out) > 0:
            return _extract_direct(out, pages, reasons)
        reasons.append("reader_print: 인쇄 결과 파일이 생성되지 않음(무인 PDF 프린터 설정 필요)")
        return None
    except Exception as e:  # noqa: BLE001
        reasons.append(f"reader_print: 인쇄 실패({type(e).__name__}: {e})")
        return None
    finally:
        if proc is not None and proc.poll() is None:
            try:
                proc.kill()
            except Exception:
                pass
        # 임시 폴더 정리 (기존 파일은 건드리지 않음)
        try:
            if os.path.exists(out):
                os.remove(out)
            os.rmdir(tmp_dir)
        except Exception:
            pass


def _extract(path: str, pages: str) -> tuple[str | None, str, list[str]]:
    """백엔드를 순서대로 시도해 (텍스트, 사용백엔드, 실패사유들)을 돌려준다."""
    reasons: list[str] = []
    for name, fn in (
        ("direct", lambda: _extract_direct(path, pages, reasons)),
        ("word_com", lambda: extract_via_word(path, reasons, WORD_TIMEOUT)),
        ("reader_print", lambda: _extract_reader_print(path, pages, reasons)),
    ):
        text = fn()
        if text:
            return text, name, reasons
    return None, "", reasons


def _no_backend_message(path: str, reasons: list[str]) -> str:
    """모든 백엔드가 실패했을 때의 안내 문자열."""
    detail = "\n".join(f"  - {r}" for r in reasons) if reasons else "  - (사유 없음)"
    return (
        f"'{os.path.basename(path)}'에서 텍스트를 추출하지 못했습니다. 시도한 경로:\n"
        f"{detail}\n"
        "DRM 암호화 PDF는 인증 앱(예: Word COM)이 복호화를 허용해야 읽힙니다. "
        "`python pdf_server.py --probe <경로>`로 어떤 경로가 되는지 진단하거나, "
        "Acrobat Reader에서 'Microsoft Print to PDF'로 평문 PDF를 만든 뒤 그 파일을 "
        "다시 읽어 보세요. 그래도 안 되면 회사 IT에 PDF 추출 권한을 문의하세요."
    )


# ─────────────────────────────── MCP 도구 (🟢 읽기 전용) ───────────────────────────────


@mcp.tool()
@pdf_tool
def read_pdf_text(path: str, pages: str = "", max_chars: int = MAX_CHARS) -> str:
    """PDF에서 본문 텍스트를 추출합니다. (🟢 읽기 전용)

    DRM으로 암호화된 PDF도 인증 앱 경로(Word COM 등)로 복호화되면 읽습니다.

    Args:
        path: PDF 파일의 전체 경로.
        pages: 읽을 페이지 범위. 빈 값이면 전체. 예: '1-5', '3', '1,3,5-7'.
               (Word COM 백엔드로 읽힌 DRM PDF는 페이지 경계가 어긋나 전체가 반환될 수 있습니다.)
        max_chars: 반환 최대 글자 수(기본 20000). 초과하면 잘라내고 안내를 붙입니다.
    """
    p = _abspath(path)
    text, backend, reasons = _extract(p, pages)
    if not text:
        return _no_backend_message(p, reasons)
    note = ""
    if backend == "word_com" and pages.strip():
        note = "\n\n(참고: Word 변환 경로로 읽어 pages 지정이 정확히 적용되지 않아 전체 본문을 반환했습니다.)"
    return _truncate(text, max_chars) + note


@mcp.tool()
@pdf_tool
def read_pdf_metadata(path: str) -> str:
    """PDF의 페이지 수·제목·작성자 등 메타데이터를 조회합니다. (🟢 읽기 전용)

    평문 PDF는 pypdf로, DRM PDF는 Word COM으로 열어 확인합니다.
    """
    p = _abspath(path)
    name = os.path.basename(p)
    # 평문이면 pypdf 메타데이터가 가장 정확하다.
    if PYPDF_AVAILABLE:
        try:
            with open(p, "rb") as f:
                head = f.read(8)
            if b"%PDF" in head:
                reader = PdfReader(p)
                if reader.is_encrypted:
                    try:
                        reader.decrypt("")
                    except Exception:  # noqa: BLE001
                        return f"'{name}': 표준 암호로 보호된 PDF(열기 암호 필요)."
                md = reader.metadata or {}
                lines = [f"파일: {name}", f"페이지 수: {len(reader.pages)}"]
                for key, label in (("/Title", "제목"), ("/Author", "작성자"),
                                   ("/Subject", "주제"), ("/Creator", "생성 프로그램"),
                                   ("/Producer", "생산 도구")):
                    val = md.get(key)
                    if val:
                        lines.append(f"{label}: {val}")
                return "\n".join(lines)
        except Exception:  # noqa: BLE001
            pass  # DRM 등 — 아래 Word 경로로 폴백
    # DRM 등 pypdf가 못 읽는 경우: Word로 열어 페이지 수/문서 속성을 확인한다.
    reasons: list[str] = []
    text = extract_via_word(p, reasons, WORD_TIMEOUT)
    if text:
        # Word 변환본에서 대략적인 정보만 제공 (정확한 원본 페이지 수는 알기 어렵다).
        approx = len(text)
        return (f"파일: {name}\n"
                f"(DRM PDF — Word 변환 경로로 확인) 추출 본문 길이: 약 {approx}자.\n"
                "정확한 페이지 수/제목 등 메타데이터는 Word 변환본에서 신뢰하기 어렵습니다.")
    return _no_backend_message(p, reasons)


@mcp.tool()
@pdf_tool
def pdf_status(path: str = "") -> str:
    """추출 백엔드의 가용 상태를 보고합니다. path를 주면 그 파일로 실제 진단합니다. (🟢 읽기 전용)

    PDF가 안 읽히거나 어떤 경로가 되는지 궁금할 때 먼저 호출하세요.
    """
    lines = ["PDF 추출 백엔드 상태:"]
    lines.append(f"  - direct(pypdf): {'가용' if PYPDF_AVAILABLE else '비활성 — ' + PYPDF_IMPORT_ERROR}")
    lines.append(f"  - word_com(pywin32): {'가용' if COM_AVAILABLE else '비활성 — ' + COM_IMPORT_ERROR}")
    exe = _find_acrobat()
    lines.append(f"  - reader_print(Acrobat): {exe if exe else '실행 파일 못 찾음(실험적 백엔드)'}")

    if not path:
        lines.append("\npath 인자에 PDF 경로를 주면 각 백엔드로 실제 추출을 진단합니다.")
        return "\n".join(lines)

    p = _abspath(path)
    lines.append(f"\n진단 대상: {p}")
    text, backend, reasons = _extract(p, "")
    if text:
        preview = text[:200].replace("\n", " ")
        lines.append(f"✅ 성공 — 백엔드 '{backend}'. 미리보기: {preview}…")
    else:
        lines.append("❌ 모든 백엔드 실패.")
    if reasons:
        lines.append("시도 로그:")
        lines += [f"  · {r}" for r in reasons]
    return "\n".join(lines)


# ─────────────────────────────── CLI / 서버 기동 ───────────────────────────────


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="DRM PDF 텍스트 추출 MCP 서버 (읽기 전용)"
    )
    parser.add_argument(
        "--probe", metavar="PDF",
        help="각 추출 백엔드를 지정 PDF로 진단하고 종료 (MCP 서버를 띄우지 않음).",
    )
    parser.add_argument(
        "--transport", choices=["stdio", "http", "sse"],
        default=os.getenv("PDF_MCP_TRANSPORT", "stdio"),
        help="stdio(기본): 로컬 클라이언트가 직접 실행. http/sse: n8n 등 네트워크 접속.",
    )
    parser.add_argument("--host", default=os.getenv("PDF_MCP_HOST", "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.getenv("PDF_MCP_PORT", "8092")))
    args = parser.parse_args()

    if args.probe:
        # CLI 진단 모드: pdf_status를 직접 호출한다(pdf_tool이 COM 초기화/예외를 처리).
        print(pdf_status(args.probe))
        sys.exit(0)

    if not PYPDF_AVAILABLE:
        print(f"[주의] pypdf 없음({PYPDF_IMPORT_ERROR}) — direct/reader_print 백엔드 비활성, "
              "word_com만 사용.", file=sys.stderr)
    if not COM_AVAILABLE:
        print(f"[주의] pywin32 없음({COM_IMPORT_ERROR}) — word_com 백엔드 비활성 "
              "(DRM PDF는 못 읽습니다).", file=sys.stderr)

    if args.transport in ("http", "sse"):
        path = "/mcp/" if args.transport == "http" else "/sse/"
        print(f"PDF MCP 서버 시작 ({args.transport}) — http://{args.host}:{args.port}{path}",
              file=sys.stderr)
        mcp.run(transport=args.transport, host=args.host, port=args.port)
    else:
        # stdio: stdout은 MCP 프로토콜 채널 — 로그는 stderr로.
        print("PDF MCP 서버 시작 (stdio)", file=sys.stderr)
        mcp.run(transport="stdio")
