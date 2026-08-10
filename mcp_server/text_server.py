"""text_server.py

DRM(사내 보안프로그램)으로 암호화된 **텍스트/코드 파일**의 내용을 읽어오는 MCP 서버입니다.
읽기 전용(🟢) — 파일을 수정하지 않습니다.

왜 이렇게 만드는가:
    사내 보안프로그램은 파일을 암호화한 뒤 **인증된 SW에만 실시간 복호화**를 해준다. 그래서
    .txt·.csv·.json·소스코드처럼 원래는 Python open()으로 바로 읽히는 파일도, DRM이 걸리면
    open()엔 암호화 바이트만 잡힌다(화이트리스트 밖 프로세스). notepad·cmd type도 마찬가지.
    → DRM을 '우회'하는 게 아니라, 복호화를 허용하는 **인증 앱(Word.exe)을 통과(through)**해
    읽는다. office_server가 DRM Word 문서를, pdf_server가 DRM PDF를 Word COM으로 읽는 것과
    똑같은 원리다. Word는 .txt/.csv/코드도 평문 문서로 열 수 있다.

추출 백엔드 (순서대로 = 우아한 저하 체인):
    1. direct   — Python으로 바이트를 읽어 utf-8/utf-8-sig/cp949로 '엄격' 디코드해 본다.
                  깨끗이 디코드되고 내용이 텍스트답게 보이면(제어문자 비율이 낮으면) 그대로
                  반환한다 = DRM이 안 걸린(또는 비활성인) 평문. 디코드가 깨지거나 암호문처럼
                  보이면 조용히 다음 백엔드로 넘어간다.
    2. word_com — Word를 백그라운드로 띄워 파일을 열고 본문 텍스트를 뽑는다(word_extract 공용
                  헬퍼). Word가 DRM 인증 앱이면 복호화된 내용이 읽힌다. hang-safe(변환/인코딩
                  대화상자를 워치독이 자동 확인, 타임아웃 시 우리 Word만 종료).

    어느 백엔드도 성공하지 못하면 예외가 아니라 **안내 문자열**을 돌려준다.

폐쇄망 반입 체크리스트:
    - 추가 pip 의존성 없음(표준 라이브러리 + pywin32). pywin32가 없으면 word_com만 비활성.
    - word_com은 office/pdf 서버와 같은 제약: **사용자 로그인 세션**에서 실행, Windows + Word.
    - DRM이 Word.exe에 해당 확장자(.txt 등) 복호화를 허용하는지는 **실기 확인 대상**이다.
      → 서버 없이 `python text_server.py --probe <파일경로>` 로 각 백엔드를 진단할 것.

사용:
    python text_server.py                     # MCP 서버 (stdio 기본)
    python text_server.py --transport http    # n8n 등 네트워크용, :8093
    python text_server.py --probe C:\a.txt     # 각 추출 백엔드를 진단 (서버 안 띄움)

llm_studio 장착: 앱에 **내장(builtin)**으로 들어가 설정 없이 text__read_text_file로 노출된다
    (llm_studio/server/builtin_servers.py). 독립 MCP 서버로도 그대로 쓸 수 있다(이중 용도).
"""

from __future__ import annotations

import argparse
import os
import sys
from functools import wraps

from fastmcp import FastMCP

# Word COM 추출은 공용 헬퍼(word_extract)와 공유한다 — pdf_server와 중복 0. 순수 헬퍼라
# import해도 서버 부작용이 없다. pywin32가 없으면 COM_AVAILABLE=False로 우아하게 저하.
from word_extract import COM_AVAILABLE, COM_IMPORT_ERROR, extract_via_word

mcp = FastMCP(
    name="text",
    instructions=(
        "DRM으로 암호화된 사내 텍스트/코드 파일(.txt, .md, .csv, .json, 소스코드 등)의 내용을 "
        "읽어오는 읽기 전용 MCP 서버입니다. 파일 내용에 관한 질문을 받으면 "
        "read_text_file(path=...)로 본문을 읽어 근거로 답하세요. 파일이 안 읽히거나 어떤 "
        "추출 경로가 되는지 궁금하면 text_status(path=...)로 먼저 진단하세요. 이 서버는 "
        "파일을 수정하지 않습니다."
    ),
)

# 출력이 컨텍스트를 통째로 삼키지 않도록 하는 기본 상한. 도구 인자로 조정한다.
MAX_CHARS = 20000

# word_com 백엔드 하드 타임아웃(초). 텍스트는 변환이 없어 대개 빠르지만, 인코딩 확인
# 대화상자에 막힐 수 있어 넉넉히 잡되 PDF(90초)보다는 짧게 둔다.
WORD_TIMEOUT = float(os.getenv("TEXT_WORD_TIMEOUT", "60"))


class TextError(Exception):
    """도구가 사용자에게 그대로 돌려줄 안내 메시지를 담은 예외."""


def text_tool(fn):
    """예외를 안내 문자열로 바꾸는 도구 데코레이터.

    (direct 백엔드는 COM을 안 쓰고, word_com은 word_extract가 자체 스레드에서 COM을
    초기화하므로, pdf_server와 달리 도구 스레드에서 CoInitialize를 하지 않아도 된다.)
    """

    @wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except TextError as e:
            return str(e)
        except Exception as e:  # noqa: BLE001 — 도구는 항상 문자열을 돌려준다
            return f"작업에 실패했습니다: {type(e).__name__}: {e}"

    return wrapper


# ─────────────────────────────── 공용 헬퍼 ───────────────────────────────


def _abspath(path: str) -> str:
    """경로를 정규화하고 존재를 확인한다."""
    if not path:
        raise TextError("path 인자에 파일 경로를 지정하세요.")
    p = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(p):
        raise TextError(f"'{p}' 경로에 파일이 없습니다. 경로를 확인하세요.")
    if os.path.isdir(p):
        raise TextError(f"'{p}'는 폴더입니다. 파일 경로를 지정하세요.")
    return p


def _truncate(text: str, limit: int) -> str:
    text = text or ""
    if limit <= 0 or len(text) <= limit:
        return text
    return text[:limit] + f"\n…(생략 — 전체 {len(text)}자 중 앞 {limit}자. max_chars로 조정하세요)"


def _looks_like_text(s: str) -> bool:
    """디코드된 문자열이 '사람이 읽는 텍스트'처럼 보이는지 휴리스틱으로 판단한다.

    DRM 암호문이 어쩌다 특정 인코딩으로 디코드되더라도 제어문자(NUL 등)가 많아 걸러진다.
    앞부분 표본만 보고 인쇄 가능 문자/공백 비율이 높은지 본다. 빈 파일은 유효로 본다.
    """
    if not s:
        return True
    sample = s[:4000]
    ok = sum(1 for ch in sample if ch.isprintable() or ch in "\n\r\t ")
    return ok / len(sample) >= 0.90


# ─────────────────────────────── 추출 백엔드 ───────────────────────────────
#
# 각 백엔드는 성공 시 텍스트(str)를, "이 백엔드로는 못 읽음(다음으로)"이면 None을 돌려준다.


def _extract_direct(path: str, reasons: list[str]) -> str | None:
    """DRM이 안 걸린(또는 비활성인) 평문 파일을 Python으로 바로 읽는다.

    utf-8 → utf-8-sig(BOM) → cp949 순으로 엄격 디코드한다. 깨끗이 디코드되고 텍스트답게
    보이면 반환한다. 디코드가 깨지거나 암호문처럼 보이면 None(다음 백엔드=word_com).
    """
    try:
        with open(path, "rb") as f:
            data = f.read()
    except OSError as e:
        reasons.append(f"direct: 파일 읽기 실패({e})")
        return None
    for enc in ("utf-8", "utf-8-sig", "cp949"):
        try:
            s = data.decode(enc)
        except UnicodeDecodeError:
            continue
        if _looks_like_text(s):
            # utf-8-sig로 읽었으면 BOM은 이미 벗겨진다.
            return s
        reasons.append(f"direct: {enc}로 디코드했으나 텍스트 같지 않음(DRM 암호문 의심) — Word 경로로")
        return None
    reasons.append("direct: utf-8/cp949로 디코드 실패 — DRM 암호문 또는 바이너리로 보임(Word 경로로)")
    return None


def _extract(path: str) -> tuple[str | None, str, list[str]]:
    """백엔드를 순서대로 시도해 (텍스트, 사용백엔드, 실패사유들)을 돌려준다."""
    reasons: list[str] = []
    text = _extract_direct(path, reasons)
    if text is not None:
        return text, "direct", reasons
    text = extract_via_word(path, reasons, WORD_TIMEOUT)
    if text is not None:
        return text, "word_com", reasons
    return None, "", reasons


def _no_backend_message(path: str, reasons: list[str]) -> str:
    """모든 백엔드가 실패했을 때의 안내 문자열."""
    detail = "\n".join(f"  - {r}" for r in reasons) if reasons else "  - (사유 없음)"
    return (
        f"'{os.path.basename(path)}'에서 텍스트를 읽지 못했습니다. 시도한 경로:\n"
        f"{detail}\n"
        "DRM으로 암호화된 파일은 인증 앱(Word COM)이 복호화를 허용해야 읽힙니다. "
        "`python text_server.py --probe <경로>`로 어떤 경로가 되는지 진단하거나, 회사 IT에 "
        "해당 파일 형식의 읽기 권한을 문의하세요. (바이너리 파일은 이 서버로 읽을 수 없습니다.)"
    )


# ─────────────────────────────── MCP 도구 (🟢 읽기 전용) ───────────────────────────────


@mcp.tool()
@text_tool
def read_text_file(path: str, max_chars: int = MAX_CHARS) -> str:
    """텍스트/코드 파일의 내용을 읽어옵니다. (🟢 읽기 전용)

    DRM으로 암호화된 파일도 인증 앱 경로(Word COM)로 복호화되면 읽습니다.

    Args:
        path: 읽을 파일의 전체 경로 (.txt, .md, .csv, .json, 소스코드 등).
        max_chars: 반환 최대 글자 수(기본 20000). 초과하면 잘라내고 안내를 붙입니다.
    """
    p = _abspath(path)
    text, _backend, reasons = _extract(p)
    if text is None:
        return _no_backend_message(p, reasons)
    return _truncate(text, max_chars)


@mcp.tool()
@text_tool
def text_status(path: str = "") -> str:
    """추출 백엔드의 가용 상태를 보고합니다. path를 주면 그 파일로 실제 진단합니다. (🟢 읽기 전용)

    파일이 안 읽히거나 어떤 경로(direct/word_com)가 되는지 궁금할 때 먼저 호출하세요.
    """
    lines = ["텍스트 추출 백엔드 상태:"]
    lines.append("  - direct(python decode): 항상 가용")
    lines.append(f"  - word_com(pywin32): {'가용' if COM_AVAILABLE else '비활성 — ' + COM_IMPORT_ERROR}")

    if not path:
        lines.append("\npath 인자에 파일 경로를 주면 각 백엔드로 실제 추출을 진단합니다.")
        return "\n".join(lines)

    p = _abspath(path)
    lines.append(f"\n진단 대상: {p}")
    text, backend, reasons = _extract(p)
    if text is not None:
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
        description="DRM 텍스트/코드 파일 읽기 MCP 서버 (읽기 전용)"
    )
    parser.add_argument(
        "--probe", metavar="FILE",
        help="각 추출 백엔드를 지정 파일로 진단하고 종료 (MCP 서버를 띄우지 않음).",
    )
    parser.add_argument(
        "--transport", choices=["stdio", "http", "sse"],
        default=os.getenv("TEXT_MCP_TRANSPORT", "stdio"),
        help="stdio(기본): 로컬 클라이언트가 직접 실행. http/sse: n8n 등 네트워크 접속.",
    )
    parser.add_argument("--host", default=os.getenv("TEXT_MCP_HOST", "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.getenv("TEXT_MCP_PORT", "8093")))
    args = parser.parse_args()

    if args.probe:
        # CLI 진단 모드: text_status를 직접 호출한다(text_tool이 예외를 처리).
        print(text_status(args.probe))
        sys.exit(0)

    if not COM_AVAILABLE:
        print(f"[주의] pywin32 없음({COM_IMPORT_ERROR}) — word_com 백엔드 비활성 "
              "(DRM 파일은 못 읽고 평문만 direct로 읽습니다).", file=sys.stderr)

    if args.transport in ("http", "sse"):
        path = "/mcp/" if args.transport == "http" else "/sse/"
        print(f"TEXT MCP 서버 시작 ({args.transport}) — http://{args.host}:{args.port}{path}",
              file=sys.stderr)
        mcp.run(transport=args.transport, host=args.host, port=args.port)
    else:
        # stdio: stdout은 MCP 프로토콜 채널 — 로그는 stderr로.
        print("TEXT MCP 서버 시작 (stdio)", file=sys.stderr)
        mcp.run(transport="stdio")
