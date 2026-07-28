"""llama-server 런처 — 임베딩(:8001)·리랭커(:8002) 백엔드 공용.

rag_indexer(구성)와 rag_server(서빙)가 함께 쓴다. 하는 일:
- 루트 `rag_embed/`(임베딩 gguf)·`rag_rerank/`(리랭커 gguf)에서 모델을 자동 선택.
- `llama-server` 실행 파일을 두 폴더 하위(재귀)·LLAMA_SERVER_BIN·PATH에서 찾는다.
- CUDA 배포 zip을 폴더째 넣어도 되도록 exe 폴더 + `cudart*` 폴더를 PATH에 덧댄다.
- 모델/실행파일이 없으면 조용히 저하(서버 없이 진행) — CLAUDE.md의 우아한 저하 규약.
- 우리가 띄운 프로세스만 종료한다(사용자가 따로 띄운 서버는 재사용하고 건드리지 않는다).

llama-server 로직이 인덱서/서버에 흩어지지 않게 여기 한 곳에 모은다. stdout은 stdio
MCP 채널일 수 있으므로 로그는 전부 stderr, 자식 서버의 stdout/stderr는 버린다.
"""

from __future__ import annotations

import atexit
import os
import shutil
import subprocess
import sys
import threading
import time
from urllib.parse import urlparse

import rag_core as core

# 루트의 반입 폴더. mcp_server의 한 단계 위(저장소 루트) 기준. gguf/exe는 gitignore.
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EMBED_DIR = os.path.join(REPO_ROOT, "rag_embed")     # 임베딩 gguf (+ 보통 llama-server.exe)
RERANK_DIR = os.path.join(REPO_ROOT, "rag_rerank")   # 리랭커 gguf
MODEL_DIRS = (EMBED_DIR, RERANK_DIR)                  # 바이너리·cudart 탐색 대상

LOAD_TIMEOUT = float(os.getenv("RAG_LLAMA_LOAD_TIMEOUT", "300"))  # 모델 로드 대기 한도(초)


# ─────────────────────────────── 탐색 헬퍼 ───────────────────────────────


def discover_gguf(model_dir: str, label: str, cli_model: str | None = None) -> str | None:
    """모델 GGUF 경로를 정한다. 인자 > model_dir의 유일한 .gguf. 없으면 None(저하).

    폴더에 .gguf가 여러 개면 자동 선택하지 않는다(무엇을 쓸지 사람이 정하게)."""
    if cli_model:
        if os.path.isfile(cli_model):
            return os.path.abspath(cli_model)
        print(f"[{label}] 지정한 모델 경로에 파일이 없습니다: {cli_model}", file=sys.stderr)
        return None
    if os.path.isdir(model_dir):
        ggufs = sorted(f for f in os.listdir(model_dir) if f.lower().endswith(".gguf"))
        if len(ggufs) == 1:
            return os.path.join(model_dir, ggufs[0])
        if len(ggufs) > 1:
            print(f"[{label}] {os.path.basename(model_dir)}에 .gguf가 여러 개"
                  f"({', '.join(ggufs)})라 자동 선택하지 않습니다 — 모델을 직접 지정하세요.",
                  file=sys.stderr)
    return None


def discover_server_bin(cli_bin: str | None = None) -> str | None:
    """llama-server 실행 파일을 찾는다. 인자 > LLAMA_SERVER_BIN > rag_embed·rag_rerank(재귀) > PATH.

    두 폴더에는 llama.cpp CUDA 배포 zip을 푼 폴더(llama-bXXXX-bin-win-cuda-.../)가
    그대로 들어오는 경우가 많아 하위까지 재귀로 훑는다. 여러 개면 경로가 가장 얕은 것을
    고른다(상위에 직접 둔 것을 우선)."""
    for cand in (cli_bin, os.environ.get("LLAMA_SERVER_BIN")):
        if cand and os.path.isfile(cand):
            return cand
    found: list[str] = []
    for d in MODEL_DIRS:
        if os.path.isdir(d):
            for root, _dirs, files in os.walk(d):
                for name in ("llama-server.exe", "llama-server"):
                    if name in files:
                        found.append(os.path.join(root, name))
    if found:
        found.sort(key=lambda p: (p.count(os.sep), p))  # 얕은 경로 우선, 그다음 사전순
        return found[0]
    return shutil.which("llama-server") or shutil.which("llama-server.exe")


def dll_search_dirs(server_bin: str) -> list[str]:
    """llama-server가 CUDA DLL을 찾도록 PATH에 덧댈 폴더들.

    exe 자신의 폴더(Windows 기본 검색 경로) + rag_embed·rag_rerank 아래 cudart* 폴더.
    llama.cpp CUDA는 본체 zip과 cudart zip이 나뉘어 있어 cudart DLL이 exe와 다른 폴더에
    풀리는 경우가 많다 — 이걸 PATH에 붙여야 폴더 병합 없이도 로드된다."""
    dirs = [os.path.dirname(server_bin)]
    for d in MODEL_DIRS:
        if os.path.isdir(d):
            for name in os.listdir(d):
                p = os.path.join(d, name)
                if os.path.isdir(p) and name.lower().startswith("cudart"):
                    dirs.append(p)
    return dirs


def _port_from_url(url: str, default: int) -> int:
    try:
        return urlparse(url).port or default
    except Exception:  # noqa: BLE001
        return default


# ─────────────────────────────── 서버 프로세스 ───────────────────────────────


class LlamaServer:
    """llama-server 프로세스 하나(임베딩 또는 리랭커)를 관리한다. start/stop.

    - 이미 해당 URL이 응답하면(사용자가 따로 띄워둠) 재사용하고 아무것도 안 띄운다.
    - 모델 또는 실행 파일을 못 찾으면 조용히 저하 — 서버 없이 진행한다.
    - 우리가 띄운 프로세스만 종료한다.
    CLAUDE.md 규약대로 -ngl 0(CPU 상주), 로그는 stderr, 자식 출력은 버린다.
    """

    def __init__(self, *, label: str, kind: str, model: str | None,
                 server_bin: str | None, port: int, ctx: str,
                 is_available, degrade_note: str = "서버 없이 진행합니다") -> None:
        self.label = label            # 로그 접두사 (예: "임베딩", "리랭커")
        self.kind = kind              # "--embeddings" | "--reranking"
        self.model = model
        self.server_bin = server_bin
        self.port = port
        self.ctx = str(ctx)
        self.is_available = is_available  # callable() -> bool (URL 응답 확인)
        self.degrade_note = degrade_note  # 저하 시 안내 문구(맥락별로 다름)
        self.proc: subprocess.Popen | None = None
        self.started = False          # 우리가 띄웠는지 (종료 책임 판단)

    def start(self) -> "LlamaServer":
        if self.is_available():
            print(f"[{self.label}] 이미 실행 중인 서버를 사용합니다.", file=sys.stderr)
            return self
        if not self.model:
            print(f"[{self.label}] 모델 .gguf가 없습니다 — {self.degrade_note}.", file=sys.stderr)
            return self
        if not self.server_bin:
            print(f"[{self.label}] llama-server 실행 파일을 못 찾았습니다 "
                  f"(LLAMA_SERVER_BIN 또는 rag_embed/rag_rerank) — {self.degrade_note}.",
                  file=sys.stderr)
            return self
        # 임베딩/리랭커 모두 입력(청크, 또는 질의+문서 쌍)이 물리 배치(-ub) 안에 통째로
        # 들어가야 처리된다. 기본 ubatch(512)면 긴 청크가 넘쳐 요청마다 에러 → 저하되므로
        # -b/-ub를 컨텍스트(-c)와 같게 키운다.
        cmd = [self.server_bin, "-m", self.model, "--host", "127.0.0.1",
               "--port", str(self.port), self.kind, "-ngl", "0",
               "-c", self.ctx, "-b", self.ctx, "-ub", self.ctx]
        print(f"[{self.label}] 서버 기동: {os.path.basename(self.model)} (:{self.port}, CPU) — "
              "모델 로드에 수십 초 걸릴 수 있습니다.", file=sys.stderr)
        env = os.environ.copy()
        env["PATH"] = os.pathsep.join(dll_search_dirs(self.server_bin) + [env.get("PATH", "")])
        try:
            # 자식 로그는 버린다(서버가 자체 stderr에 찍음). 우리 stdout(MCP 채널)과 안 섞이게.
            self.proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL,
                                         stderr=subprocess.DEVNULL, env=env)
        except OSError as e:
            print(f"[{self.label}] 서버 기동 실패({e}) — {self.degrade_note}.", file=sys.stderr)
            self.proc = None
            return self
        self.started = True
        if self._wait_ready():
            print(f"[{self.label}] 준비 완료.", file=sys.stderr)
        else:
            print(f"[{self.label}] 로드 시간 초과/실패 — {self.degrade_note}.", file=sys.stderr)
            self.stop()
        return self

    def _wait_ready(self) -> bool:
        deadline = time.time() + LOAD_TIMEOUT
        while time.time() < deadline:
            if self.proc is not None and self.proc.poll() is not None:
                return False  # 서버가 로드 중 죽음(모델 손상·메모리 부족 등)
            if self.is_available():
                return True
            time.sleep(2)
        return False

    def stop(self) -> None:
        if self.proc is None:
            return
        try:
            self.proc.terminate()
            try:
                self.proc.wait(timeout=10)
            except subprocess.TimeoutExpired:
                self.proc.kill()
        except Exception:  # noqa: BLE001
            pass
        self.proc = None

    def __enter__(self) -> "LlamaServer":
        return self.start()

    def __exit__(self, *exc) -> None:
        if self.started and self.proc is not None:
            print(f"[{self.label}] 서버를 종료합니다.", file=sys.stderr)
        self.stop()


# ─────────────────────────────── 빌더 ───────────────────────────────


def make_embed_server(cli_model: str | None = None, cli_bin: str | None = None,
                      *, spawn: bool = True, degrade_note: str = "키워드 전용으로 진행합니다"
                      ) -> LlamaServer:
    """임베딩(llama-server --embeddings, rag_embed의 gguf) 서버 핸들.

    spawn=False면 모델/바이너리를 찾지 않아 새로 띄우지 않는다(이미 떠 있으면 재사용)."""
    model = discover_gguf(EMBED_DIR, "임베딩", cli_model) if spawn else None
    server_bin = discover_server_bin(cli_bin) if spawn else None
    return LlamaServer(
        label="임베딩", kind="--embeddings", model=model, server_bin=server_bin,
        port=_port_from_url(core.EMBED_URL, 8001), ctx=os.getenv("RAG_EMBED_CTX", "4096"),
        is_available=core._embed_available, degrade_note=degrade_note)


def make_rerank_server(cli_model: str | None = None, cli_bin: str | None = None,
                       *, spawn: bool = True, degrade_note: str = "RRF 순서를 그대로 씁니다"
                       ) -> LlamaServer:
    """리랭커(llama-server --reranking, rag_rerank의 gguf) 서버 핸들."""
    model = discover_gguf(RERANK_DIR, "리랭커", cli_model) if spawn else None
    server_bin = discover_server_bin(cli_bin) if spawn else None
    return LlamaServer(
        label="리랭커", kind="--reranking", model=model, server_bin=server_bin,
        port=_port_from_url(core.RERANK_URL, 8002), ctx=os.getenv("RAG_RERANK_CTX", "4096"),
        is_available=core._rerank_available, degrade_note=degrade_note)


def start_in_background(servers: list[LlamaServer]):
    """서빙용: 여러 llama-server를 데몬 스레드에서 기동(블로킹 없이)하고 종료 시 정리한다.

    모델 로드에 수십 초 걸리므로 메인 스레드(MCP 핸드셰이크)를 막지 않도록 백그라운드로
    띄운다 — 준비 전 검색은 우아하게 저하했다가 준비되면 자동으로 하이브리드가 된다.
    프로세스 종료 시 우리가 띄운 것만 정리한다(atexit). 반환값은 수동 정리용 stop 함수.
    """
    def _run() -> None:
        for s in servers:
            try:
                s.start()
            except Exception as e:  # noqa: BLE001 — 한 백엔드 실패가 나머지를 막지 않게
                print(f"[{s.label}] 기동 중 예외({e}) — 건너뜁니다.", file=sys.stderr)

    def _stop() -> None:
        for s in servers:
            if s.started and s.proc is not None:
                print(f"[{s.label}] 서버를 종료합니다.", file=sys.stderr)
            s.stop()

    atexit.register(_stop)
    threading.Thread(target=_run, name="rag-llama-backends", daemon=True).start()
    return _stop
