"""llama_proc.py

llama-server를 두 가지 방식 중 하나로 확보한다.

    managed  : 이 앱이 llama-server 프로세스를 직접 띄우고 소유한다 (기본).
    external : serve_llm.py 등으로 **이미 떠 있는** 서버에 붙기만 한다.

external 모드는 LLM 서버(serve_llm.py)와 인터페이스(llm_studio)를 분리해서
독립적으로 돌리기 위한 것이다. 이때 서버의 주인은 이쪽이 아니므로,
**절대 그 프로세스를 종료시키지 않는다** — 앱을 껐다 켜도 모델은 계속 떠 있다.
config의 llama_external_url이 비어 있지 않으면 external 모드가 된다.

managed 모드에서는 도구 호출(MCP)을 위해 --jinja를 항상 켠다. external 모드에서는
서버를 띄운 쪽이 --jinja를 켰는지 여기서 강제할 수 없다 (serve_llm.py는 기본으로 켠다).
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import time
import urllib.error
import urllib.request
from pathlib import Path

from .config import app_dir

# Gemma 3 계열 권장 샘플링. 서버 기본값으로 두고, 요청별 값이 이를 덮어쓴다.
GEMMA_SAMPLING = ["--temp", "1.0", "--top-p", "0.95", "--top-k", "64", "--min-p", "0.0"]


def _as_int(value, default: int) -> int:
    """config 값을 정수로 강제한다. None·빈값·이상값이면 default로 물러선다.

    UI에서 숫자 입력란을 비우면 값이 null로 저장될 수 있고, config.json을 손으로
    잘못 고칠 수도 있다. 그 값이 그대로 llama-server 인자(-c/-ngl 등)에 들어가면
    프로세스가 죽으므로, 여기서 방어한다(설정 오염이 서빙을 막지 못하게 한다)."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return default

HEALTH_TIMEOUT_SEC = 600
# 외부 서버는 이미 떠 있어야 하므로 오래 기다리지 않는다. 다만 모델 로드 직후일 수
# 있어 약간의 여유는 준다 (managed의 600초는 이쪽엔 과하다 — 없는 서버를 기다리게 됨).
ATTACH_TIMEOUT_SEC = 30


class LlamaServerError(RuntimeError):
    pass


class LlamaServer:
    """llama-server 하나를 확보한다 — 직접 띄우거나(managed), 붙거나(external)."""

    def __init__(self, data_dir: Path):
        self.data_dir = data_dir
        self.proc: subprocess.Popen | None = None
        self.model_path: Path | None = None
        self.port: int = 8000
        self.alias: str = "local-model"
        self.external_url: str | None = None  # 값이 있으면 external 모드
        self._log_handle = None
        # 프로세스가 떠 있는 것(running)과 헬스체크를 통과해 요청을 받을 수 있는 것(ready)은
        # 다르다. 모델 로딩 중에는 running=True지만 ready=False다.
        self._ready = False

    # ---------- 탐색 ----------

    def find_binary(self) -> Path:
        """llama-server 실행 파일을 찾는다. 환경 변수 > 앱 폴더 동봉 > 데이터 폴더 > PATH 순."""
        candidates = [
            os.environ.get("LLAMA_SERVER_BIN"),
            app_dir() / "llama" / "llama-server.exe",
            app_dir() / "llama-server.exe",
            self.data_dir / "llama" / "llama-server.exe",
            shutil.which("llama-server"),
            shutil.which("llama-server.exe"),
        ]
        for cand in candidates:
            if cand and Path(cand).is_file():
                return Path(cand)
        raise LlamaServerError(
            "llama-server 실행 파일을 찾지 못했습니다. "
            "앱 폴더의 llama\\ 하위에 두거나 LLAMA_SERVER_BIN 환경 변수로 지정하세요."
        )

    def find_model(self, configured: str = "") -> Path:
        """모델 파일을 찾는다. 설정값이 있으면 그것을, 없으면 models/의 최신 .gguf."""
        if configured:
            path = Path(configured)
            if path.is_file():
                return path
            raise LlamaServerError(f"설정된 모델 파일이 없습니다: {configured}")
        ggufs = sorted(
            (self.data_dir / "models").glob("*.gguf"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not ggufs:
            raise LlamaServerError(
                f"모델이 없습니다. {self.data_dir / 'models'} 폴더에 .gguf 파일을 넣으세요."
            )
        return ggufs[0]

    def list_models(self) -> list[dict]:
        """models/ 폴더의 .gguf 파일 목록을 최신순으로 돌려준다 (UI 모델 선택용).

        각 항목: {"name", "path", "size", "mtime"}. 폴더가 없거나 비어 있으면 빈 목록.
        직접 경로(폴더 밖)를 쓰고 싶으면 UI에서 경로를 손으로 넣는다 — 여기선 스캔만 한다.
        """
        models_dir = self.data_dir / "models"
        items: list[dict] = []
        try:
            for p in models_dir.glob("*.gguf"):
                try:
                    st = p.stat()
                except OSError:
                    continue
                items.append(
                    {"name": p.name, "path": str(p), "size": st.st_size, "mtime": st.st_mtime}
                )
        except OSError:
            pass
        return sorted(items, key=lambda m: m["mtime"], reverse=True)

    # ---------- 실행 ----------

    def start(self, config: dict) -> None:
        """설정에 따라 llama-server를 확보한다.

        llama_external_url이 있으면 이미 떠 있는 서버에 붙고(external),
        없으면 직접 띄운다(managed).
        """
        self._ready = False  # 시작할 때마다 준비 상태를 초기화한다
        external = (config.get("llama_external_url") or "").strip()
        if external:
            self._attach(external, config)
            return

        self.external_url = None
        binary = self.find_binary()
        self.model_path = self.find_model(config.get("model_path", ""))
        self.port = _as_int(config.get("llama_port"), 8000)
        self.alias = config.get("model_alias", "local-model")

        cmd = [
            str(binary),
            "-m", str(self.model_path),
            "--host", "127.0.0.1",  # llama-server는 앱 내부 전용. 외부 공개는 UI 서버가 맡는다.
            "--port", str(self.port),
            "-c", str(_as_int(config.get("ctx"), 32768)),
            # GPU 오프로드 레이어 수. 기본 99(전 레이어)는 VRAM이 넉넉한 서버(A4000 16GB)용.
            # VRAM이 작으면(노트북 dGPU 등) 이 값을 낮춰 일부만 올린다. 0이면 CPU 전용.
            "-ngl", str(_as_int(config.get("gpu_layers"), 99)),
            "--alias", self.alias,
            "--jinja",  # chat template 기반 도구 호출(function calling) 활성화
            *GEMMA_SAMPLING,
        ]
        if config.get("kv_quant"):
            cmd += ["-ctk", "q8_0", "-ctv", "q8_0", "-fa", "on"]

        log_path = self.data_dir / "logs" / "llama-server.log"
        self._log_handle = open(log_path, "a", encoding="utf-8", errors="replace")
        self._log_handle.write(f"\n===== 시작: {' '.join(cmd)} =====\n")
        self._log_handle.flush()

        print(f"[정보] llama-server 시작: {self.model_path.name} (포트 {self.port})")
        print(f"[정보] 로그: {log_path}")
        self.proc = subprocess.Popen(
            cmd, stdout=self._log_handle, stderr=subprocess.STDOUT,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        )
        if not self._wait_healthy():
            code = self.proc.poll()
            self.stop()
            raise LlamaServerError(
                f"llama-server가 준비되지 않았습니다 (exit={code}). "
                f"{log_path}에서 원인을 확인하세요 (VRAM 부족, 모델 손상 등)."
            )
        self._ready = True  # 헬스체크 통과 — 이제 요청을 받을 수 있다
        print("[정보] llama-server 준비 완료")

    def _attach(self, url: str, config: dict) -> None:
        """이미 떠 있는 llama-server(serve_llm.py 등)에 붙는다.

        프로세스를 띄우지 않으므로 바이너리/모델 파일이 없어도 된다.
        연결에 실패하면 LlamaServerError를 던져 앱이 목 모드로 물러서게 한다.
        """
        # "127.0.0.1:8000"처럼 스킴이 빠진 주소를 보정한다. urlparse는 "host:port"의
        # host를 스킴으로 오인하므로(숫자·점도 스킴 문자라서) 접두사로 직접 판별한다.
        if not url.lower().startswith(("http://", "https://")):
            url = "http://" + url
        self.external_url = url.rstrip("/")
        self.proc = None
        self.model_path = None
        self.alias = config.get("model_alias", "local-model")

        print(f"[정보] 외부 llama-server에 연결 시도: {self.external_url}")
        if not self._wait_healthy():
            failed = self.external_url
            self.external_url = None
            raise LlamaServerError(
                f"외부 llama-server({failed})에 연결하지 못했습니다. "
                f"serve_llm.py 등으로 먼저 띄웠는지, 주소가 맞는지 확인하세요."
            )
        self._ready = True  # 외부 서버 헬스체크 통과 — 요청을 받을 수 있다
        # 모델 이름은 서버가 알고 있다 — /v1/models에서 읽어 상태 표시에 쓴다.
        self.alias = self._fetch_alias() or self.alias
        print(f"[정보] 외부 서버 연결됨 (모델: {self.alias})")

    def _fetch_alias(self) -> str | None:
        """붙은 서버의 모델 id(=alias)를 /v1/models에서 읽는다. 실패하면 None."""
        try:
            with urllib.request.urlopen(f"{self.base_url}/v1/models", timeout=5) as res:
                data = json.loads(res.read().decode("utf-8"))
            models = data.get("data") or []
            if models:
                return models[0].get("id")
        except (urllib.error.URLError, OSError, ValueError, KeyError):
            pass
        return None

    def _wait_healthy(self) -> bool:
        # external은 이미 떠 있어야 하므로 짧게, managed는 모델 로드까지 길게 기다린다.
        external = self.external_url is not None
        deadline = time.time() + (ATTACH_TIMEOUT_SEC if external else HEALTH_TIMEOUT_SEC)
        while time.time() < deadline:
            # managed에서 프로세스가 죽었으면 더 기다릴 이유가 없다.
            # external은 우리가 띄운 프로세스가 없으므로 이 검사를 건너뛴다.
            if not external and (self.proc is None or self.proc.poll() is not None):
                return False
            try:
                with urllib.request.urlopen(f"{self.base_url}/health", timeout=3) as res:
                    if res.status == 200:
                        return True
            except (urllib.error.URLError, OSError):
                pass
            time.sleep(2)
        return False

    def stop(self) -> None:
        self._ready = False
        if self.external_url is not None:
            # 외부 서버는 우리 것이 아니다 — 프로세스를 건드리지 않고 연결만 끊는다.
            # (serve_llm.py로 띄운 서버는 이 앱을 껐다 켜도 계속 떠 있어야 한다.)
            self.external_url = None
            return
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()
            try:
                self.proc.wait(timeout=10)
            except subprocess.TimeoutExpired:
                self.proc.kill()
        self.proc = None
        if self._log_handle:
            self._log_handle.close()
            self._log_handle = None

    def restart(self, config: dict) -> None:
        self.stop()
        self.start(config)

    # ---------- 상태 ----------

    @property
    def base_url(self) -> str:
        if self.external_url is not None:
            return self.external_url
        return f"http://127.0.0.1:{self.port}"

    @property
    def running(self) -> bool:
        if self.external_url is not None:
            # 외부 서버는 우리가 소유하지 않으므로 프로세스 핸들이 없다.
            # 연결에 성공해 external_url이 살아 있으면 running으로 본다.
            return True
        return self.proc is not None and self.proc.poll() is None

    @property
    def ready(self) -> bool:
        """헬스체크를 통과해 실제로 요청을 받을 수 있는 상태인지.

        running(프로세스가 떠 있음)과 구분한다 — 모델 로딩 중에는 프로세스는 살아
        있지만(running=True) 아직 준비되지 않았다(ready=False). 채팅 가드와 UI는
        이 값으로 '진짜 서빙 가능'을 판단한다(로딩 중인 서버로 요청을 보내지 않게)."""
        return self.running and self._ready

    def status(self) -> dict:
        external = self.external_url is not None
        return {
            "running": self.running,
            "ready": self.ready,
            # external은 로컬 모델 파일이 없다 — 서버에서 읽어온 alias를 모델명으로 쓴다.
            "model": self.model_path.name if self.model_path else (self.alias if external else None),
            "alias": self.alias,
            "url": self.base_url if self.ready else None,
            "external": external,
        }
