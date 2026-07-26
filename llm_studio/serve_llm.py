"""serve_llm.py

로컬 GGUF 모델(예: Gemma 12B IT QAT)을 llama.cpp의 llama-server로 서빙한다.
**LLM 서버 역할만** 맡는 독립 실행 스크립트다 (UI 없음).

llama-server는 OpenAI 호환 API(/v1/chat/completions, 스트리밍 포함)를 열어주므로,
브라우저의 HTML 페이지나 LangChain(base_url 지정) 어디서든 같은 방식으로 붙을 수 있다.
llm_studio도 `--llama-url`로 여기에 붙을 수 있다 (아래 참고).

도구 호출(function calling)은 --jinja로 기본 활성화된다. 붙는 쪽이 대부분 도구를
쓰므로 켜는 게 기본이고, 끄려면 --no-jinja를 준다.

이 스크립트가 하는 일:
1. llama-server 실행 파일과 GGUF 모델 파일 위치를 확인한다.
2. GPU 오프로드, 컨텍스트 길이, 샘플링 등 권장 옵션으로 llama-server를 띄운다.
3. /health 엔드포인트를 폴링해 모델 로드가 끝날 때까지 기다린 뒤 접속 정보를 출력한다.
4. Ctrl+C를 누르면 서버 프로세스를 정리하고 종료한다.

llm_studio(인터페이스)와의 관계 — 둘은 독립적으로 실행된다:
    (a) 각자 실행    : llm_studio가 자기 llama-server를 직접 띄운다 (기본 동작).
    (b) 분리해서 연결 : 이 스크립트로 LLM 서버를 먼저 띄운 뒤,
                       `python app.py --llama-url http://127.0.0.1:8000`으로 붙인다.
                       모델을 다시 로드하지 않고 여러 클라이언트가 한 서버를 공유한다.
    (b)에서는 이 스크립트가 서버의 주인이다. llm_studio를 껐다 켜도 모델은 계속 떠 있고,
    llm_studio는 이 프로세스를 종료시키지 않는다.
    여러 클라이언트가 동시에 붙는다면 --parallel로 슬롯을 늘릴 것 (기본 1은 순차 처리).

사전 준비 (폐쇄망 반입물):
- llama-server 실행 파일: llama.cpp 릴리스의 Windows CUDA 빌드
  (llama-bXXXX-bin-win-cuda-x64.zip 압축 해제, 같은 폴더의 DLL 포함)
- 모델 파일: gemma 12B IT QAT .gguf 1개

사용 예시:
    python serve_llm.py --model C:/models/gemma-12b-it-qat.gguf
    python serve_llm.py --model C:/models/gemma-12b-it-qat.gguf --host 0.0.0.0
    python serve_llm.py --model ... --server-bin C:/llama.cpp/llama-server.exe
    python serve_llm.py --model ... --ctx 65536 --kv-quant   # 컨텍스트 확장 시 KV 캐시 양자화
    python serve_llm.py --model ... --parallel 4             # 여러 클라이언트가 붙을 때

환경 변수로도 지정 가능 (인자가 우선):
    LLM_MODEL_PATH    모델 .gguf 경로
    LLAMA_SERVER_BIN  llama-server 실행 파일 경로
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.request

# Gemma 3 계열의 권장 샘플링 값. 직접 서빙에서는 이 값들이 응답 품질을 좌우하므로
# 서버 기본값에 맡기지 않고 명시적으로 지정한다. (build_agent.py의 VLLM_DECODING과 같은 취지)
GEMMA_SAMPLING = {
    "--temp": "1.0",
    "--top-p": "0.95",
    "--top-k": "64",
    "--min-p": "0.0",
}

HEALTH_TIMEOUT_SEC = 600  # 모델 로드 대기 한도. 12B QAT는 디스크 속도에 따라 수 분 걸릴 수 있다.


def find_server_bin(cli_value: str | None) -> str:
    """llama-server 실행 파일을 찾는다. 인자 > 환경 변수 > PATH 순."""
    candidates = [
        cli_value,
        os.environ.get("LLAMA_SERVER_BIN"),
        shutil.which("llama-server"),
        shutil.which("llama-server.exe"),
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    sys.exit(
        "[오류] llama-server 실행 파일을 찾지 못했습니다.\n"
        "  --server-bin 인자나 LLAMA_SERVER_BIN 환경 변수로 경로를 지정하거나,\n"
        "  llama-server.exe가 있는 폴더를 PATH에 추가하세요.\n"
        "  (llama.cpp 릴리스의 Windows CUDA 빌드 zip을 풀면 나옵니다)"
    )


def find_model(cli_value: str | None) -> str:
    """GGUF 모델 파일을 찾는다. 인자 > 환경 변수 순."""
    path = cli_value or os.environ.get("LLM_MODEL_PATH")
    if not path:
        sys.exit("[오류] 모델 경로가 없습니다. --model 인자나 LLM_MODEL_PATH 환경 변수로 지정하세요.")
    if not os.path.isfile(path):
        sys.exit(f"[오류] 모델 파일이 없습니다: {path}")
    if not path.lower().endswith(".gguf"):
        print(f"[주의] 확장자가 .gguf가 아닙니다: {path} — llama-server는 GGUF 포맷만 읽습니다.")
    return path


def build_command(args, server_bin: str, model_path: str) -> list[str]:
    """llama-server 실행 커맨드를 조립한다."""
    cmd = [
        server_bin,
        "-m", model_path,
        "--host", args.host,
        "--port", str(args.port),
        "-c", str(args.ctx),
        "-ngl", str(args.gpu_layers),  # 99 = 전 레이어 GPU 오프로드 (A4000 16GB에 QAT 12B는 전부 올라간다)
        "--alias", args.alias,         # API에서 보이는 모델 이름 (HTML 쪽 model 필드와 맞춘다)
    ]
    if args.jinja:
        # chat template 기반 function calling(도구 호출)을 켠다. 이게 없으면 tools를
        # 넘겨도 모델이 도구를 부르지 못한다 — 붙는 쪽(LangChain/n8n/llm_studio)이
        # 대부분 도구를 쓰므로 기본값으로 켜 둔다.
        cmd += ["--jinja"]
    for key, value in GEMMA_SAMPLING.items():
        cmd += [key, value]
    if args.kv_quant:
        # KV 캐시를 q8_0으로 양자화해 긴 컨텍스트에서 VRAM을 절약한다. 품질 영향은 미미하다.
        cmd += ["-ctk", "q8_0", "-ctv", "q8_0", "-fa", "on"]
    if args.api_key:
        cmd += ["--api-key", args.api_key]
    if args.parallel > 1:
        cmd += ["-np", str(args.parallel)]  # 동시 요청 슬롯 수 (컨텍스트가 슬롯 수만큼 나뉜다)
    if args.extra:
        cmd += args.extra
    return cmd


def wait_until_healthy(base_url: str, proc: subprocess.Popen) -> bool:
    """모델 로드가 끝나 /health가 200을 줄 때까지 기다린다."""
    deadline = time.time() + HEALTH_TIMEOUT_SEC
    while time.time() < deadline:
        if proc.poll() is not None:
            return False  # 서버가 먼저 죽음 (모델 로드 실패, VRAM 부족 등)
        try:
            with urllib.request.urlopen(f"{base_url}/health", timeout=3) as res:
                if res.status == 200:
                    return True
        except (urllib.error.URLError, OSError):
            pass  # 아직 로드 중
        time.sleep(2)
    return False


def main() -> None:
    parser = argparse.ArgumentParser(
        description="llama-server로 로컬 GGUF 모델을 OpenAI 호환 API로 서빙한다.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--model", help="GGUF 모델 파일 경로 (또는 LLM_MODEL_PATH 환경 변수)")
    parser.add_argument("--server-bin", help="llama-server 실행 파일 경로 (또는 LLAMA_SERVER_BIN 환경 변수)")
    parser.add_argument("--host", default="127.0.0.1",
                        help="바인딩 주소. 같은 망의 다른 PC에서 접속하려면 0.0.0.0")
    parser.add_argument("--port", type=int, default=8000, help="포트")
    parser.add_argument("--ctx", type=int, default=32768, help="컨텍스트 길이 (토큰)")
    parser.add_argument("--gpu-layers", type=int, default=99, help="GPU에 올릴 레이어 수 (99=전부)")
    parser.add_argument("--alias", default="gemma-12b-it-qat", help="API에 노출할 모델 이름")
    parser.add_argument("--kv-quant", action="store_true",
                        help="KV 캐시를 q8_0으로 양자화 (긴 컨텍스트에서 VRAM 절약)")
    parser.add_argument("--no-jinja", dest="jinja", action="store_false",
                        help="chat template 기반 도구 호출을 끈다 (기본은 켬). "
                             "모델의 chat template이 깨져 서버가 안 뜰 때만 쓴다 — "
                             "끄면 도구 호출이 동작하지 않는다")
    parser.set_defaults(jinja=True)
    parser.add_argument("--api-key", default=None,
                        help="지정 시 이 키가 있는 요청만 허용 (Authorization: Bearer <키>)")
    parser.add_argument("--parallel", type=int, default=1, help="동시 처리 슬롯 수")
    parser.add_argument("--extra", nargs=argparse.REMAINDER, default=None,
                        help="이후의 모든 인자를 llama-server에 그대로 전달")
    args = parser.parse_args()

    server_bin = find_server_bin(args.server_bin)
    model_path = find_model(args.model)
    cmd = build_command(args, server_bin, model_path)

    size_gb = os.path.getsize(model_path) / (1024 ** 3)
    print(f"[정보] 서버 바이너리 : {server_bin}")
    print(f"[정보] 모델 파일     : {model_path} ({size_gb:.1f} GB)")
    print(f"[정보] 실행 커맨드   : {' '.join(cmd)}")
    print("[정보] 서버를 시작합니다. 모델 로드에 몇 분 걸릴 수 있습니다...\n")

    # 로그를 그대로 콘솔에 흘려보낸다 (로드 진행 상황, VRAM 사용량 등이 여기 찍힌다).
    proc = subprocess.Popen(cmd)

    # 헬스 체크는 항상 로컬 주소로 한다 (0.0.0.0은 접속 주소가 아니므로).
    check_host = "127.0.0.1" if args.host == "0.0.0.0" else args.host
    base_url = f"http://{check_host}:{args.port}"

    try:
        if not wait_until_healthy(base_url, proc):
            if proc.poll() is not None:
                sys.exit(f"\n[오류] llama-server가 종료됐습니다 (exit code {proc.returncode}). "
                         "위 로그에서 원인을 확인하세요 (VRAM 부족, 모델 파일 손상 등).")
            proc.terminate()
            sys.exit(f"\n[오류] {HEALTH_TIMEOUT_SEC}초 안에 서버가 준비되지 않았습니다.")

        print("\n" + "=" * 60)
        print("서버 준비 완료!")
        print(f"  OpenAI 호환 API : {base_url}/v1/chat/completions")
        print(f"  모델 이름       : {args.alias}")
        print(f"  내장 웹 UI      : {base_url}  (브라우저에서 바로 테스트 가능)")
        if args.host == "0.0.0.0":
            print("  같은 망의 다른 PC에서는 이 컴퓨터의 IP로 접속하세요.")
        print("  종료: Ctrl+C")
        print("=" * 60 + "\n")

        proc.wait()  # 서버가 살아있는 동안 대기
    except KeyboardInterrupt:
        print("\n[정보] 서버를 종료합니다...")
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
    sys.exit(proc.returncode or 0)


if __name__ == "__main__":
    main()
