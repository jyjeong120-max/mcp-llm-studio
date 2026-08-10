# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 빌드 스펙: pyinstaller --noconfirm LocalLLMStudio.spec
#
# onedir 방식(폴더 배포)을 쓴다. onefile은 실행마다 임시 폴더에 풀어야 해서
# 시작이 느리고, llama-server/DLL을 함께 두기도 onedir이 깔끔하다.

import os

# 내장 MCP 도구 서버들이 있는 repo의 mcp_server/ 폴더 — builtin_servers.py가 __import__로
# 지연 로드한다. pathex에 넣어야 PyInstaller가 아래 hiddenimports를 찾아 번들에 담는다.
# (소스 실행은 builtin_servers._mcp_server_dir()가 sys.path에 얹지만, 프리즈된 exe에선 그
#  경로가 없어 실패한다 — 그래서 번들 시점에 PYZ로 포함해야 내장 도구가 exe에서도 뜬다.)
MCP_SERVER_DIR = os.path.join(SPECPATH, '..', 'mcp_server')

a = Analysis(
    ['app.py'],
    pathex=[MCP_SERVER_DIR],
    binaries=[],
    datas=[('static', 'static')],  # 웹 UI 정적 파일을 번들에 포함
    hiddenimports=[
        # uvicorn이 문자열로 지연 임포트하는 모듈들 — 정적 분석에 안 잡힌다
        'uvicorn.logging',
        'uvicorn.loops.auto',
        'uvicorn.loops.asyncio',
        'uvicorn.protocols.http.auto',
        'uvicorn.protocols.http.h11_impl',
        'uvicorn.protocols.websockets.auto',
        'uvicorn.lifespan.on',
        # 선택 의존성: 설치돼 있으면 함께 묶는다
        'pypdf',
        'docx',
        # 내장 MCP 도구 서버(mcp_server/) — builtin_servers.py가 __import__로 지연 로드하므로
        # 정적 분석에 안 잡힌다. 이게 빠지면 exe 배포판에서 내장 도구 4종이 전부 로드 실패한다.
        'pdf_server',
        'text_server',
        'office_server',
        'outlook_server',
        # 위 office/outlook 서버가 COM 조종에 쓰는 pywin32 — 이것도 지연 import라 명시한다
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
    ],
    excludes=['tkinter', 'matplotlib', 'numpy.tests'],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='LocalLLMStudio',
    debug=False,
    upx=False,
    console=True,   # 콘솔 창에 서버 로그를 보여준다. 숨기려면 False
    icon=None,      # .ico 파일이 있으면 여기에 경로 지정
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    name='LocalLLMStudio',
)
