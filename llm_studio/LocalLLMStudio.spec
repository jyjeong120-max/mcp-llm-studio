# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 빌드 스펙: pyinstaller --noconfirm LocalLLMStudio.spec
#
# onedir 방식(폴더 배포)을 쓴다. onefile은 실행마다 임시 폴더에 풀어야 해서
# 시작이 느리고, llama-server/DLL을 함께 두기도 onedir이 깔끔하다.

a = Analysis(
    ['app.py'],
    pathex=[],
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
