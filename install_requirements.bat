@echo off
setlocal
rem Installs requirements.txt on the closed network using pip.
rem Sources: the local .whl files in .\wheelhouse AND (optionally) the internal
rem PyPI mirror. Edit MIRROR_INDEX/MIRROR_HOST below to enable the mirror; leave
rem them blank to install fully OFFLINE from the wheelhouse only (--no-index).
rem
rem Populate the wheelhouse on an internet-connected PC first, e.g.:
rem   pip download -r requirements.txt -d wheelhouse
rem   pip download -r llm_studio\requirements.txt -d wheelhouse   (if INSTALL_LLM_STUDIO=1)
rem then copy the wheelhouse\ folder over (USB etc.) next to this .bat.
rem
rem Installs into a local venv\ virtualenv (created if missing; python -m venv is
rem offline-safe via ensurepip). Set USE_VENV=0 below to use the system python.
cd /d "%~dp0"

rem ===== EDIT: internal PyPI mirror (leave blank for offline wheelhouse-only) =====
set "MIRROR_INDEX=http://10.42.86.106:8081/repository/py-pi-local/simple"
set "MIRROR_HOST=10.42.86.106:8081"
rem Example:
rem   set "MIRROR_INDEX=http://pypi.company.local/simple"
rem   set "MIRROR_HOST=pypi.company.local"

rem ===== EDIT: also write a per-user pip.ini so the mirror attaches to EVERY pip =====
rem   1 = create %USERPROFILE%\pip\pip.ini (only when it does not exist yet)
rem   0 = skip, and only use this run's command-line options
rem Needs MIRROR_INDEX/MIRROR_HOST above; ignored when they are blank.
set "WRITE_PIP_INI=1"
set "PIPDIR=%USERPROFILE%\pip"

rem ===== EDIT: install into a local venv\ folder (1) or the system python (0) =====
set "USE_VENV=1"

rem ===== EDIT: also install llm_studio\requirements.txt into the same venv (1/0) =====
rem The shared venv then runs both the MCP servers and the LocalLLM Studio app.
set "INSTALL_LLM_STUDIO=1"

set "WHEELS=%~dp0wheelhouse"

set "PY=python"
if "%USE_VENV%"=="1" (
    if not exist "venv\Scripts\python.exe" (
        echo Creating virtualenv: %~dp0venv
        python -m venv venv
    )
    if not exist "venv\Scripts\python.exe" (
        echo [ERROR] Could not create venv - is python on PATH?
        pause
        exit /b 1
    )
    set "PY=venv\Scripts\python.exe"
) else (
    if exist "venv\Scripts\python.exe" set "PY=venv\Scripts\python.exe"
)

if not exist "%WHEELS%" (
    echo [ERROR] wheelhouse folder not found: %WHEELS%
    echo Put the .whl files there ^(see the pip download note at the top^).
    pause
    exit /b 1
)

rem ===== Write per-user pip.ini so future pip calls hit the mirror automatically =====
if "%WRITE_PIP_INI%"=="1" if not "%MIRROR_INDEX%"=="" (
    if not exist "%PIPDIR%" mkdir "%PIPDIR%"
    if exist "%PIPDIR%\pip.ini" (
        echo [SKIP] pip.ini already exists - leaving it as is: %PIPDIR%\pip.ini
    ) else (
        > "%PIPDIR%\pip.ini" echo [global]
        >>"%PIPDIR%\pip.ini" echo index-url=%MIRROR_INDEX%
        >>"%PIPDIR%\pip.ini" echo trusted-host=%MIRROR_HOST%
        echo [OK] Wrote pip.ini: %PIPDIR%\pip.ini
    )
)

call :install_reqs "requirements.txt"
if errorlevel 1 goto :failed

if "%INSTALL_LLM_STUDIO%"=="1" (
    if exist "llm_studio\requirements.txt" (
        call :install_reqs "llm_studio\requirements.txt"
        if errorlevel 1 goto :failed
    ) else (
        echo [SKIP] llm_studio\requirements.txt not found - skipping.
    )
)

echo.
echo [OK] Requirements installed.
pause
endlocal
exit /b 0

:failed
echo.
echo [ERROR] pip install failed - see messages above.
echo If a package is missing from the wheelhouse, download it on an online PC
echo with: pip download ^<name^> -d wheelhouse
pause
endlocal
exit /b 1

rem ----- installs one requirements file (%~1) from mirror+wheelhouse or offline -----
:install_reqs
if "%MIRROR_INDEX%"=="" (
    echo Installing OFFLINE from wheelhouse only: %~1
    "%PY%" -m pip install -r "%~1" --no-index --find-links "%WHEELS%"
) else (
    echo Installing %~1 from mirror %MIRROR_INDEX% + wheelhouse %WHEELS%
    "%PY%" -m pip install -r "%~1" --find-links "%WHEELS%" -i "%MIRROR_INDEX%" --trusted-host "%MIRROR_HOST%"
)
exit /b %errorlevel%
