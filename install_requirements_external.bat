@echo off
setlocal
rem Installs requirements.txt on an INTERNET-CONNECTED PC using pip.
rem This is the online twin of install_requirements.bat: it pulls straight from
rem the public PyPI (pypi.org), so there is NO internal mirror and NO wheelhouse.
rem
rem Use this to build/verify the venv on a dev PC that has internet. For the
rem closed network, use install_requirements.bat (mirror + wheelhouse) instead.
rem
rem Installs into a local venv\ virtualenv (created if missing). Set USE_VENV=0
rem below to use the system python.
cd /d "%~dp0"

rem ===== EDIT: install into a local venv\ folder (1) or the system python (0) =====
set "USE_VENV=1"

rem ===== EDIT: also install llm_studio\requirements.txt into the same venv (1/0) =====
rem The shared venv then runs both the MCP servers and the LocalLLM Studio app.
set "INSTALL_LLM_STUDIO=1"

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

rem ===== Make sure pip itself is fresh (optional but avoids old-pip surprises) =====
"%PY%" -m pip install --upgrade pip

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
echo Check your internet connection and that the package names in the
echo requirements file are available on pypi.org.
pause
endlocal
exit /b 1

rem ----- installs one requirements file (%~1) from the public PyPI -----
:install_reqs
echo Installing %~1 from public PyPI (pypi.org)
"%PY%" -m pip install -r "%~1"
exit /b %errorlevel%
