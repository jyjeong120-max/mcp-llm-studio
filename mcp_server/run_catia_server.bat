@echo off
rem Launches the CATIA V5 MCP server (part/sketch/3D features).
rem Double-click -> runs over HTTP (:8089) so n8n / llm_studio can connect by URL.
rem Any arguments are passed through as-is, e.g.: run_catia_server.bat --transport stdio
rem Uses the venv\ virtualenv at the repo root if present, otherwise system python.
rem NOTE: must run in the user session where CATIA is running (COM restriction).
cd /d "%~dp0"
set "PY=python"
if exist "..\venv\Scripts\python.exe" set "PY=..\venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" catia_server.py --transport http
) else (
    "%PY%" catia_server.py %*
)
pause
