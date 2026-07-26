@echo off
rem Launches the office MCP server (Word/PowerPoint read + Excel write).
rem Double-click -> runs over HTTP (:8087) so n8n / llm_studio can connect by URL.
rem Any arguments are passed through as-is, e.g.: run_office_server.bat --transport stdio
rem Uses the venv\ virtualenv at the repo root if present, otherwise system python.
rem NOTE: must run in the user session where Office is logged in (COM restriction).
cd /d "%~dp0"
set "PY=python"
if exist "..\venv\Scripts\python.exe" set "PY=..\venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" office_server.py --transport http
) else (
    "%PY%" office_server.py %*
)
pause
