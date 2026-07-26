@echo off
rem Launches the outlook MCP server (mail/calendar/contacts, confirm-gated writes).
rem Double-click -> runs over HTTP (:8088) so n8n / llm_studio can connect by URL.
rem Any arguments are passed through as-is, e.g.: run_outlook_server.bat --transport stdio
rem Uses the venv\ virtualenv at the repo root if present, otherwise system python.
rem NOTE: must run in the user session where Outlook is logged in (COM restriction).
cd /d "%~dp0"
set "PY=python"
if exist "..\venv\Scripts\python.exe" set "PY=..\venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" outlook_server.py --transport http
) else (
    "%PY%" outlook_server.py %*
)
pause
