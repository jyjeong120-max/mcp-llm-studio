@echo off
rem Launches the DRM text/code file reader MCP server (read-only).
rem Double-click -> runs over HTTP (:8093) so n8n / llm_studio can connect by URL.
rem Any arguments are passed through as-is, e.g.: run_text_server.bat --transport stdio
rem Diagnose a file:  run_text_server.bat --probe C:\path	oile.txt
rem Uses the venv\ virtualenv at the repo root if present, otherwise system python.
rem NOTE: reading DRM-encrypted files needs Word COM (pywin32 + Word) in the user's
rem       logged-in session. Plain (non-DRM) files are read directly. Without pywin32
rem       the server still starts; DRM files then return guidance messages only.
cd /d "%~dp0"
set "PY=python"
if exist "..env\Scripts\python.exe" set "PY=..env\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" text_server.py --transport http
) else (
    "%PY%" text_server.py %*
)
pause
