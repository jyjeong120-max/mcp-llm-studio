@echo off
rem Launches the DRM PDF text-extraction MCP server (read-only).
rem Double-click -> runs over HTTP (:8092) so n8n / llm_studio can connect by URL.
rem Any arguments are passed through as-is, e.g.: run_pdf_server.bat --transport stdio
rem Diagnose a file:  run_pdf_server.bat --probe C:\path\to\file.pdf
rem Uses the venv\ virtualenv at the repo root if present, otherwise system python.
rem NOTE: reading DRM-encrypted PDFs needs Word COM (pywin32 + Word) in the user's
rem       logged-in session. Without pypdf/pywin32 the server still starts and the
rem       tools return guidance messages only (graceful degradation).
cd /d "%~dp0"
set "PY=python"
if exist "..\venv\Scripts\python.exe" set "PY=..\venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" pdf_server.py --transport http
) else (
    "%PY%" pdf_server.py %*
)
pause
