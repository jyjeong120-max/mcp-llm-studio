@echo off
rem Launches the ANSYS MAPDL thermal-analysis MCP server (PyMAPDL, gRPC).
rem Double-click -> runs over HTTP (:8091) so n8n / llm_studio can connect by URL.
rem Any arguments are passed through as-is, e.g.: run_ansys_server.bat --transport stdio
rem Uses the venv\ virtualenv at the repo root if present, otherwise system python.
rem NOTE: requires ANSYS MAPDL installed + license on this PC (checked out on use).
rem       Without the ansys-mapdl-core package the server still starts and the
rem       tools return guidance messages only.
cd /d "%~dp0"
set "PY=python"
if exist "..\venv\Scripts\python.exe" set "PY=..\venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" ansys_server.py --transport http
) else (
    "%PY%" ansys_server.py %*
)
pause
