@echo off
rem Launches the document RAG serving MCP server (hybrid search, read-only).
rem Index building (adding/removing documents) is handled by run_rag_indexer.bat.
rem Double-click -> runs over HTTP (:8090) so n8n / llm_studio can connect by URL.
rem Any arguments are passed through as-is, e.g.: run_rag_server.bat --transport stdio
rem Uses the venv\ virtualenv at the repo root if present, otherwise system python.
rem NOTE: Qdrant local storage is single-process locked - stop this server
rem       while running the indexer.
cd /d "%~dp0"
set "PY=python"
if exist "..\venv\Scripts\python.exe" set "PY=..\venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" rag_server.py --transport http
) else (
    "%PY%" rag_server.py %*
)
pause
