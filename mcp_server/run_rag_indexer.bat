@echo off
rem RAG index building CLI (build only - serving is run_rag_server.bat).
rem Usage: run_rag_indexer.bat ..\rag_docs           (index the rag_docs folder, incremental)
rem        run_rag_indexer.bat ..\rag_docs --reindex (rebuild all, with embeddings)
rem        run_rag_indexer.bat C:\any\folder         (or point at any folder)
rem        run_rag_indexer.bat --status           (show index status)
rem        run_rag_indexer.bat --clear --yes      (delete the whole index)
rem Running without arguments shows status and usage.
rem Uses the venv\ virtualenv at the repo root if present, otherwise system python.
rem NOTE: refuses to start while rag_server (MCP serving) holds the Qdrant lock.
rem       Uses Word COM, so run in the logged-in user session.
cd /d "%~dp0"
set "PY=python"
if exist "..\venv\Scripts\python.exe" set "PY=..\venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" rag_indexer.py --status
    echo.
    "%PY%" rag_indexer.py --help
) else (
    "%PY%" rag_indexer.py %*
)
pause
