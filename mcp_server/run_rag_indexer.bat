@echo off
rem RAG index building CLI (build only - serving is run_rag_server.bat).
rem Indexes Word (.docx/.doc), PowerPoint (.pptx/.ppt) and PDF (.pdf) files.
rem Auto-starts an embedding server from the repo-root rag_embed\ folder (a
rem single .gguf there is picked automatically), indexes rag_docs, then stops it.
rem If no model/binary is found it degrades to a keyword-only index (no crash).
rem Usage: run_rag_indexer.bat ..\rag_docs           (index the rag_docs folder, incremental)
rem        run_rag_indexer.bat ..\rag_docs --reindex (rebuild all, with embeddings)
rem        run_rag_indexer.bat C:\any\folder         (or point at any folder)
rem        run_rag_indexer.bat ..\rag_docs --no-embed-server  (do not auto-start the embed server)
rem        run_rag_indexer.bat --status           (show index status)
rem        run_rag_indexer.bat --clear --yes      (delete the whole index)
rem Running without arguments indexes the repo-root rag_docs folder (double-click use).
rem Uses the venv\ virtualenv at the repo root if present, otherwise system python.
rem NOTE: refuses to start while rag_server (MCP serving) holds the Qdrant lock.
rem       Uses Word COM for Word/PPT, so run in the logged-in user session.
cd /d "%~dp0"
set "PY=python"
if exist "..\venv\Scripts\python.exe" set "PY=..\venv\Scripts\python.exe"
if "%~1"=="" (
    "%PY%" rag_indexer.py "..\rag_docs"
) else (
    "%PY%" rag_indexer.py %*
)
pause
