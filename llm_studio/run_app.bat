@echo off
rem Runs LocalLLM Studio from source (starts idle - model serving is enabled in the UI).
rem Any arguments are passed through as-is, e.g.: run_app.bat --mock (UI only, no model)
rem Virtualenv auto-detection: llm_studio\venv\ first, then repo-root venv\,
rem otherwise system python.
rem First run: pip install -r requirements.txt
cd /d "%~dp0"
set "PY=python"
if exist "..\venv\Scripts\python.exe" set "PY=..\venv\Scripts\python.exe"
if exist "venv\Scripts\python.exe" set "PY=venv\Scripts\python.exe"
"%PY%" app.py %*
pause
