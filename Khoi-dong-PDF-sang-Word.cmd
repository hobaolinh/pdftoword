@echo off
setlocal
set "PYTHON=C:\Program Files\WindowsApps\OpenAI.CodexPrimaryRuntime.v26-1105-1297-0_26.1105.1297.0_x64__3k8sg7r9htsxt\dependencies\python\python.exe"
set "LAUNCH=C:\Users\Linhhb\Documents\Codex\2026-11-07\h-y-x-y-d-ng\work\khoi-dong-ocr.py"
set "PYTHONDONTWRITEBYTECODE=1"
set "PYTHONUTF8=1"
chcp 65001 >nul
"%PYTHON%" "%LAUNCH%" %*
if errorlevel 1 pause
