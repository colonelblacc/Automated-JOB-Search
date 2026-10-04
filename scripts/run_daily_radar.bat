@echo off
REM ====================================================================
REM Local Daily Embedded Job Radar Runner for Ajith Shajan
REM Automatically runs every morning at 8:30 AM via Windows Task Scheduler
REM ====================================================================

cd /d "c:\JOB SEARCH\AI TOOL"

set "PYTHON_EXE=C:\Users\colon\scoop\apps\python313\current\python.exe"

echo [%DATE% %TIME%] Running Relevance Engine...
"%PYTHON_EXE%" scripts\relevance_engine.py

echo [%DATE% %TIME%] Updating Excel Tracker...
"%PYTHON_EXE%" scripts\generate_excel_tracker.py

echo [%DATE% %TIME%] Syncing to Google Sheets...
"%PYTHON_EXE%" scripts\sync_to_google_sheets.py

echo [%DATE% %TIME%] Pushing Daily Radar Alert to Telegram...
"%PYTHON_EXE%" scripts\telegram_notifier.py

echo [%DATE% %TIME%] Daily Job Radar completed successfully.
