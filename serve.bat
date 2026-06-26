@echo off
REM ============================================================================
REM  RENEE B2B PO web app  --  manual start (office LAN)
REM  Double-click this file to start the server. Leave the window OPEN; close
REM  it (or press Ctrl+C) to stop. See HOSTING.md for the full walkthrough.
REM ============================================================================

cd /d "%~dp0"

REM ── Production-ish settings for LAN hosting ────────────────────────────────
set DJANGO_DEBUG=0
REM Trust any host on the internal network. To lock down, list the server's
REM LAN IP + name, comma-separated, e.g.:  set DJANGO_ALLOWED_HOSTS=192.168.1.50,renee-pc
set DJANGO_ALLOWED_HOSTS=*
REM Set a real secret in production (any long random string):
REM set DJANGO_SECRET_KEY=change-me-to-something-long-and-random

set PY=.venv\Scripts\python.exe
if not exist "%PY%" (
    echo [ERROR] .venv not found. Create it first:
    echo     py -3.13 -m venv .venv
    echo     .venv\Scripts\pip install -r requirements.txt
    pause
    exit /b 1
)

REM ── Kill any STALE server still holding port 8000 ──────────────────────────
REM Without this, a previous run keeps the port + serves the OLD code, so your
REM changes never appear. This guarantees a clean restart every time.
echo Stopping any previous server on port 8000...
for /f "tokens=5" %%p in ('netstat -ano ^| findstr /r /c:":8000 .*LISTENING"') do taskkill /F /PID %%p >nul 2>&1

REM ── Refresh static files (cheap; safe to run every start) ──────────────────
echo Collecting static files...
"%PY%" manage.py collectstatic --noinput >nul

REM ── Show the LAN address so people know where to connect ───────────────────
echo.
echo ============================================================
echo   RENEE B2B is starting on port 8000
echo   On THIS machine:        http://localhost:8000/
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do echo   On the office network:  http://%%a:8000/  (try this one)
echo.
echo   Leave this window OPEN. Close it to stop the server.
echo ============================================================
echo.

REM ── Serve via waitress (4 worker threads is plenty for an office) ──────────
"%PY%" -m waitress --listen=0.0.0.0:8000 --threads=4 renee_cosmetics.wsgi:application

echo.
echo Server stopped.
pause
