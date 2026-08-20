@echo off
REM ============================================================================
REM Daily Issues email. Runs the Django management command with --if-not-sent, so
REM it mails the day's issue lines to stakeholders once and won't re-send if it
REM already went out for that date. Register in Windows Task Scheduler with a
REM daily trigger at end-of-day (e.g. 20:00) + StartWhenAvailable (catches up if
REM the laptop was off at the trigger time). Output appends to logs\daily_issues.log.
REM
REM Morning schedule instead? Add --yesterday so it sends the PREVIOUS day's issues.
REM ============================================================================
cd /d "d:\PO tracking\Automation\warehouse"
py manage.py send_daily_issues --if-not-sent >> "logs\daily_issues.log" 2>&1
