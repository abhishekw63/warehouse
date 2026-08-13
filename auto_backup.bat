@echo off
REM ============================================================================
REM Auto TiDB -> local MySQL backup. Runs the Django management command with
REM --if-due, so it only actually backs up once per day (skips if a successful
REM backup ran in the last 20h). Registered in Windows Task Scheduler with a
REM daily trigger + at-logon, StartWhenAvailable (catches up a day missed while
REM the laptop was off). Output is appended to logs\auto_backup.log.
REM ============================================================================
cd /d "d:\PO tracking\Automation\warehouse"
py manage.py backup_local --if-due >> "logs\auto_backup.log" 2>&1
