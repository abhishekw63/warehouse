@echo off
REM Daily TiDB -> local MySQL backup (run by Windows Task Scheduler).
REM Overwrites local MySQL to mirror the TiDB server. Logs to logs\tidb_backup.log.
cd /d "d:\PO tracking\Automation\warehouse"
echo ============================================================ >> "logs\tidb_backup.log"
"C:\Users\renee\AppData\Local\Programs\Python\Python313\python.exe" manage.py backup_local >> "logs\tidb_backup.log" 2>&1
