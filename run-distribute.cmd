@echo off
REM ============================================================
REM  Runs the Row-Level Report Distributor. Task Scheduler (or a
REM  double-click) points at THIS file. It self-locates via
REM ============================================================

REM %~dp0 = the folder this file lives in (the project folder).
cd /d "%~dp0"

title Row-Level Report Distributor
echo.
echo  Row-Level Report Distributor
echo  Starting up...
echo.

REM Run with the project's own Python environment (created by setup.ps1 / uv sync).
REM Calling the venv's python directly needs no PATH lookup, which is the most
REM reliable option under Task Scheduler.
".venv\Scripts\python.exe" distribute_reports.py --config "config.toml"

REM To hide the console window during scheduled runs, use pythonw instead:
REM ".venv\Scripts\pythonw.exe" distribute_reports.py --config "config.toml"

REM Propagate the exit code (0 = all ok, 1 = a failure) to Task Scheduler.
exit /b %ERRORLEVEL%
