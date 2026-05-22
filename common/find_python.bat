@echo off
set "PY_OK=0"
where python >nul 2>&1 && set "PY_OK=1"
if "%PY_OK%"=="0" where py >nul 2>&1 && set "PY_OK=1"
if "%PY_OK%"=="0" where python3 >nul 2>&1 && set "PY_OK=1"
exit /b 0
