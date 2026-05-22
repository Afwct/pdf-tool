@echo off
REM Run python: prefer "python", else "py -3"
where python >nul 2>&1
if %ERRORLEVEL%==0 (
  python %*
  exit /b %ERRORLEVEL%
)
where py >nul 2>&1
if %ERRORLEVEL%==0 (
  py -3 %*
  exit /b %ERRORLEVEL%
)
echo ERROR: Python not found.
exit /b 1
