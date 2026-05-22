@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0\.."

call "%~dp0..\common\find_python.bat"
if not "%PY_OK%"=="1" (
  echo Run install.bat in the parent folder first.
  pause
  exit /b 1
)

call "%~dp0..\common\python.bat" "%~dp0..\common\install_deps.py"
if errorlevel 1 (
  pause
  exit /b 1
)

set "TOOL=%~dp0..\core\pdf_tool.py"

if "%~1"=="" (
  echo No files passed. Opening image picker...
  call "%~dp0..\common\python.bat" "%TOOL%" image2pdf --pick
  goto done
)

set "ARGS=image2pdf"
:loop
if "%~1"=="" goto run
set "ARGS=!ARGS! "%~1""
shift
goto loop

:run
echo Converting images to PDF...
call "%~dp0..\common\python.bat" "%TOOL%" !ARGS!

:done
echo.
if errorlevel 1 (echo FAILED) else (echo OK)
pause
exit /b %ERRORLEVEL%
