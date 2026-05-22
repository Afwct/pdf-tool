@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo   PDF Toolkit - Install
echo ========================================
echo.

call "%~dp0common\find_python.bat"
if not "%PY_OK%"=="1" goto nopy

echo Python version:
call "%~dp0common\python.bat" --version
if errorlevel 1 goto fail
echo.

call "%~dp0common\python.bat" "%~dp0common\install_deps.py"
if errorlevel 1 goto fail

echo.
echo Done. Open a tool folder and double-click run.bat
echo.
pause
exit /b 0

:nopy
echo.
echo Python not found. Install from https://www.python.org/
echo Check "Add python to PATH" during setup.
echo.
pause
exit /b 1

:fail
echo.
echo Install failed.
echo.
pause
exit /b 1
