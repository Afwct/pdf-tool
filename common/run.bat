@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0\.."

set "SUB=%~1"
if "%SUB%"=="" goto badsub
shift /1

call "%~dp0find_python.bat"
if not "%PY_OK%"=="1" goto nopy

echo [1/2] Checking dependencies...
call "%~dp0python.bat" "%~dp0install_deps.py"
if errorlevel 1 goto fail

set "TOOL=%~dp0..\core\pdf_tool.py"
if not exist "%TOOL%" goto nomodule

if /I "%SUB%"=="image2pdf" goto image2pdf
if /I "%SUB%"=="encrypt" goto encrypt
if /I "%SUB%"=="decrypt" goto decrypt
if "%~1"=="" goto usage

echo [2/2] Converting...
if "%~2"=="" (
  call "%~dp0python.bat" "%TOOL%" %SUB% "%~1"
) else (
  call "%~dp0python.bat" "%TOOL%" %SUB% %*
)
goto finish

:image2pdf
if not "%~1"=="" goto image2pdf_files
echo [2/2] Select images...
call "%~dp0python.bat" "%TOOL%" image2pdf --pick
goto finish

:image2pdf_files
echo [2/2] Converting images...
set "ARGS=image2pdf"
:img_next
if "%~1"=="" goto img_do
set "ARGS=!ARGS! "%~1""
shift
goto img_next
:img_do
call "%~dp0python.bat" "%TOOL%" !ARGS!
goto finish

:encrypt
if "%~1"=="" goto usage
set /p PWD=Enter password: 
if "!PWD!"=="" goto usage
call "%~dp0python.bat" "%TOOL%" encrypt "%~1" -p "!PWD!"
goto finish

:decrypt
if "%~1"=="" goto usage
set /p PWD=Enter password: 
call "%~dp0python.bat" "%TOOL%" decrypt "%~1" -p "!PWD!"
goto finish

:usage
echo.
echo ========================================
echo   Tool: %SUB%
echo ========================================
echo.
echo 1) Drag your file onto run.bat in this folder
echo 2) Or: run.bat "C:\path\file.ext"
echo.
echo Run install.bat in the parent folder first.
goto finish

:badsub
echo ERROR: missing tool name.
goto finish

:nopy
echo.
echo Python not found. Run install.bat in parent folder.
goto finish

:nomodule
echo ERROR: core\pdf_tool.py not found.
goto finish

:fail
echo.
echo Dependency check failed. Run install.bat first.
goto finish

:finish
set "EC=!ERRORLEVEL!"
echo.
if "!EC!"=="0" (
  echo ===== SUCCESS =====
) else (
  echo ===== FAILED code !EC! =====
)
echo.
pause
endlocal & exit /b %EC%
