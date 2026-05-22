@echo off
cd /d "%~dp0\.."
call "%~dp0python.bat" -c "import sys; sys.path.insert(0,'core'); import converters as c; p=c._find_libreoffice(); print('Found:', p if p else 'NOT FOUND')"
pause
