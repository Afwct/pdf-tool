@echo off
cd /d "%~dp0"
call "%~dp0..\common\run.bat" ppt2pdf %*
