@echo off
cd /d "%~dp0"
call "%~dp0..\common\run.bat" decrypt %*
