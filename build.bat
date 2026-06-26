@echo off
cd /d "%~dp0"
if not exist dist mkdir dist
xcopy app dist\app /E /I /Y
copy run.bat dist\run.bat
copy README.md dist\README.md
echo.
echo Da copy ban chay tinh vao thu muc dist.
pause
