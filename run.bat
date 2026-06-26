@echo off
cd /d "%~dp0app"
echo Dang mo web app tai dia chi: http://localhost:5173
echo Neu trinh duyet khong tu mo, hay copy dia chi tren de mo.
start http://localhost:5173
py -m http.server 5173
if errorlevel 1 (
  python -m http.server 5173
)
pause
