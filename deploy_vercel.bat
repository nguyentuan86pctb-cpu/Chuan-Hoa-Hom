@echo off
cd /d "%~dp0app"
echo Dang deploy len Vercel tu thu muc app...
echo Neu Vercel yeu cau dang nhap, hay lam theo link/ma hien tren man hinh.
echo Khi hoi link project, hay chon project Vercel cu, khong tao project moi.
echo.
set npm_config_strict_ssl=false
set NODE_TLS_REJECT_UNAUTHORIZED=0
npx vercel link
if errorlevel 1 (
  echo.
  echo Chua link duoc project Vercel. Hay dang nhap/link xong roi chay lai file nay.
  pause
  exit /b 1
)
npx vercel --prod
echo.
pause
