@echo off
echo ============================================================
echo  Generating Public Internet Link for DIFFERENT Wi-Fi / Mobile Data
echo ============================================================
echo.
ssh -o StrictHostKeyChecking=no -R 80:127.0.0.1:5000 serveo.net
pause
