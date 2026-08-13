@echo off
echo Creating Public Internet Link for DIFFERENT Wi-Fi / Mobile Data...
echo.
ssh -o StrictHostKeyChecking=no -R 80:localhost:5000 nokey@localhost.run
pause

