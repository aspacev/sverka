@echo off
chcp 65001 >nul
cd /d "%~dp0"
docker compose down
echo.
echo Сайт остановлен. Данные сохранены в папке data.
timeout /t 5 >nul 2>&1
