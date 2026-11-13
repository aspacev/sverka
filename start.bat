@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Запускаю сайт сверок. Первый запуск занимает несколько минут.
echo.

rem Порт берём из файла .env, если он есть; иначе 5000
set SVERKA_PORT=5000
if exist ".env" for /f "usebackq tokens=1,* delims==" %%a in (".env") do if /i "%%a"=="SVERKA_PORT" set SVERKA_PORT=%%b

docker info >nul 2>&1
if errorlevel 1 (
  echo Docker не запущен.
  echo Откройте Docker Desktop, дождитесь надписи "Engine running" и запустите этот файл снова.
  echo.
  pause
  exit /b 1
)

docker compose up -d --build
if errorlevel 1 (
  echo.
  echo Не удалось запустить сайт. Текст ошибки - выше.
  echo.
  pause
  exit /b 1
)

echo.
echo Готово. Сайт открыт по адресу http://localhost:%SVERKA_PORT%
timeout /t 3 /nobreak >nul 2>&1
start "" http://localhost:%SVERKA_PORT%
timeout /t 5 >nul 2>&1
exit /b 0
