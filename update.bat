@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".git" (
  echo Эта папка скачана архивом, а не командой git clone.
  echo Чтобы обновить: скачайте новый архив с GitHub и замените файлы. Папку data не трогайте.
  echo.
  pause
  exit /b 1
)

where git >nul 2>&1
if errorlevel 1 (
  echo Программа Git не найдена. Установите её командой:  winget install Git.Git
  echo Затем закройте это окно и запустите update.bat снова.
  echo.
  pause
  exit /b 1
)

for /f %%i in ('git rev-parse HEAD') do set BEFORE=%%i

echo Проверяю, есть ли новая версия...
git pull --ff-only
if errorlevel 1 (
  echo.
  echo Не удалось скачать обновление. Текст ошибки - выше.
  echo Частая причина: изменён один из файлов приложения. Настройки храните в файле .env.
  echo.
  pause
  exit /b 1
)

for /f %%i in ('git rev-parse HEAD') do set AFTER=%%i

echo.
if "%BEFORE%"=="%AFTER%" echo У вас уже последняя версия.
if not "%BEFORE%"=="%AFTER%" echo Скачана новая версия. Пересобираю сайт...
git log -1 --date=format:%%d.%%m.%%Y "--format=Версия от %%cd: %%s"
echo.

call "%~dp0start.bat"
