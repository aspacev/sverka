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

echo Скачиваю новую версию...
git pull
if errorlevel 1 (
  echo.
  echo Не удалось скачать обновление. Текст ошибки - выше.
  echo.
  pause
  exit /b 1
)

echo.
call "%~dp0start.bat"
