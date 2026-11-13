@echo off
rem Запуск сайта сверок. Окно не закрывайте, пока работаете с сайтом.
cd /d "%~dp0"
start "" http://127.0.0.1:5000
python app.py
pause
