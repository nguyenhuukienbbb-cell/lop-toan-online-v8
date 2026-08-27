@echo off
chcp 65001 >nul
title Lop Toan Online V8.0
where python >nul 2>nul
if errorlevel 1 (
 echo [LOI] Chua tim thay Python. Hay chay install.bat truoc.
 pause
 exit /b 1
)
set APP_TIMEZONE=Asia/Ho_Chi_Minh
set COOKIE_SECURE=0
start "" http://127.0.0.1:5000
python app.py
pause
