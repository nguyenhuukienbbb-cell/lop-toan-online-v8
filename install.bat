@echo off
chcp 65001 >nul
title Cai dat Lop Toan Online V8.0
echo =============================================
echo        CAI DAT LOP TOAN ONLINE V8.0
echo =============================================
where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo [LOI] Chua tim thay Python.
  echo Hay cai Python 3.12 64-bit va tich "Add python.exe to PATH".
  pause
  exit /b 1
)

python -c "import sys; print('Python:', sys.version.split()[0]); print('64-bit:', sys.maxsize > 2**32)"
python -c "import sys; raise SystemExit(0 if sys.maxsize > 2**32 else 1)"
if errorlevel 1 (
  echo.
  echo [LOI] Ban dang dung Python 32-bit.
  echo Hay cai Python 3.12 64-bit. Ban 32-bit khong cai duoc psycopg binary.
  pause
  exit /b 1
)

echo.
echo Dang cap nhat pip...
python -m pip install --upgrade pip setuptools wheel
if errorlevel 1 goto :fail

echo.
echo Dang cai thu vien...
python -m pip install -r requirements.txt
if errorlevel 1 goto :fail

echo.
echo [OK] Cai dat thanh cong.
echo Bam run.bat de chay phan mem tren may.
pause
exit /b 0

:fail
echo.
echo [LOI] Cai dat chua thanh cong.
echo Hay chup man hinh loi va gui lai de kiem tra.
pause
exit /b 1
