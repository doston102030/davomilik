@echo off
setlocal
cd /d "%~dp0"
title SVOD TIZIMI

echo ==========================================
echo          SVOD TIZIMI ishga tushmoqda
echo ==========================================
echo.

set "VENV_PY=.venv\Scripts\python.exe"

rem Lokal muhit bor va ishlayapti - Python qidirish shart emas.
if exist "%VENV_PY%" (
  "%VENV_PY%" -c "import sys" >nul 2>&1
  if not errorlevel 1 goto :deps
  echo Lokal muhit buzilgan ^(boshqa kompyuterdan ko'chirilgan^). Qayta yaratiladi...
  rmdir /s /q ".venv"
)

rem Python 3.9+ qidirish. Microsoft Store "python.exe" niqobi ishlamaydi, shuning uchun
rem mavjudlik emas, haqiqiy ishga tushishi tekshiriladi.
set "PYEXE="
py -3 -c "import sys; sys.exit(sys.version_info < (3, 9))" >nul 2>&1
if not errorlevel 1 set "PYEXE=py -3"
if defined PYEXE goto :mkvenv
python -c "import sys; sys.exit(sys.version_info < (3, 9))" >nul 2>&1
if not errorlevel 1 set "PYEXE=python"
if defined PYEXE goto :mkvenv

echo XATO: Python 3.9 yoki undan yangi versiya topilmadi.
echo Python 3 ni kompyuterga o'rnating va qayta ishga tushiring.
echo O'rnatishda "Add python.exe to PATH" belgisini qo'ying.
echo Rasmiy manba: https://www.python.org/downloads/windows/
echo.
pause
exit /b 1

:mkvenv
echo Birinchi ishga tushirish: lokal muhit yaratilmoqda...
%PYEXE% -m venv .venv
if errorlevel 1 goto :fail

:deps
"%VENV_PY%" -c "import flask, xlsxwriter" >nul 2>&1
if not errorlevel 1 goto :run
echo Kerakli modullar o'rnatilmoqda...
"%VENV_PY%" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto :fail

:run
echo Tizim tayyor. Brauzer avtomatik ochiladi.
echo Ushbu qora oynani dastur ishlayotgan paytda yopmang.
echo.
"%VENV_PY%" app.py
if errorlevel 1 (
  echo.
  echo Dastur xato bilan to'xtadi. Tafsilotlar: logs\app.log
  pause
  exit /b 1
)
goto :eof

:fail
echo.
echo XATO: tizimni ishga tushirib bo'lmadi.
echo Internet yoki Python sozlamalarini tekshiring.
pause
exit /b 1
