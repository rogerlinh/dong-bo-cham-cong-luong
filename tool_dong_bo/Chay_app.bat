@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Dong bo cham cong luong - dong cua so nay de tat

set PY=python
where python >nul 2>nul || set PY=py
%PY% --version >nul 2>nul
if errorlevel 1 (
  echo Chua cai Python. Tai tai https://www.python.org/downloads/ va tick "Add python.exe to PATH" khi cai.
  pause
  exit /b 1
)

%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Dang cai thu vien openpyxl lan dau...
  %PY% -m pip install --user -r ..\requirements.txt
)

%PY% app_web.py
if errorlevel 1 pause
