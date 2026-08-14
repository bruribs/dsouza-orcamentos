@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo O sistema ainda nao foi instalado.
  echo Execute primeiro: INSTALAR_E_INICIAR.bat
  pause
  exit /b 1
)
start "" http://127.0.0.1:5000
".venv\Scripts\python.exe" app.py
pause
