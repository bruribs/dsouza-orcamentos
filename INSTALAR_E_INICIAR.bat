@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==============================================
echo   DSouza Orcamentos - Instalacao local
echo ==============================================
echo.
where python >nul 2>nul
if errorlevel 1 (
  echo Python nao foi encontrado neste computador.
  echo Instale o Python e marque a opcao "Add Python to PATH".
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
  echo Criando ambiente do sistema...
  python -m venv .venv
)
echo Instalando/atualizando componentes...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Houve um erro na instalacao.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" inicializar.py
start "" http://127.0.0.1:5000
".venv\Scripts\python.exe" app.py
pause
