@echo off
setlocal

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

set "VENV_PYTHON=%~dp0.venv\Scripts\python.exe"
set "SPEC_FILE=%~dp0RelatorioFotografico.spec"

if not exist "%VENV_PYTHON%" (
  echo Python da virtualenv nao encontrado em "%VENV_PYTHON%".
  pause
  exit /b 1
)

"%VENV_PYTHON%" -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
  echo PyInstaller nao esta instalado na virtualenv.
  echo Instale com: .venv\Scripts\python.exe -m pip install pyinstaller
  pause
  exit /b 1
)

if not exist "%SPEC_FILE%" (
  echo Arquivo de build nao encontrado em "%SPEC_FILE%".
  pause
  exit /b 1
)

"%VENV_PYTHON%" -m PyInstaller --noconfirm --clean "%SPEC_FILE%"
if errorlevel 1 (
  echo Falha ao gerar o executavel.
  pause
  exit /b 1
)

set "RELATORIO_SMOKE_TEST=1"
"%~dp0dist\RelatorioFotografico.exe"
if errorlevel 1 (
  echo Smoke test do executavel falhou.
  pause
  exit /b 1
)
set "RELATORIO_SMOKE_TEST="

echo Build concluido e smoke test aprovado.

pause
