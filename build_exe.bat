@echo off
setlocal

where py >nul 2>&1
if errorlevel 1 (
  echo No se encontro Python Launcher ^(py^). Instale Python 3.11 o superior.
  exit /b 1
)

if not exist .venv (
  py -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist SistemaFotocopias.spec del /q SistemaFotocopias.spec

pyinstaller --noconfirm --clean --windowed --onefile --name SistemaFotocopias app.py

if errorlevel 1 (
  echo.
  echo ERROR: no se pudo generar el ejecutable.
  exit /b 1
)

echo.
echo Ejecutable generado correctamente:
echo dist\SistemaFotocopias.exe
endlocal
