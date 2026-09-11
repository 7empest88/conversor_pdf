@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo No se encontro el entorno virtual .venv.
    echo Ejecuta primero instalar_windows.ps1.
    echo.
    pause
    exit /b 1
)

echo.
echo Iniciando la aplicacion local...
echo Abre esta direccion en tu navegador:
echo http://127.0.0.1:5000
echo.

".venv\Scripts\python.exe" -m waitress --host=127.0.0.1 --port=5000 app_conversor_documentos_local:app

pause
