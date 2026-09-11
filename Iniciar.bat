@echo off
REM Abre la interfaz con doble clic, sin usar la terminal.
REM Si existe el entorno virtual lo usa; si no, usa el Python del sistema.

cd /d "%~dp0"

if exist "venv\Scripts\pythonw.exe" (
    start "" "venv\Scripts\pythonw.exe" app.py
    exit /b
)

where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw app.py
    exit /b
)

echo.
echo No se encontro Python en esta computadora.
echo.
echo Instalalo desde https://www.python.org/downloads/
echo y marca la casilla "Add Python to PATH" en la primera pantalla.
echo.
pause