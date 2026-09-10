@echo off
REM ============================================================
REM  Genera el ejecutable para Windows.
REM  Se corre UNA vez, con doble clic, en una maquina Windows.
REM  Al terminar queda "dist\Descarga de fotos.exe".
REM ============================================================

cd /d "%~dp0"
echo.
echo   Generando el ejecutable. Tarda unos minutos.
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo   No se encontro Python.
    echo   Instalalo desde https://www.python.org/downloads/
    echo   y marca "Add Python to PATH".
    pause
    exit /b 1
)

python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo.
    echo   Fallo la instalacion de dependencias.
    echo   Si el error menciona curl_cffi y esta maquina es ARM, ver el README.
    pause
    exit /b 1
)

REM --collect-all curl_cffi: se lleva el paquete ENTERO. Hace falta porque
REM   curl_cffi incluye una libcurl compilada que PyInstaller no detecta solo,
REM   y sin ella Lanidor y Michael Kors no funcionan.
REM --hidden-import: los modulos de marca se importan tarde (dentro de una
REM   funcion), asi que se nombran explicitamente para que no queden afuera.

python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "Descarga de fotos" ^
    --collect-all curl_cffi ^
    --hidden-import runner --hidden-import comun ^
    --hidden-import geox --hidden-import lanidor ^
    --hidden-import hugo --hidden-import mk ^
    app.py

if errorlevel 1 (
    echo.
    echo   Fallo la generacion. Revisar el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo   ============================================
echo     Listo: dist\Descarga de fotos.exe
echo   ============================================
echo.
echo   Ese archivo solo se copia y se abre con doble clic.
echo   No necesita Python ni nada instalado.
echo.
pause
