@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

REM ============================================================
REM  actualizar_datos.bat
REM  --------------------------------------------------------
REM  Regenera data/apoyos.json a partir de data/apoyos.xlsx
REM  con un doble clic. No hay que abrir la terminal.
REM
REM  ESTRUCTURA DE CARPETAS ESPERADA (este .bat va en "scripts"):
REM    proyecto/
REM     |- data/      apoyos.xlsx y apoyos.json
REM     |- scripts/   excel_to_json.py y este actualizar_datos.bat
REM     |- source/    portal_v2.html y la carpeta multimedia/
REM
REM  REQUISITO (solo la primera vez): tener Python instalado.
REM  Si no lo tienes, descargalo de https://www.python.org/downloads/
REM  y durante la instalacion marca la casilla
REM  "Add Python to PATH".
REM ============================================================

cd /d "%~dp0"

echo ============================================================
echo   Actualizando datos del portal Cuenta Conmigo - Mas Chihuahua
echo ============================================================
echo.

REM --- Buscar Python (puede llamarse "python" o "py" segun la instalacion) ---
set PYCMD=
where python >nul 2>nul
if %errorlevel%==0 (
    set PYCMD=python
) else (
    where py >nul 2>nul
    if %errorlevel%==0 (
        set PYCMD=py
    )
)

if "%PYCMD%"=="" (
    echo [ERROR] No se encontro Python instalado en este equipo.
    echo.
    echo Para solucionarlo:
    echo   1. Ve a https://www.python.org/downloads/
    echo   2. Descarga e instala la version mas reciente.
    echo   3. Durante la instalacion, marca la casilla que dice
    echo      "Add Python to PATH" antes de darle a Instalar.
    echo   4. Vuelve a abrir este archivo .bat.
    echo.
    pause
    exit /b 1
)

echo Usando: %PYCMD%
echo.

REM --- Verificar/instalar la libreria openpyxl (necesaria para leer el Excel) ---
%PYCMD% -c "import openpyxl" >nul 2>nul
if not %errorlevel%==0 (
    echo La libreria "openpyxl" no esta instalada todavia.
    echo Instalandola automaticamente, espera un momento...
    %PYCMD% -m pip install openpyxl --quiet
    if not %errorlevel%==0 (
        echo.
        echo [ERROR] No se pudo instalar openpyxl automaticamente.
        echo Intenta manualmente abriendo una terminal aqui y escribiendo:
        echo     %PYCMD% -m pip install openpyxl
        echo.
        pause
        exit /b 1
    )
    echo Listo, openpyxl quedo instalado.
    echo.
)

REM --- Verificar que exista el Excel antes de intentar convertirlo ---
if not exist "..\data\apoyos.xlsx" (
    echo [ERROR] No se encontro el archivo apoyos.xlsx en:
    echo     %cd%\..\data\apoyos.xlsx
    echo.
    echo Verifica que la carpeta "data" exista junto a la carpeta "scripts"
    echo (deben ser carpetas hermanas, ambas dentro de la misma carpeta
    echo principal del proyecto) y que dentro de "data" este apoyos.xlsx.
    echo.
    pause
    exit /b 1
)

REM --- Correr la conversion ---
echo Leyendo ..\data\apoyos.xlsx y generando ..\data\apoyos.json...
echo --------------------------------------------------------------
%PYCMD% excel_to_json.py
echo --------------------------------------------------------------

if %errorlevel%==0 (
    echo.
    echo ============================================================
    echo   LISTO. data\apoyos.json se actualizo correctamente.
    echo   Solo recarga la pagina del portal en el navegador
    echo   (Ctrl+F5 para forzar la recarga) y ya veras los cambios.
    echo ============================================================
) else (
    echo.
    echo ============================================================
    echo   Ocurrio un error al generar el JSON.
    echo   Revisa el mensaje de arriba para ver el detalle.
    echo   Causas comunes: una columna del Excel fue renombrada o
    echo   borrada, o el archivo esta abierto en Excel en este momento
    echo   (cierralo y vuelve a intentar).
    echo ============================================================
)

echo.
pause

