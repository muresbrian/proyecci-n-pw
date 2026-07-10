@echo off
title Automatizacion de Analisis de Clientes
color 0A

echo ========================================================
echo       INICIANDO ANALISIS DE TENDENCIAS EN PYTHON
echo ========================================================
echo.

:: 1. Validar que Python este disponible
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    color 0C
    echo [ERROR] Python no esta instalado.
    echo.
    pause
    exit /b
)

:: 2. Instalar librerias necesarias (pandas y openpyxl para leer excel)
echo Verificando librerias necesarias...
pip install pandas openpyxl >nul 2>&1

:: 3. Validar que los archivos existan
if not exist "analizar_tendencias.py" (
    color 0E
    echo [ADVERTENCIA] No se encontro el archivo 'analizar_tendencias.py'.
    echo.
    pause
    exit /b
)

if not exist "TRX WU_BP.xlsx" (
    color 0E
    echo [ADVERTENCIA] No se encontro el archivo Excel 'TRX WU_BP.xlsx'.
    echo.
    pause
    exit /b
)

:: 4. Ejecutar el script
echo.
echo Ejecutando el modelo de analisis...
echo.
python analizar_tendencias.py
set "ANALYSIS_ERROR=%ERRORLEVEL%"

:: Buscar el ejecutable de git fuera de bloques complejos para evitar errores de expansion de variables
set "GIT_CMD=git"
where git >nul 2>nul
if %ERRORLEVEL% neq 0 (
    if exist "%LOCALAPPDATA%\Programs\Git\cmd\git.exe" set "GIT_CMD=%LOCALAPPDATA%\Programs\Git\cmd\git.exe"
    if exist "%LOCALAPPDATA%\Programs\Git\bin\git.exe" set "GIT_CMD=%LOCALAPPDATA%\Programs\Git\bin\git.exe"
    if exist "C:\Program Files\Git\cmd\git.exe" set "GIT_CMD=C:\Program Files\Git\cmd\git.exe"
    if exist "C:\Program Files\Git\bin\git.exe" set "GIT_CMD=C:\Program Files\Git\bin\git.exe"
)

:: 5. Validar resultado
if %ANALYSIS_ERROR% equ 0 (
    echo.
    echo ========================================================
    echo  [EXITO] El analisis ha finalizado correctamente.
    echo  Se genero: Analisis_Historico_Variaciones.xlsx
    echo ========================================================
    
    echo.
    echo Generando reporte de clientes perdidos...
    echo.
    python generar_reporte_perdidos.py
    
    echo.
    echo Generando reporte de Health Score de la cartera 2026...
    echo.
    python generar_reporte_salud_2026.py
    
    echo.
    echo Generando reporte de Health Score Avanzado de la cartera 2026...
    echo.
    python generar_reporte_salud_avanzado.py
    
    echo.
    echo Separando las hojas del archivo Excel en archivos CSV...
    echo.
    python separar_hojas_csv.py

    :: Sincronizar datos con GitHub
    echo.
    echo Sincronizando reportes con tu repositorio de GitHub...
    
    "%GIT_CMD%" add webapp/Reportes_Individuales_CSV/*.csv
    "%GIT_CMD%" commit -m "Actualizacion automatica de reportes" >nul 2>&1
    "%GIT_CMD%" push origin main
    echo Sincronizacion completada con exito.
) else (
    color 0C
    echo.
    echo ========================================================
    echo  [ERROR] Ocurrio un problema en Python.
    echo ========================================================
)

echo.
pause