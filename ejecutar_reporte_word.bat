@echo off
title Ejecutar Reporte Word - Diagnostico
color 0E

echo ========================================================
echo        EJECUTANDO GENERAR_REPORTE_WORD.PY
echo ========================================================
echo.

:: Cambiar al directorio donde esta guardado este archivo .bat
cd /d "%~dp0"

:: Validar la ubicacion del archivo Python y ejecutarlo
if exist "generar_reporte_word.py" (
    echo [INFO] Ejecutando desde la carpeta local...
    python generar_reporte_word.py
) else if exist "REPORTES ALETRAS\generar_reporte_word.py" (
    echo [INFO] Ejecutando desde la subcarpeta REPORTES ALETRAS...
    python "REPORTES ALETRAS\generar_reporte_word.py"
) else (
    color 0C
    echo [ERROR] No se encontro el archivo 'generar_reporte_word.py'.
    echo Por favor, asegúrate de colocar este archivo .bat dentro de la carpeta 
    echo principal del proyecto o dentro de la carpeta 'REPORTES ALETRAS'.
)

echo.
echo ========================================================
echo  Proceso terminado. Revisa arriba si hay algun error.
echo  Esta ventana NO se cerrara hasta que presiones una tecla.
echo ========================================================
echo.
pause
