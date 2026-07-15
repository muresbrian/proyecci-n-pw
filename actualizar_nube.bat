@echo off
title Actualizar Dashboard en la Nube
color 0B

echo ========================================================
echo       ACTUALIZANDO EL DASHBOARD EN STREAMLIT CLOUD
echo ========================================================
echo.

:: Buscar el ejecutable de git
set "GIT_CMD=git"
where git >nul 2>nul
if %ERRORLEVEL% neq 0 (
    if exist "%LOCALAPPDATA%\Programs\Git\cmd\git.exe" set "GIT_CMD=%LOCALAPPDATA%\Programs\Git\cmd\git.exe"
    if exist "%LOCALAPPDATA%\Programs\Git\bin\git.exe" set "GIT_CMD=%LOCALAPPDATA%\Programs\Git\bin\git.exe"
    if exist "C:\Program Files\Git\cmd\git.exe" set "GIT_CMD=C:\Program Files\Git\cmd\git.exe"
    if exist "C:\Program Files\Git\bin\git.exe" set "GIT_CMD=C:\Program Files\Git\bin\git.exe"
)

:: Cambiar al directorio del script
cd /d "%~dp0"

echo Configurando la conexion separada para el nuevo repositorio...
if not exist "Nube_Dashboard" (
    "%GIT_CMD%" clone https://github.com/muresbrian/dashboard-alertas-diarias.git Nube_Dashboard
)

:: 1. Sincronizar la carpeta de la nube ANTES de copiar nuevos archivos (asegura que este limpia)
echo Sincronizando repositorio de la nube con GitHub...
cd Nube_Dashboard
"%GIT_CMD%" pull origin main
cd ..

echo Copiando archivos de codigo y datos...
copy /Y "app.py" "Nube_Dashboard\app.py" >nul
copy /Y "data_processing.py" "Nube_Dashboard\data_processing.py" >nul
copy /Y "export_utils.py" "Nube_Dashboard\export_utils.py" >nul
copy /Y "requirements.txt" "Nube_Dashboard\requirements.txt" >nul
copy /Y "TRX WU_BP.xlsx" "Nube_Dashboard\TRX WU_BP.xlsx" >nul
if not exist "Nube_Dashboard\webapp\Reportes_Individuales_CSV" mkdir "Nube_Dashboard\webapp\Reportes_Individuales_CSV"
copy /Y "webapp\Reportes_Individuales_CSV\Ranking.csv" "Nube_Dashboard\webapp\Reportes_Individuales_CSV\Ranking.csv" >nul

echo Subiendo actualizaciones...
cd Nube_Dashboard
"%GIT_CMD%" add "app.py"
"%GIT_CMD%" add "data_processing.py"
"%GIT_CMD%" add "export_utils.py"
"%GIT_CMD%" add "requirements.txt"
"%GIT_CMD%" add -f "TRX WU_BP.xlsx"
"%GIT_CMD%" add -f "webapp/Reportes_Individuales_CSV/Ranking.csv"
"%GIT_CMD%" commit -m "Actualizacion diaria del Dashboard de Alertas" >nul 2>&1
"%GIT_CMD%" push origin main

echo.
echo ========================================================
echo  [EXITO] La nube se ha actualizado correctamente.
echo  Los cambios se reflejaran en tu Dashboard en segundos.
echo ========================================================
echo.
pause
