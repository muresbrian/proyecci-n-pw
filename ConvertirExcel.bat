@echo off
color 0A
echo ===================================================
echo Convirtiendo archivo de texto a Excel...
echo Por favor espera unos segundos.
echo ===================================================

:: Definimos las rutas de los archivos de forma dinámica
for %%i in ("%~dp0..") do set "PARENT_DIR=%%~fi"
set "TXT_FILE=%PARENT_DIR%\PRUEBa hob.txt"
set "EXCEL_FILE=%PARENT_DIR%\PRUEBa hob.xlsx"

:: Usamos PowerShell de fondo para abrir el txt con Excel y guardarlo como .xlsx
powershell -NoProfile -Command "$excel = New-Object -ComObject Excel.Application; $excel.DisplayAlerts = $false; $excel.Visible = $false; $wb = $excel.Workbooks.Open('%TXT_FILE%'); $wb.SaveAs('%EXCEL_FILE%', 51); $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null"

echo.
echo ===================================================
echo !Conversion terminada exitosamente! 
echo Tu archivo de Excel esta listo en la carpeta de descargas.
echo ===================================================
echo.
pause