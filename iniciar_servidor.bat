@echo off
title Servidor de Proyección
color 0B

echo ========================================================
echo         INICIANDO TU WEB APP DE PROYECCIÓN WUZI
echo ========================================================
echo.

echo Levantando servidor local en el puerto 8080...
echo (Manten esta ventana abierta mientras uses el Dashboard)
echo.

:: Abrir el navegador
start http://localhost:8080/

:: Cambiar directorio a webapp y lanzar servidor estatico
cd webapp
python -m http.server 8080

pause
