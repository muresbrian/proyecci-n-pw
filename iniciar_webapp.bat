@echo off
title Lanzador del Dashboard de Clientes
color 0B

echo ========================================================
echo       INICIANDO TU WEB APP DE INTELIGENCIA DE NEGOCIO
echo ========================================================
echo.

:: Instalar Streamlit por si no lo tienes
echo Validando librerias de visualizacion (esto puede tomar unos segundos)...
pip install streamlit pandas openpyxl >nul 2>&1

:: Lanzar la aplicacion web
echo.
echo Levantando servidor local... Tu navegador se abrira automaticamente.
echo (Manten esta ventana negra abierta mientras uses el Dashboard)
echo.

streamlit run dashboard.py

pause