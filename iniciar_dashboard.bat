@echo off
echo Iniciando Dashboard de Alertas...
cd /d "%~dp0"
streamlit run app.py
pause
