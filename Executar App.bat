@echo off
cd /d "%~dp0"
echo Iniciando Gerador de Atas...
echo.
python -m streamlit run app.py
pause
