@echo off
title Motor de Escoragem
echo Iniciando o app...
"C:\Users\fog\AppData\Local\Python\bin\python.exe" -m streamlit run "%~dp0projeto_leitor_excel\app.py" --server.port 8501
pause
