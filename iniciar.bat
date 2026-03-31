@echo off
title Descarga Automática Certificados Tope
echo.
echo  =====================================================
echo   Descarga Automatica Certificados Tope - Activa IT
echo   Grupo Campbell
echo  =====================================================
echo.
echo  Verificando dependencias...
pip show playwright >nul 2>&1 || pip install playwright
playwright install chromium >nul 2>&1
echo.
echo  Iniciando servidor local...
echo  Abre tu navegador en: http://127.0.0.1:5001
echo.
echo  Para cerrar presiona Ctrl+C
echo.
python app.py
pause
