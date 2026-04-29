@echo off
title Capaciteitsplanning - Beheer

cd /d "%~dp0"

echo.
echo  ============================================
echo   Capaciteitsplanning Coatinc Groningen v2.3
echo   Beheerapplicatie
echo  ============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo  [FOUT] Python niet gevonden in PATH.
    pause
    exit /b
)

echo  Controleren en installeren van benodigde packages...
python -m pip install -q --prefer-binary streamlit pandas numpy openpyxl matplotlib
python -m pip install -q --prefer-binary --force-reinstall "httpcore>=1.0.8" "httpx==0.27.2"
python -m pip install -q --prefer-binary --no-deps supabase storage3 gotrue postgrest realtime supafunc
if errorlevel 1 (
    echo.
    echo  [FOUT] Installatie mislukt. Zie foutmelding hierboven.
    pause
    exit /b
)
echo  Packages gereed.
echo.

set STREAMLIT_CMD=streamlit
where streamlit >nul 2>&1
if errorlevel 1 (
    set STREAMLIT_CMD=python -m streamlit
)

echo  Beheerapplicatie wordt gestart op http://localhost:8502
echo  Sluit dit venster om de app te stoppen.
echo.
%STREAMLIT_CMD% run manager_app.py --server.port 8502
if errorlevel 1 (
    echo.
    echo  [FOUT] De beheerapplicatie kon niet worden gestart.
    pause
)
exit /b
