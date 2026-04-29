@echo off
title Capaciteitsplanning - Viewer

cd /d "%~dp0"

echo.
echo  ============================================
echo   Capaciteitsplanning Coatinc Groningen v2.3
echo   Viewer
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

echo  Viewer wordt gestart op http://localhost:8501
echo  Sluit dit venster om de app te stoppen.
echo.
%STREAMLIT_CMD% run viewer_app.py --server.port 8501
if errorlevel 1 (
    echo.
    echo  [FOUT] De viewer kon niet worden gestart.
    pause
)
exit /b
