@echo off
chcp 65001 >nul
echo =====================================================
echo   CAP NHAT BOT NHAT GO
echo =====================================================
echo.
echo Buoc 1: Rebuild index tu file Word moi...
cd /d "%~dp0"
python rebuild_index.py
if errorlevel 1 (
    echo LOI: Khong the rebuild index!
    pause
    exit /b 1
)
echo.
echo Buoc 2: Dang push len GitHub...
git add ToanBoQuyDinh_NhatGo_2026.docx bm25_index.pkl app.py
git commit -m "Cap nhat quy dinh %date%"
git push
echo.
echo =====================================================
echo   XONG! Bot se tu dong cap nhat trong 3-5 phut.
echo   Kiem tra tai: https://nhat-go-assistant.onrender.com
echo =====================================================
pause
