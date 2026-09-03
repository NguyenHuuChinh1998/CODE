@echo off
set PYTHON=C:\Users\huuchinh.nguyen\AppData\Local\anaconda3\python.exe
set STREAMLIT=C:\Users\huuchinh.nguyen\AppData\Local\anaconda3\Scripts\streamlit.exe
set DASHBOARD=C:\Users\huuchinh.nguyen\Concentrix Corporation\WFM-Expedia-HCM - Branding files\BI_Task\CODE\Python_Code\dashboard

cd /d "%DASHBOARD%"

:: Kiểm tra streamlit tồn tại không
if not exist "%STREAMLIT%" (
    echo [LOI] Khong tim thay streamlit tai: %STREAMLIT%
    echo Vui long chay: pip install streamlit
    pause
    exit /b 1
)

:: Chạy Streamlit trong background
echo Dang khoi dong WFM Dashboard...
start /b "" "%STREAMLIT%" run app.py --server.headless true --server.port 8501

:: Chờ Streamlit sẵn sàng
timeout /t 5 /nobreak >nul

:: Mở Edge
start msedge http://localhost:8501

echo.
echo ✓ Dashboard dang chay tai: http://localhost:8501
echo   Dong cua so nay se TAT dashboard.
echo.
pause >nul
