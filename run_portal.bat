@echo off
title MLS Procedures Documentation Portal
echo ========================================================
echo   MLS & Research Operations - Documentation UI Portal
echo ========================================================
echo.
echo Launching documentation portal in your default browser...
"C:\Users\karthik.padmanathan\AppData\Local\Programs\Python\Python312\python.exe" -m streamlit run "C:\Users\karthik.padmanathan\mls_docs_portal\app.py" --server.port 8501 --server.headless false
pause
