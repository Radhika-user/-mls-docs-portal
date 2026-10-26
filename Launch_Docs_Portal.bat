@echo off
title MLS & Research Operations - Documentation Portal
color 1F

echo ==============================================================================
echo             MLS & RESEARCH OPERATIONS - DOCUMENTATION PORTAL
echo ==============================================================================
echo.
echo   [Network Share Source]
echo   \\pfbafileserver\File Server\Admin Ops\Research\MLS\Procedures
echo.
echo   [Access URLs for your team]
echo     - Local (Your PC):   http://localhost:8501
echo     - Network (Host):    http://PFBA-LAP-PT-032:8501
echo     - Network (IP):      http://192.168.10.38:8501
echo.
echo   [Authentication]
echo     - Domain: OCNHOU (Windows Active Directory)
echo.
echo ==============================================================================
echo   Launching Portal and opening browser...
echo   (Keep this window open while you and your team are using the portal)
echo ==============================================================================
echo.

"C:\Users\karthik.padmanathan\AppData\Local\Programs\Python\Python312\python.exe" -m streamlit run "C:\Users\karthik.padmanathan\mls_docs_portal\app.py" --server.port 8501 --server.headless false

echo.
echo Portal has stopped.
pause
