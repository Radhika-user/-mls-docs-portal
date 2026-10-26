@echo off
title Sync Procedures & Push to GitHub
color 0A

echo ==============================================================================
echo             MLS & RESEARCH PORTAL - SYNC & PUSH TO GITHUB
echo ==============================================================================
echo.

echo [1/3] Checking for new documents in File Server...
set "SRC=\\pfbafileserver\File Server\Admin Ops\Research\MLS\Procedures"
set "DEST=C:\Users\karthik.padmanathan\mls_docs_portal\sample_docs"

if exist "%SRC%" (
    echo Syncing latest files from network share to cloud package...
    robocopy "%SRC%" "%DEST%" /MIR /FFT /Z /NP /NDL /NJH /NJS /XF ~$* ~* *.tmp
    echo Network share sync complete.
) else (
    echo Note: Network share not reachable right now, pushing local files.
)

echo.
echo [2/3] Adding changes to Git...
cd /d "C:\Users\karthik.padmanathan\mls_docs_portal"
"C:\Users\karthik.padmanathan\AppData\Local\Programs\Git\cmd\git.exe" add .
"C:\Users\karthik.padmanathan\AppData\Local\Programs\Git\cmd\git.exe" commit -m "Auto-sync latest documents and updates"

echo.
echo [3/3] Uploading to GitHub & Streamlit Cloud...
"C:\Users\karthik.padmanathan\AppData\Local\Programs\Git\cmd\git.exe" push origin main

echo.
echo ==============================================================================
echo Done! New documents have been synced and pushed to the Cloud app.
echo ==============================================================================
echo.
pause