Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ""C:\Users\karthik.padmanathan\mls_docs_portal\auto_sync.ps1""", 0, False
