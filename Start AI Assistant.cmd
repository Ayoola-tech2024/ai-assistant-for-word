@echo off
rem AI Assistant for Word — launcher. Double-click this, nothing else needed.
rem 1) Open your Word document first (if it's not open yet, open it now).
rem 2) This opens the assistant window next to Word.
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if errorlevel 1 (
    echo ============================================================
    echo  Please open your Word document first, then press any key
    echo  to open the assistant. The assistant will wait for Word.
    echo ============================================================
    pause >nul
)
start "" "%~dp0AIAssistantForWord\AIAssistantForWord.exe"
