@echo off
setlocal
set "APP_DIR=%~dp0"
set "BUNDLED_NODE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
if exist "%BUNDLED_NODE%" (
  start "Piece Rate Server" cmd /k "cd /d ""%APP_DIR%"" && ""%BUNDLED_NODE%"" server.js"
) else (
  start "Piece Rate Server" cmd /k "cd /d ""%APP_DIR%"" && node server.js"
)
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:3100"
endlocal
