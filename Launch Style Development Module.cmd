@echo off
setlocal
set "APP_DIR=%~dp0garment_erp"
set "BUNDLED_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%BUNDLED_PYTHON%" (
  start "Style Development Module" cmd /k "cd /d ""%APP_DIR%"" && ""%BUNDLED_PYTHON%"" main.py"
) else (
  start "Style Development Module" cmd /k "cd /d ""%APP_DIR%"" && py -3 main.py"
)

endlocal
