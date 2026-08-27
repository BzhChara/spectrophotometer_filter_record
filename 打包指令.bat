@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$spec = Get-ChildItem -LiteralPath . -Filter '*.spec' | Select-Object -First 1; if (-not $spec) { Write-Error 'spec not found'; exit 1 }; python -m PyInstaller --clean $spec.FullName"
if errorlevel 1 goto :build_failed

if not exist "%~dp0dist\config.ini" (
    copy /y "%~dp0config.ini" "%~dp0dist\config.ini" >nul
    if errorlevel 1 goto :config_failed
    echo [INFO] Created dist\config.ini from the project default config.
) else (
    echo [INFO] Kept the existing dist\config.ini.
)

pause
exit /b 0

:build_failed
echo [ERROR] Packaging failed. dist\config.ini was not changed.
pause
exit /b 1

:config_failed
echo [ERROR] Packaging succeeded, but dist\config.ini could not be created.
pause
exit /b 1
