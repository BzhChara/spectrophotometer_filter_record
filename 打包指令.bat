@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$spec = Get-ChildItem -LiteralPath . -Filter '*.spec' | Select-Object -First 1; if (-not $spec) { Write-Error 'spec not found'; exit 1 }; pyinstaller --clean $spec.FullName"
pause