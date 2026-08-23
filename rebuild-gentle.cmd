@echo off
REM Rebuild + redeploy the helper pinned to the E-cores at below-normal priority.
cd /d C:\TabsPortalHelper
start /B /WAIT /AFFINITY FFFF0000 /BELOWNORMAL "" powershell -NoProfile -ExecutionPolicy Bypass -File C:\TabsPortalHelper\Rebuild.ps1 > "%TEMP%\helper-rebuild.txt" 2>&1
set RC=%errorlevel%
findstr /C:"error" /C:"Error" /C:"Done." /C:"Copied" /C:"failed" "%TEMP%\helper-rebuild.txt"
echo rebuild exit %RC%
exit /b %RC%
