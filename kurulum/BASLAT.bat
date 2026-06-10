@echo off
cd /d "%~dp0"
echo ============================================
echo  PERSONEL ICRA TAKIP - KURULUM BASLATILIYOR
echo ============================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Kur.ps1"
echo.
echo Pencereyi kapatabilirsiniz.
pause
