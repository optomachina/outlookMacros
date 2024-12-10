@echo off
echo Installing Prescription Request Add-in...

:: Copy the add-in to Outlook's addins folder
copy "PrescriptionRequestAddin.otm" "%APPDATA%\Microsoft\Outlook\PrescriptionRequestAddin.otm"

:: Enable macros in registry
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v "Level" /t REG_DWORD /d "1" /f

echo Installation complete!
pause 