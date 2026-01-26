@echo off
setlocal EnableExtensions

REM === USTAWIENIA ===
set "DEST=C:\ahk"
set "HERE=%~dp0"

REM 1) znajdz ZIP AutoHotkey_* w folderze BAT
set "ZIP="
for %%F in ("%HERE%AutoHotkey_*.zip") do (
    set "ZIP=%%F"
    goto :zipfound
)

:zipfound
if not defined ZIP (
    echo Nie znaleziono pliku AutoHotkey_*.zip w:
    echo %HERE%
    pause
    exit /b 1
)

echo ZIP: "%ZIP%"

REM 2) utworz folder docelowy
if not exist "%DEST%" mkdir "%DEST%" 2>nul
if not exist "%DEST%" (
    echo Nie moge utworzyc "%DEST%".
    echo Uruchom jako administrator albo zmien DEST na %%USERPROFILE%%\ahk
    pause
    exit /b 1
)

REM 3) rozpakuj ZIP
echo Rozpakowywanie...
powershell -NoProfile -ExecutionPolicy Bypass ^
  "Expand-Archive -LiteralPath '%ZIP%' -DestinationPath '%DEST%' -Force" >nul 2>&1

REM 4) znajdz AutoHotkey64.exe (moze byc w podfolderze)
set "AHKEXE="
for /r "%DEST%" %%E in (AutoHotkey64.exe) do (
    set "AHKEXE=%%E"
    goto :exefound
)

:exefound
if not defined AHKEXE (
    echo Nie znaleziono AutoHotkey64.exe w "%DEST%"
    pause
    exit /b 1
)

echo EXE: "%AHKEXE%"

REM 5) ustaw asocjacje (bez admina - HKCU)
reg add "HKCU\Software\Classes\.ahk" /ve /d "AutoHotkeyScript" /f >nul
reg add "HKCU\Software\Classes\AutoHotkeyScript\shell\open\command" /ve /d "\"%AHKEXE%\" \"%%1\"" /f >nul

echo.
echo GOTOWE.
echo Rozpakowano do: %DEST%
echo .ahk powiazane z: %AHKEXE%
echo.
pause
