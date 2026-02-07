@echo off
setlocal

set PROJECT_DIR=%~dp0
set BUILD_DIR=%PROJECT_DIR%build\exe.win-amd64-3.13
set PFX_PATH=C:\Certs\RonsinCodeSign.pfx
set SIGNTOOL="C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x64\signtool.exe"

:: Prompt for certificate password
set /p PFX_PASS="Enter code signing certificate password: "

echo.
echo === Step 1: Building with cx_Freeze ===
echo.
cd /d "%PROJECT_DIR%"
python setup.py build
if errorlevel 1 (
    echo BUILD FAILED
    pause
    exit /b 1
)

echo.
echo === Step 2: Signing executables ===
echo.
for %%F in ("%BUILD_DIR%\Study Aggregator.exe" "%BUILD_DIR%\update_checker.exe" "%BUILD_DIR%\reg.exe" "%BUILD_DIR%\unreg.exe") do (
    echo Signing %%F ...
    %SIGNTOOL% sign /f "%PFX_PATH%" /p "%PFX_PASS%" /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 /d "Study Aggregator" %%F
    if errorlevel 1 (
        echo SIGNING FAILED for %%F
        pause
        exit /b 1
    )
)

echo.
echo === Build and signing complete ===
echo Now compile StudyAggSetup.iss in Inno Setup to create the installer.
echo.
pause
