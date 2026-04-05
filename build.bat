@echo off
setlocal

set PROJECT_DIR=%~dp0

echo.
echo === Step 1: Building Rust engine ===
echo.
cd /d "%PROJECT_DIR%engine"
cargo build --release
if errorlevel 1 (
    echo RUST BUILD FAILED
    pause
    exit /b 1
)
echo Copying engine binary...
copy /Y "target\release\study-agg-engine.exe" "%PROJECT_DIR%"
cd /d "%PROJECT_DIR%"

echo.
echo === Step 2: Building with Coil ===
echo.
coil build . --entry "Study Aggregator.py"
if errorlevel 1 (
    echo COIL BUILD FAILED
    pause
    exit /b 1
)

:: Copy engine binary into dist
copy /Y "study-agg-engine.exe" "dist\Study Aggregator\study-agg-engine.exe"

echo.
echo === Build complete ===
echo Now compile StudyAggSetup.iss in Inno Setup to create the installer.
echo.
pause
