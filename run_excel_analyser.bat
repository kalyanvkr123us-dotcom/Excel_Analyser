@echo off
setlocal

cd /d "%~dp0"

set "INPUT_PATH=%~1"
set "OUTPUT_PATH=%~2"

if "%INPUT_PATH%"=="" set "INPUT_PATH=Source Data\GoCardless Generator Jan JNL.xlsx"

if "%OUTPUT_PATH%"=="" (
    node excel_analyser.js "%INPUT_PATH%"
) else (
    node excel_analyser.js "%INPUT_PATH%" "%OUTPUT_PATH%"
)

echo.
if errorlevel 1 (
    echo The Excel analyser failed.
) else (
    echo The Excel analyser completed successfully.
)
pause