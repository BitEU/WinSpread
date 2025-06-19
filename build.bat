REM BUILD SCRIPT - Save as build.bat

@echo off
echo Building Windows Terminal Spreadsheet Calculator...

REM Try to initialize Visual Studio environment
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\VC\Auxiliary\Build\vcvars64.bat" (
    call "C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul 2>nul
)

REM Check for Visual Studio compiler
where cl >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo Using MSVC compiler...
    cl /O2 /W3 /TC main.c /Fe:wtsc.exe /link user32.lib
) else (
    REM Check for MinGW
    where gcc >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        echo Using MinGW compiler...
        gcc -O2 -Wall -std=c99 main.c -o wtsc.exe
    ) else (
        echo Error: No compiler found!
        echo Please install Visual Studio Build Tools or MinGW
        exit /b 1
    )
)

if %ERRORLEVEL% EQU 0 (
    echo Build successful!
    echo Run wtsc.exe to start the spreadsheet
) else (
    echo Build failed!
    exit /b 1
)