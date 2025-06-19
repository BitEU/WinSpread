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
    echo Compiling resource file...
    rc resource.rc
    if %ERRORLEVEL% EQU 0 (        echo Compiling and linking with icon...
        cl /O2 /W3 /TC main.c /Fe:wtsc.exe /link resource.res user32.lib
    ) else (
        echo Warning: Resource compilation failed, building without icon...
        cl /O2 /W3 /TC main.c /Fe:wtsc.exe /link user32.lib
    )
) else (
    echo Error: Visual Studio compiler not found!
    echo Please install Visual Studio Build Tools
    exit /b 1
)

if %ERRORLEVEL% EQU 0 (
    echo Build successful!
    echo Run wtsc.exe to start the spreadsheet
    REM Clean up temporary resource files
    if exist resource.res del resource.res >nul 2>nul
) else (
    echo Build failed!
    exit /b 1
)