@echo off
REM Quick launcher for C# Test Generator
REM Usage: generate-tests.bat [project-path] [--file filename.cs]

echo.
echo ===================================
echo  C# Automated Test Generator
echo ===================================
echo.

if "%1"=="" (
    echo Usage: generate-tests.bat [project-path] [--file filename.cs]
    echo.
    echo Examples:
    echo   generate-tests.bat FindMissingAppointments
    echo   generate-tests.bat FindMissingAppointments --file Helper.cs
    echo   generate-tests.bat .
    echo.
    exit /b 1
)

node scripts\test-generator\generate-tests.js %*

echo.
echo ===================================
echo  Generation Complete!
echo ===================================
