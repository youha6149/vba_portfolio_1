chcp 932
@echo off
cd %~dp0
powershell -NoProfile -ExecutionPolicy Unrestricted -File "%USERPROFILE%\MyWorkSpace\sql_server\mydb_precious_metal_table\switch.ps1"

if %ERRORLEVEL% neq 0 (
    echo error occuered error code: %ERRORLEVEL%
    pause
    exit /b %ERRORLEVEL%
)

exit /b 0
