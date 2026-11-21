@echo off
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
if not exist "%~dp0bin" mkdir "%~dp0bin"
cl /nologo /O2 /MT /EHsc /std:c++17 /utf-8 /DUNICODE /D_UNICODE /I"%~dp0mapi-headers" /LD "%~dp0bulk.cpp" /Fo"%~dp0bin\bulk.obj" /Fe"%~dp0bin\ktg-mapi-x64.dll" /link ole32.lib oleaut32.lib
