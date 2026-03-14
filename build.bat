@echo off
setlocal

:: Setup MSVC 2022 Build Tools environment
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul 2>&1

set OUTDIR=%~dp0dist
if not exist "%OUTDIR%" mkdir "%OUTDIR%"

echo [1/3] Compiling context_menu.dll ...
cl /nologo /EHsc /std:c++17 /O2 /LD ^
   /DWIN32_LEAN_AND_MEAN /DUNICODE /D_UNICODE ^
   /Fe:"%OUTDIR%\context_menu.dll" ^
   /Fo:"%OUTDIR%\\" ^
   "%~dp0src\context_menu.cpp" ^
   /link /DEF:"%~dp0src\context_menu.def" ^
   /OPT:REF /OPT:ICF ^
   shlwapi.lib shell32.lib runtimeobject.lib advapi32.lib ole32.lib onecore.lib

if errorlevel 1 (
    echo [FAILED] DLL compilation failed.
    exit /b 1
)

echo [2/3] Compiling stub.exe ...
cl /nologo /O2 ^
   /Fe:"%OUTDIR%\stub.exe" ^
   /Fo:"%OUTDIR%\\" ^
   "%~dp0src\stub.c" ^
   /link /SUBSYSTEM:WINDOWS

if errorlevel 1 (
    echo [FAILED] stub.exe compilation failed.
    exit /b 1
)

:: Clean intermediate files
del /q "%OUTDIR%\*.obj" "%OUTDIR%\*.exp" 2>nul

echo [3/3] Build complete. Output: %OUTDIR%
