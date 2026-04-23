@echo off
rem ===================================================================
rem Build 2htm.exe from 2htm.cs.
rem
rem Requires:
rem   - Windows 10 or later.
rem   - A Roslyn-based C# compiler (csc.exe). Roslyn is the modern
rem     compiler that supports current language features. It ships
rem     with:
rem       - Visual Studio 2017 or later (any edition, including the
rem         free Community edition), OR
rem       - Visual Studio Build Tools 2019/2022 (a free, smaller
rem         download that installs just the compiler). Download from:
rem         https://visualstudio.microsoft.com/downloads/
rem
rem   NOTE: The compiler at
rem         %WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
rem         that ships with .NET Framework 4.x is the older
rem         pre-Roslyn compiler and cannot build 2htm (it only
rem         supports C# 5 and earlier).
rem
rem If Markdig.dll is not present in the build folder, it is
rem fetched from nuget.org on first build (requires internet access
rem that first time only). Markdig.dll is embedded as a manifest
rem resource; the resulting 2htm.exe is a true single-file
rem executable and does NOT need Markdig.dll at runtime.
rem ===================================================================

setlocal
cd /d "%~dp0"

rem ---- Locate a Roslyn csc.exe ---------------------------------------
rem
rem Search known install locations for Roslyn. The .NET Framework
rem redistributable csc.exe is intentionally NOT used here — it is
rem pre-Roslyn and cannot compile modern C#.
rem -------------------------------------------------------------------
set "c_sCsc="

rem Visual Studio Build Tools 2022
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe" (
    set "c_sCsc=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe"
    goto :cscFound
)

rem Visual Studio 2022 editions (Community, Professional, Enterprise)
for %%E in (Community Professional Enterprise) do (
    if exist "C:\Program Files\Microsoft Visual Studio\2022\%%E\MSBuild\Current\Bin\Roslyn\csc.exe" (
        set "c_sCsc=C:\Program Files\Microsoft Visual Studio\2022\%%E\MSBuild\Current\Bin\Roslyn\csc.exe"
        goto :cscFound
    )
    if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\%%E\MSBuild\Current\Bin\Roslyn\csc.exe" (
        set "c_sCsc=C:\Program Files (x86)\Microsoft Visual Studio\2022\%%E\MSBuild\Current\Bin\Roslyn\csc.exe"
        goto :cscFound
    )
)

rem Visual Studio Build Tools 2019
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe" (
    set "c_sCsc=C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe"
    goto :cscFound
)

rem Visual Studio 2019 editions
for %%E in (Community Professional Enterprise) do (
    if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\%%E\MSBuild\Current\Bin\Roslyn\csc.exe" (
        set "c_sCsc=C:\Program Files (x86)\Microsoft Visual Studio\2019\%%E\MSBuild\Current\Bin\Roslyn\csc.exe"
        goto :cscFound
    )
)

rem Visual Studio 2017 editions
for %%E in (Community Professional Enterprise BuildTools) do (
    if exist "C:\Program Files (x86)\Microsoft Visual Studio\2017\%%E\MSBuild\15.0\Bin\Roslyn\csc.exe" (
        set "c_sCsc=C:\Program Files (x86)\Microsoft Visual Studio\2017\%%E\MSBuild\15.0\Bin\Roslyn\csc.exe"
        goto :cscFound
    )
)

rem Roslyn standalone in PATH
where csc.exe >nul 2>&1
if not errorlevel 1 (
    rem Make sure it's Roslyn, not the old Framework compiler. The
    rem Framework path contains "Microsoft.NET\Framework" so we
    rem reject it. Roslyn csc typically lives under a "Roslyn" dir.
    for /f "delims=" %%P in ('where csc.exe') do (
        echo %%P | findstr /i "Microsoft.NET\\Framework" >nul
        if errorlevel 1 (
            set "c_sCsc=%%P"
            goto :cscFound
        )
    )
)

echo [ERROR] Could not find a Roslyn C# compiler.
echo.
echo         2htm uses modern C# features and requires the Roslyn
echo         compiler that ships with Visual Studio 2017 or later.
echo.
echo         Install either:
echo           - Visual Studio 2022 Community (free):
echo             https://visualstudio.microsoft.com/downloads/
echo           - Visual Studio Build Tools 2022 (smaller, free):
echo             https://visualstudio.microsoft.com/downloads/
echo             (Under "Tools for Visual Studio" at the bottom of
echo              the page. During install, select the workload
echo              ".NET desktop build tools".)
echo.
echo         NOTE: The csc.exe included with .NET Framework at
echo         %%WINDIR%%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
echo         is the older pre-Roslyn compiler and will NOT work.
exit /b 2

:cscFound
echo [INFO] Using compiler: %c_sCsc%

rem ---- Markdig dependency -------------------------------------------
set "c_sMarkdigVersion=0.18.3"
set "c_sMarkdigUrl=https://www.nuget.org/api/v2/package/Markdig/%c_sMarkdigVersion%"

if not exist "Markdig.dll" call :fnFetchMarkdig
if not exist "Markdig.dll" (
    echo [ERROR] Markdig.dll could not be obtained. Build cannot proceed.
    exit /b 2
)

rem ---- Compile ------------------------------------------------------
rem Target platform: x64, to match the 64-bit Office that Microsoft
rem installs by default since Office 2019. If your Office is 32-bit,
rem change /platform:x64 to /platform:x86 and rebuild. A 64-bit
rem process cannot automate a 32-bit Office COM server and vice
rem versa.
"%c_sCsc%" /nologo /target:exe /platform:x64 /optimize+ ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    /reference:System.Xml.dll ^
    /reference:System.IO.Compression.dll ^
    /reference:System.IO.Compression.FileSystem.dll ^
    /reference:System.Drawing.dll ^
    /reference:System.Windows.Forms.dll ^
    /reference:Microsoft.CSharp.dll ^
    /reference:Markdig.dll ^
    /resource:Markdig.dll,Markdig.dll ^
    /out:2htm.exe ^
    2htm.cs

if errorlevel 1 (
    echo [ERROR] Build failed.
    exit /b 1
)
echo [INFO] Built 2htm.exe successfully.

endlocal
goto :eof

rem ===================================================================
rem Fetch Markdig.dll from nuget.org. The .nupkg is a ZIP archive;
rem we extract lib/net40/Markdig.dll (or net35) into the build folder.
rem Uses built-in curl (Windows 10 1803+) and PowerShell's
rem Expand-Archive for extraction. Temp files are cleaned up.
rem ===================================================================
:fnFetchMarkdig
    echo [INFO] Markdig.dll not found. Fetching Markdig %c_sMarkdigVersion% from nuget.org ...
    set "sTempDir=%TEMP%\2htm_markdig_%RANDOM%%RANDOM%"
    mkdir "%sTempDir%" 2>nul
    if not exist "%sTempDir%" (
        echo [ERROR] Could not create temp directory %sTempDir%.
        goto :eof
    )

    curl -sL -o "%sTempDir%\markdig.zip" "%c_sMarkdigUrl%"
    if errorlevel 1 (
        echo [ERROR] Download failed. Check your internet connection or firewall.
        echo         URL: %c_sMarkdigUrl%
        goto :fnFetchCleanup
    )

    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "Expand-Archive -Path '%sTempDir%\markdig.zip' -DestinationPath '%sTempDir%\unpacked' -Force"
    if errorlevel 1 (
        echo [ERROR] Could not extract Markdig nupkg.
        goto :fnFetchCleanup
    )

    rem The nupkg targets .NET Framework 3.5 and 4.0. Prefer net40.
    if exist "%sTempDir%\unpacked\lib\net40\Markdig.dll" (
        copy /y "%sTempDir%\unpacked\lib\net40\Markdig.dll" "Markdig.dll" >nul
    ) else if exist "%sTempDir%\unpacked\lib\net35\Markdig.dll" (
        copy /y "%sTempDir%\unpacked\lib\net35\Markdig.dll" "Markdig.dll" >nul
    ) else (
        echo [ERROR] Expected Markdig.dll not found in the nupkg. Package contents:
        dir /s /b "%sTempDir%\unpacked\lib"
        goto :fnFetchCleanup
    )

    echo [INFO] Markdig.dll fetched successfully.

:fnFetchCleanup
    if exist "%sTempDir%" rmdir /s /q "%sTempDir%"
    goto :eof
