@echo off
setlocal enabledelayedexpansion

REM Specify the directory containing the folders
set "directory=C:\Users\aravng\Desktop\CDW\Archiving\FY24\UK"

REM Professional welcome message
echo ============================================================
echo Welcome to the Automated Zipping Utility
echo ============================================================
echo Current zipping directory: %directory%
echo Zipping process initiated. Please wait...
echo ============================================================

REM Change to the specified directory
cd /d "%directory%"

REM Initialize counter
set /a count=0

REM Loop through each folder in the directory
for /d %%f in (*) do (
    REM Create or update a zip file for each folder with the same name
    if exist "%%f.zip" (
        powershell Compress-Archive -Path "%%f\*" -Update -DestinationPath "%%f.zip"
        echo Updated %%f.zip
    ) else (
        powershell Compress-Archive -Path "%%f" -DestinationPath "%%f.zip"
        echo Created %%f.zip
    )
    REM Increment counter
    set /a count+=1
)

REM Print the total number of folders zipped
echo ============================================================
echo Total number of folders processed: %count%
echo ============================================================

REM Get the current date and time
for /f "tokens=1-4 delims=/ " %%a in ('date /t') do set curdate=%%a-%%b-%%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set curtime=%%a:%%b

REM Print the current date and time
echo Zipping completed on %curdate% at %curtime% . Please press C to close.
echo ============================================================

endlocal
pause



