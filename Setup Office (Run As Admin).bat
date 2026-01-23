@echo off
setlocal enabledelayedexpansion
mode con: cols=45 lines=40

:: Force BAT path
cd /d "%~dp0"

:: Check setup.exe
if not exist "%~dp0setup.exe" (
    echo ERROR: setup.exe not found in the BAT folder.
    pause
    exit /b
)

:: ===== APPLICATION LIST =====
:: 1 = enabled (NOT excluded)
:: 0 = disabled (excluded with ExcludeApp)
set Word=1
set Excel=1
set PowerPoint=0
set Outlook=0
set Access=0
set OneNote=0
set Teams=0
set Publisher=0
set Project=0
set Visio=0

:MENU
cls
echo ============================================
echo   Select the applications to install
echo   (1 = enabled / 0 = disabled)
echo ============================================
echo.
echo  1. Word.............. !Word!
echo  2. Excel............. !Excel!
echo  3. PowerPoint........ !PowerPoint!
echo  4. Outlook........... !Outlook!
echo  5. Access............ !Access!
echo  6. OneNote........... !OneNote!
echo  7. Teams............. !Teams!
echo  8. Publisher......... !Publisher!
echo  9. Project LTSC...... !Project!
echo 10. Visio LTSC........ !Visio!
echo.
echo  S. Install
echo  X. Exit
echo.
set /p opt=Choose an option: 

if "%opt%"=="1"  (if !Word!==1       (set Word=0)      else (set Word=1))       & goto MENU
if "%opt%"=="2"  (if !Excel!==1      (set Excel=0)     else (set Excel=1))      & goto MENU
if "%opt%"=="3"  (if !PowerPoint!==1 (set PowerPoint=0) else (set PowerPoint=1)) & goto MENU
if "%opt%"=="4"  (if !Outlook!==1    (set Outlook=0)   else (set Outlook=1))    & goto MENU
if "%opt%"=="5"  (if !Access!==1     (set Access=0)    else (set Access=1))     & goto MENU
if "%opt%"=="6"  (if !OneNote!==1    (set OneNote=0)   else (set OneNote=1))    & goto MENU
if "%opt%"=="7"  (if !Teams!==1      (set Teams=0)     else (set Teams=1))      & goto MENU
if "%opt%"=="8"  (if !Publisher!==1  (set Publisher=0) else (set Publisher=1))  & goto MENU
if "%opt%"=="9"  (if !Project!==1    (set Project=0)   else (set Project=1))    & goto MENU
if "%opt%"=="10" (if !Visio!==1      (set Visio=0)     else (set Visio=1))      & goto MENU

if /i "%opt%"=="X" (
    if exist "%~dp0config_auto.xml" del /f /q "%~dp0config_auto.xml"
    exit /b
)

if /i not "%opt%"=="S" goto MENU


:: ===== GENERATE XML =====
echo.
echo Generating XML...
echo.

(
echo ^<Configuration ID="b53935ac-439f-4c47-ab4b-48899f68c2c2"^>
echo   ^<Add OfficeClientEdition="64" Channel="PerpetualVL2024"^>
echo     ^<Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB"^>
echo       ^<Language ID="es-es" /^>

if %Word%==0        echo       ^<ExcludeApp ID="Word" /^>
if %Excel%==0       echo       ^<ExcludeApp ID="Excel" /^>
if %PowerPoint%==0  echo       ^<ExcludeApp ID="PowerPoint" /^>
if %Outlook%==0     echo       ^<ExcludeApp ID="Outlook" /^>
if %Access%==0      echo       ^<ExcludeApp ID="Access" /^>
if %OneNote%==0     echo       ^<ExcludeApp ID="OneNote" /^>
if %Teams%==0       echo       ^<ExcludeApp ID="Teams" /^>
if %Publisher%==0   echo       ^<ExcludeApp ID="Publisher" /^>
if %Project%==0     echo       ^<ExcludeApp ID="Project" /^>
if %Visio%==0       echo       ^<ExcludeApp ID="Visio" /^>

echo     ^</Product^>
echo   ^</Add^>

echo   ^<Property Name="SharedComputerLicensing" Value="0" /^>
echo   ^<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /^>
echo   ^<Property Name="DeviceBasedLicensing" Value="0" /^>
echo   ^<Property Name="SCLCacheOverride" Value="0" /^>
echo   ^<Property Name="AUTOACTIVATE" Value="1" /^>

echo   ^<Updates Enabled="TRUE" /^>
echo   ^<RemoveMSI /^>
echo   ^<Display Level="Full" AcceptEULA="TRUE" /^>

echo ^</Configuration^>
) > "%~dp0config_auto.xml"

echo.
echo XML generated: config_auto.xml
echo.


:: ===== INSTALL =====
echo Starting installation...
echo Please wait while the process completes...
setup.exe /configure config_auto.xml

:: ===== DELETE XML =====
if exist "%~dp0config_auto.xml" del /f /q "%~dp0config_auto.xml"

echo.
echo Process completed.
pause

endlocal