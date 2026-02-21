@echo off
setlocal enabledelayedexpansion
mode con: cols=45 lines=40
powershell -command "& {$h=$host.ui.rawui;$b=$h.buffersize;$w=$h.windowsize;$b.height=999;$h.buffersize=$b}"

:: Force BAT path
cd /d "%~dp0"

:: Check setup.exe
if not exist "%~dp0setup.exe" (
    echo ERROR: setup.exe not found in the BAT folder.
	echo Download from https://www.microsoft.com/en-us/download/details.aspx?id=49117
	start https://www.microsoft.com/en-us/download/details.aspx?id=49117
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
echo  I. Install
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

if /i not "%opt%"=="I" goto MENU

::Language Select
cls
:LANGUAGE
echo ============================================
echo   Select the language to install
echo ============================================
echo.
echo  1. English (en-us)
echo  2. English UK (en-gb)
echo  3. Spanish (es-es)
echo  4. Spanish LATAM (es-mx)
echo  5. French (fr-fr)
echo  6. German (de-de)
echo  7. Italian (it-it)
echo  8. Portuguese BR (pt-br)
echo  9. Portuguese PT (pt-pt)
echo 10. Japanese (ja-jp)
echo 11. Korean (ko-kr)
echo 12. Chinese Simplified (zh-cn)
echo 13. Chinese Traditional (zh-tw)
echo 14. Russian (ru-ru)
echo 15. Arabic (ar-sa)
echo 16. Dutch (nl-nl)
echo 17. Swedish (sv-se)
echo 18. Norwegian (nb-no)
echo 19. Danish (da-dk)
echo 20. Finnish (fi-fi)
echo 21. Polish (pl-pl)
echo 22. Czech (cs-cz)
echo 23. Hungarian (hu-hu)
echo 24. Turkish (tr-tr)
echo 25. Greek (el-gr)
echo 26. Romanian (ro-ro)
echo 27. Ukrainian (uk-ua)
echo 28. Indonesian (id-id)
echo 29. Vietnamese (vi-vn)
echo 30. Thai (th-th)
echo 31. Hebrew (he-il)
echo 32. Hindi (hi-in)
echo.

set /p lang=Choose an option: 

if "%lang%"=="1"  set "language=en-us"
if "%lang%"=="2"  set "language=en-gb"
if "%lang%"=="3"  set "language=es-es"
if "%lang%"=="4"  set "language=es-mx"
if "%lang%"=="5"  set "language=fr-fr"
if "%lang%"=="6"  set "language=de-de"
if "%lang%"=="7"  set "language=it-it"
if "%lang%"=="8"  set "language=pt-br"
if "%lang%"=="9"  set "language=pt-pt"
if "%lang%"=="10" set "language=ja-jp"
if "%lang%"=="11" set "language=ko-kr"
if "%lang%"=="12" set "language=zh-cn"
if "%lang%"=="13" set "language=zh-tw"
if "%lang%"=="14" set "language=ru-ru"
if "%lang%"=="15" set "language=ar-sa"
if "%lang%"=="16" set "language=nl-nl"
if "%lang%"=="17" set "language=sv-se"
if "%lang%"=="18" set "language=nb-no"
if "%lang%"=="19" set "language=da-dk"
if "%lang%"=="20" set "language=fi-fi"
if "%lang%"=="21" set "language=pl-pl"
if "%lang%"=="22" set "language=cs-cz"
if "%lang%"=="23" set "language=hu-hu"
if "%lang%"=="24" set "language=tr-tr"
if "%lang%"=="25" set "language=el-gr"
if "%lang%"=="26" set "language=ro-ro"
if "%lang%"=="27" set "language=uk-ua"
if "%lang%"=="28" set "language=id-id"
if "%lang%"=="29" set "language=vi-vn"
if "%lang%"=="30" set "language=th-th"
if "%lang%"=="31" set "language=he-il"
if "%lang%"=="32" set "language=hi-in"

if not defined language (
	echo.
	echo Language not found
	echo.
	pause
	goto menu
)

:: ===== GENERATE XML =====
echo.
echo Generating XML...
echo.

(
echo ^<Configuration ID="b53935ac-439f-4c47-ab4b-48899f68c2c2"^>
echo   ^<Add OfficeClientEdition="64" Channel="PerpetualVL2024"^>
echo     ^<Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB"^>
echo       ^<Language ID="%language%" /^>

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