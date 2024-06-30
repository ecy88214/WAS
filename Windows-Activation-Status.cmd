@echo off
title WIndows Activation Status / Hardware Status Menu
:menu
cls
echo ========================================================================
echo           WIndows Activation Status / Hardware Status Menu         
echo ========================================================================
echo.
echo 1. Check Windows Activation
echo 2. Check Office Activation
echo 3. Check Hardware Configuration
echo 4. Exit
echo.
set /p choice=Please select an option (1, 2, 3, 4): 

if "%choice%"=="1" goto check_windows
if "%choice%"=="2" goto office_menu
if "%choice%"=="3" goto hardware_menu
if "%choice%"=="4" goto end
echo Invalid choice, please select 1, 2, 3 or 4.
pause
goto menu

:check_windows
cls
echo Checking Windows Activation Status...
cscript /nologo %SystemRoot%\System32\slmgr.vbs /dlv
pause
goto menu

:office_menu
cls
echo =========================================
echo         Office Version Selection         
echo =========================================
echo.
echo 1. Office 2016
echo 2. Office 2019
echo 3. Office 365
echo 4. Default (Auto-detect)
echo 5. Back to Main Menu
echo.
set /p office_choice=Please select your Office version (1, 2, 3, 4, 5): 

if "%office_choice%"=="1" goto check_office_2016
if "%office_choice%"=="2" goto check_office_2019
if "%office_choice%"=="3" goto check_office_365
if "%office_choice%"=="4" goto check_office_default
if "%office_choice%"=="5" goto menu
echo Invalid choice, please select 1, 2, 3, 4 or 5.
pause
goto office_menu

:check_office_2016
cls
echo Checking Office 2016 Activation Status...
cscript /nologo "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus
pause
goto office_menu

:check_office_2019
cls
echo Checking Office 2019 Activation Status...
cscript /nologo "C:\Program Files\Microsoft Office\Office17\OSPP.VBS" /dstatus
pause
goto office_menu

:check_office_365
cls
echo Checking Office 365 Activation Status...
cscript /nologo "C:\Program Files\Microsoft Office\Office18\OSPP.VBS" /dstatus
pause
goto office_menu

:check_office_default
cls
echo Auto-detecting Office version...
set "detected_version="
for %%O in (16 17 18) do (
    if exist "C:\Program Files\Microsoft Office\Office%%O\OSPP.VBS" (
        set "detected_version=Office%%O"
        goto detected
    ) else if exist "C:\Program Files (x86)\Microsoft Office\Office%%O\OSPP.VBS" (
        set "detected_version=Office%%O"
        goto detected
    )
)

:detected
if defined detected_version (
    echo Detected %detected_version%
    if exist "C:\Program Files\Microsoft Office\%detected_version%\OSPP.VBS" (
        cscript /nologo "C:\Program Files\Microsoft Office\%detected_version%\OSPP.VBS" /dstatus
    ) else (
        cscript /nologo "C:\Program Files (x86)\Microsoft Office\%detected_version%\OSPP.VBS" /dstatus
    )
) else (
    echo No Office installation detected.
)
pause
goto office_menu

:hardware_menu
cls
echo =========================================
echo       Hardware Configuration Menu       
echo =========================================
echo.
echo 1. Check CPU Information
echo 2. Check RAM Information
echo 3. Check Hard Disk Information
echo 4. Check Graphics Card Information
echo 5. Check Network Adapter Information
echo 6. Check BaseBoard Information
echo 7. Back to Main Menu
echo.
set /p hardware_choice=Please select an option (1-7): 

if "%hardware_choice%"=="1" goto check_cpu
if "%hardware_choice%"=="2" goto check_ram
if "%hardware_choice%"=="3" goto check_hard_disk
if "%hardware_choice%"=="4" goto check_graphics_card
if "%hardware_choice%"=="5" goto check_network_adapter
if "%hardware_choice%"=="6" goto check_baseboard
if "%hardware_choice%"=="7" goto menu
echo Invalid choice, please select 1, 2, 3, 4, 5, 6, or 7.
pause
goto hardware_menu

:check_cpu
cls
echo CPU Information:
wmic cpu get name, maxclockspeed, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, SocketDesignation
pause
goto hardware_menu

:check_ram
cls
echo RAM Information:
wmic memorychip get capacity, manufacturer, partnumber, speed, BankLabel
pause
goto hardware_menu

:check_hard_disk
cls
echo Hard Disk Information:
wmic diskdrive get model, size, Name, InterfaceType
pause
goto hardware_menu

:check_graphics_card
cls
echo Graphics Card Information:
wmic path win32_videocontroller get caption, name, adapterram, driverversion
pause
goto hardware_menu

:check_network_adapter
cls
echo Network Adapter Information:
wmic nic get name, macaddress
pause
goto hardware_menu

:check_baseboard
cls
echo BaseBoard Information:
wmic baseboard get Manufacturer, Model, Product, SerialNumber, Version
pause
goto hardware_menu

:end
echo Goodbye!
pause
exit
