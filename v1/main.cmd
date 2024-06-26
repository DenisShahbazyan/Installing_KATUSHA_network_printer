@echo off
chcp 65001

echo Перед продолжением - удалите все принтеры KATUSHA из панели управления, если некоторые принтеры не удаляются - можно их оставить

pause

rem Очиста очереди печати
echo Остановка службы Spooler...
net stop Spooler

echo Очистка папки PRINTERS...
del /Q /F C:\Windows\System32\spool\PRINTERS\
echo.

echo Запуск службы Spooler...
net start Spooler


set "prndrvr=C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prndrvr.vbs"
set "prnport=C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prnport.vbs"
set "prnmngr=C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prnmngr.vbs"
set "prnqctl=C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prnqctl.vbs"

set "printerName=KATUSHA M247"
set "driverName=%printerName% PCL6"
set "driverPath=C:\KATUSHA_M247_PCL6\Driver\KTSMXS.inf"
set /p printerIP=Введите IP принтера:
set "printerPort=9100"

rem Установка драйвера принтера
cscript "%prndrvr%" -a -m "%driverName%" -i "%driverPath%"

rem Создаем TCP/IP порт для сетевого принтера
cscript "%prnport%" -a -r "IP_%printerIP%" -h "%printerIP%" -o raw -n "%printerPort%"

rem Установка нового принтера
cscript "%prnmngr%" -a -p "%printerName%" -m "%driverName%" -r "IP_%printerIP%"

set /p testPrint=Введите Y или нажмите Enter(чтобы напечатать тестовую страницу) или N для пропуска: 
if /i "%testPrint%"=="Y" (
    cscript "%prnqctl%" -e -p "%printerName%"
) 


pause

rem https://winitpro.ru/index.php/2014/03/03/ustanovka-printera-iz-komandnoj-stroki-v-windows-8/
