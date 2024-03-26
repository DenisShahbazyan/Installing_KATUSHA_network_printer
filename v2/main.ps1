function Start-Administrator {
    $isAdmin = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544"
    if (-not $isAdmin) {
        Start-Process powershell -ArgumentList " -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
}

function Uninstall-PrinterAndComponents {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PrinterName
    )

    # Получаем информацию о принтере
    $printer = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
    if ($null -eq $printer) {
        Write-Host "Принтер с именем '$PrinterName' не найден."
        return
    }

    # Отменяем все задания печати
    Get-PrintJob -PrinterName $PrinterName | Remove-PrintJob

    # Удаляем принтер
    Remove-Printer -Name $PrinterName
    Write-Host "Принтер '$PrinterName' удален."

    # Удаляем порт, если он больше не используется
    $portName = $printer.PortName
    if (-not (Get-Printer | Where-Object { $_.PortName -eq $portName })) {
        Remove-PrinterPort -Name $portName
        Write-Host "Порт '$portName' удален."
    }

    # Удаляем драйвер, если он больше не используется
    $driverName = $printer.DriverName
    if (-not (Get-Printer | Where-Object { $_.DriverName -eq $driverName })) {
        Remove-PrinterDriver -Name $driverName
        Write-Host "Драйвер '$driverName' удален."
    }
}

function Uninstall-AllPrinters {
    $printers = Get-Printer | Where-Object {
        $_.Name -imatch "katusha" -or
        $_.Name -imatch "м247" -or
        $_.Name -imatch "m247" -or
        $_.DriverName -imatch "katusha" -or
        $_.DriverName -imatch "м247" -or
        $_.DriverName -imatch "m247"
    }

    # Удаляем принтеры, соответствующие условию
    foreach ($printer in $printers) {
        Uninstall-PrinterAndComponents -PrinterName $printer.Name
    }
}

function Restart-EthernetAdapter {
    $adapters = Get-NetAdapter | Where-Object { 
        $_.Status -eq "Up" -and 
        $_.MacAddress -ne $null -and 
        $_.MacAddress -ne "" 
    }
    
    foreach ($adapter in $adapters) {
        # Отключаем адаптер
        Disable-NetAdapter -Name $adapter.Name -Confirm:$false

        # Ждем, пока адаптер не отключится, но не более 20 секунд
        $timeout = 20
        do {
            Start-Sleep -Seconds 1
            $timeout--
            $pingResult = Test-Connection -ComputerName "8.8.8.8" -Count 1 -ErrorAction SilentlyContinue
        } while ($pingResult -and $timeout -gt 0)

        if ($timeout -le 0) {
            Write-Host "Время ожидания отключения адаптера $($adapter.Name) истекло."
        }
        else {
            Write-Host "Адаптер $($adapter.Name) успешно отключен."
        }

        # Включаем адаптер
        Enable-NetAdapter -Name $adapter.Name -Confirm:$false

        # Ждем, пока адаптер не включится, но не более 20 секунд
        $timeout = 20
        do {
            Start-Sleep -Seconds 1
            $timeout--
            $pingResult = Test-Connection -ComputerName "8.8.8.8" -Count 1 -ErrorAction SilentlyContinue
        } while (-not $pingResult -and $timeout -gt 0)

        if ($timeout -le 0) {
            Write-Host "Время ожидания включения адаптера $($adapter.Name) истекло."
        }
        else {
            Write-Host "Адаптер $($adapter.Name) успешно включен."
        }
    }
}

function Clear-SpoolerQueue {
    # Останавливаем службу
    $spoolerStatus = Get-Service -Name Spooler
    if ($spoolerStatus.Status -eq "Running") {
        Stop-Service -Name Spooler -Force
        
        $timeout = 10
        do {
            Start-Sleep -Seconds 1
            $timeout--
            $spoolerStatus = Get-Service -Name Spooler
        } while ($spoolerStatus.Status -ne "Stopped" -and $timeout -gt 0)

        if ($spoolerStatus.Status -ne "Stopped") {
            Write-Host "Не удалось остановить службу диспетчера печати."
            return
        }
    }

    # Очищаем очередь печати
    $printerQueuePath = "C:\Windows\System32\spool\PRINTERS\"
    Remove-Item -Path "$printerQueuePath*" -Force -Recurse -ErrorAction SilentlyContinue

    # Проверяем, пуста ли папка после удаления файлов
    $itemsInQueueAfter = Get-ChildItem -Path $printerQueuePath
    if ($itemsInQueueAfter) {
        Write-Host "Не удалось очистить очередь печати. В папке остались файлы."
    }
    else {
        Write-Host "Очередь печати успешно очищена."
    }

    # Запускаем службу
    Start-Service -Name Spooler
    $timeout = 10
    do {
        Start-Sleep -Seconds 1
        $timeout--
        $spoolerStatus = Get-Service -Name Spooler
    } while ($spoolerStatus.Status -ne "Running" -and $timeout -gt 0)

    if ($spoolerStatus.Status -ne "Running") {
        Write-Host "Не удалось запустить службу диспетчера печати."
    }
}

function Install-Printer {
    # Пути к скриптам управления принтерами
    $prndrvr = "C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prndrvr.vbs"
    $prnport = "C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prnport.vbs"
    $prnmngr = "C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prnmngr.vbs"
    $prnqctl = "C:\Windows\System32\Printing_Admin_Scripts\ru-RU\prnqctl.vbs"

    # Настройки принтера
    $printerName = "KATUSHA M247"
    $driverName = "$printerName PCL6"
    $driverPath = "C:\KATUSHA_M247_PCL6\Driver\KTSMXS.inf"
    $printerIP = Read-Host -Prompt "Введите IP принтера"
    $printerPort = "9100"

    # Установка драйвера принтера
    & cscript $prndrvr -a -m $driverName -i $driverPath

    # Создание TCP/IP порта для сетевого принтера
    & cscript $prnport -a -r "IP_$printerIP" -h $printerIP -o raw -n $printerPort

    # Установка нового принтера
    & cscript $prnmngr -a -p $printerName -m $driverName -r "IP_$printerIP"

    # Запрос на печать тестовой страницы
    $testPrint = Read-Host -Prompt "Введите Y и нажмите Enter (чтобы напечатать тестовую страницу) или любую другую клавишу для пропуска"
    if ($testPrint -ieq "Y") {
        & cscript $prnqctl -e -p $printerName
    }
}

Start-Administrator

$isClean = Read-Host -Prompt "Введите Y и нажмите Enter для запуска очистки, Или любую другую клавишу для пропуска"
if ($isClean -ieq "Y") {
    for ($i = 0; $i -lt 2; $i++) {
        Clear-SpoolerQueue
        Uninstall-AllPrinters
        Restart-EthernetAdapter
    }
}

Install-Printer

Pause