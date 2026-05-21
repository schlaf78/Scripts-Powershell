# ==========================================
# Инвентаризация принтеров + аудит GPO
# SECURITY = зафиксированный рабочий CIM метод (НЕ ЛОМАЕМ)
# GPO = парсинг PrinterConnections из XML (q1:Path) по UNC
# ==========================================

Import-Module ActiveDirectory
Import-Module GroupPolicy

# ImportExcel
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

$DomainName = (Get-ADDomain).DNSRoot

$PrintServers = @("prn-1","prn-2","prn-3","prn-4")

# ===== EXPORT =====
$ExportFolder = "C:\Temp"
$DateStamp = Get-Date -Format "yyyy-MM-dd"
$ExportPath = "$ExportFolder\Printers-Inventory_$DateStamp.xlsx"

if (-not (Test-Path $ExportFolder)) {
    New-Item -Path $ExportFolder -ItemType Directory | Out-Null
}

# ==========================================
# 1) GPO MAP (printer #N -> UNC из PrinterConnections)
# ==========================================

Write-Host "Сканирую GPO printer #..." -ForegroundColor Cyan

$GpoMap = @{}

# Берём все политики, начинающиеся на "printer #<число>"
$PrinterGPOs = Get-GPO -All | Where-Object {
    $_.DisplayName -match '^printer\s*#\s*\d+'
}

foreach ($Gpo in $PrinterGPOs) {

    # Достаём номер из имени политики
    if ($Gpo.DisplayName -notmatch '#\s*(\d+)') { continue }
    $Num = $Matches[1]

    try {
        # ВАЖНО: берём XML как строку (Get-GPOReport отдаёт строку)
        $GpoXmlText = Get-GPOReport -Guid $Gpo.Id -ReportType Xml

        # Ищем именно UNC (начинается с \\) внутри любого префикса *:Path
        # Это попадает в <q1:Path>\\server\Printer #10</q1:Path>
        $unc = $null
        if ($GpoXmlText -match '<[^:>]+:Path>(\\\\[^<]+)</[^:>]+:Path>') {
            $unc = $Matches[1]
        }

        if ($unc) {
            $GpoMap[$Num] = [PSCustomObject]@{
                GPOName = $Gpo.DisplayName
                UNCPath = $unc
            }
        }
        else {
            # Политика есть, но PrinterConnections path не найден
            $GpoMap[$Num] = [PSCustomObject]@{
                GPOName = $Gpo.DisplayName
                UNCPath = ""
            }
        }
    }
    catch {
        # Если какая-то политика не читается (права/битая) — фиксируем как пустую
        $GpoMap[$Num] = [PSCustomObject]@{
            GPOName = $Gpo.DisplayName
            UNCPath = ""
        }
    }
}

# ==========================================
# 2) INVENTORY printers + SECURITY + DRIVER + GPO columns
# ==========================================

$PrinterInventory = @()

foreach ($Server in $PrintServers) {

    Write-Host "Сканирую $Server ..." -ForegroundColor Cyan

    try {
        # Принтеры через PrintManagement
        $Printers = Get-Printer -ComputerName $Server -ErrorAction Stop

        # КЕШ Win32_Printer (ускорение)
        $CimPrinters = Get-CimInstance Win32_Printer -ComputerName $Server |
                       Group-Object Name -AsHashTable -AsString

        foreach ($Printer in $Printers) {

            # ========== SECURITY (ЗАФИКСИРОВАНО) ==========
            try {
                if ($CimPrinters.ContainsKey($Printer.Name)) {

                    $SD = $CimPrinters[$Printer.Name] |
                          Invoke-CimMethod -MethodName GetSecurityDescriptor

                    if ($SD.ReturnValue -eq 0) {
                        $SecurityGroups = (
                            $SD.Descriptor.DACL |
                                ForEach-Object { $_.Trustee.Name } |
                                Where-Object { $_ } |
                                Sort-Object -Unique
                        ) -join "; "
                    }
                    else {
                        $SecurityGroups = "Нет доступа к ACL (код $($SD.ReturnValue))"
                    }
                }
                else {
                    $SecurityGroups = "Принтер не найден в CIM"
                }
            }
            catch {
                $SecurityGroups = "Ошибка получения ACL"
            }

            # ========== ДРАЙВЕР ==========
            $Driver = Get-PrinterDriver -ComputerName $Server -Name $Printer.DriverName -ErrorAction SilentlyContinue

            if ($Driver) {
                $DriverVersion = $Driver.DriverVersion
                $DriverType    = $Driver.MajorVersion
            }
            else {
                $DriverVersion = "Unknown"
                $DriverType    = "Unknown"
            }

            # ========== GPO AUDIT (по номеру #) ==========
            $GPOName = "GPO not found"
            $UNCPath = ""

            if ($Printer.Name -match '#\s*(\d+)') {
                $Num = $Matches[1]

                if ($GpoMap.ContainsKey($Num)) {
                    $GPOName = $GpoMap[$Num].GPOName
                    $UNCPath = $GpoMap[$Num].UNCPath
                }
            }

            $PrinterInventory += [PSCustomObject]@{
                PrintServer            = $Server
                PrinterName            = $Printer.Name
                FQDN                   = "$($Printer.Name).$DomainName"
                DriverName             = $Printer.DriverName
                DriverVersion          = $DriverVersion
                DriverType             = $DriverType
                ShareName              = $Printer.ShareName
                PortName               = $Printer.PortName
                SecurityGroups         = $SecurityGroups
                GPOName                = $GPOName
                MappedPrinterFromGPO   = $UNCPath
            }
        }
    }
    catch {
        Write-Warning "Не удалось подключиться к $Server : $_"
    }
}

# ==========================================
# 3) SORT (как у тебя: PrintServer + номер принтера)
# ==========================================

$PrinterInventory = $PrinterInventory |
    Sort-Object PrintServer, @{
        Expression = {
            if ($_.PrinterName -match '#\s*(\d+)') { [int]$Matches[1] }
            else { $_.PrinterName }
        }
    }

# ==========================================
# 4) EXPORT
# ==========================================

if (Test-Path $ExportPath) {
    Remove-Item $ExportPath -Force
}

$PrinterInventory | Export-Excel `
    -Path $ExportPath `
    -AutoSize `
    -Title "Printers Inventory + GPO" `
    -FreezeTopRow `
    -BoldTopRow

Write-Host "`nГотово. Файл: $ExportPath" -ForegroundColor Green
