# Список серверов
#$servers = @("infra-test-1", "test-karavan")
$servers = Get-ADComputer -Filter 'Name -like "1c-*"' | Select-Object -ExpandProperty Name

# Функция получения размеров папок на диске C:
function Get-FolderSizes {
    param(
        [string]$ServerName
    )

    Invoke-Command -ComputerName $ServerName -ScriptBlock {
        Get-ChildItem -Path 'C:\' -Directory | ForEach-Object {
            $folderPath = $_.FullName
            try {
                $size = (Get-ChildItem -Path $folderPath -Recurse -Force -ErrorAction SilentlyContinue |
                    Measure-Object -Property Length -Sum).Sum
                [PSCustomObject]@{
                    ServerName = $env:COMPUTERNAME
                    FolderPath = $folderPath
                    SizeGB     = [math]::Round($size / 1GB, 2)
                }
            } catch {
                [PSCustomObject]@{
                    ServerName = $env:COMPUTERNAME
                    FolderPath = $folderPath
                    SizeGB     = "Error"
                }
            }
        }
    } -ErrorAction SilentlyContinue | Select-Object ServerName, FolderPath, SizeGB
}

# Инициализируем список результатов
$finalResults = @()

# Обрабатываем каждый сервер
foreach ($server in $servers) {
    $serverResults = Get-FolderSizes -ServerName $server
    $finalResults += $serverResults

    # Добавляем пустую строку (объект с пустыми значениями) между серверами
    $finalResults += [PSCustomObject]@{
        ServerName = ''
        FolderPath = ''
        SizeGB     = ''
    }
}

# Путь к файлу
$exportPath = "C:\Reports\CDriveFolderSizes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Экспорт с разделителем ';' для Excel
$finalResults | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'

Write-Host "CSV сохранен: $exportPath"
