#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    AD Group Report - Bulk Load версия (быстрая)
    Вместо тысяч LDAP-запросов делает ДВА запроса и считает в памяти.
    Compatible: Windows Server 2019, PowerShell 5.1+
.NOTES
    Сохранять в UTF-8 with BOM
    В VS Code: правый нижний угол -> UTF-8 -> Save with Encoding -> UTF-8 with BOM
#>

param(
    [int]$ThrottleLimit = 16,
    [string]$OutputPath = "C:\Temp\AD_Groups_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Import-Module ActiveDirectory -ErrorAction Stop

# Все русские строки — только здесь, в одном месте.
# Внутри хэш-таблиц @{N=...; E=...} кириллица не используется,
# чтобы парсер PowerShell 5.1 не спотыкался на кодировке.
$col1 = 'Группа'
$col2 = 'Вложенных групп (1й уровень)'
$col3 = 'Пользователей напрямую'
$col4 = 'Пользователей всего (глубина)'
$col5 = 'Дата последнего изменения'
$col6 = 'Дней без изменений'

# ─── ШАГ 1: Два больших запроса вместо тысяч маленьких ───────────────────────
# Вся магия ускорения здесь: мы забираем ВСЕ группы и ВСЕХ пользователей
# одним запросом каждый, а дальше работаем только в памяти — DC больше не трогаем.
Write-Host "[$([datetime]::Now.ToString('HH:mm:ss'))] Загрузка всех групп из AD..." -ForegroundColor Cyan

$allGroups = Get-ADGroup -Filter * -Properties Members, whenChanged |
             Select-Object Name, DistinguishedName, Members, whenChanged

$totalGroups = $allGroups.Count
Write-Host "[$([datetime]::Now.ToString('HH:mm:ss'))] Загружено групп: $totalGroups" -ForegroundColor Green

# ─── ШАГ 2: Индекс групп в памяти ────────────────────────────────────────────
# Хэш-таблица DN -> объект группы позволяет находить любую группу
# за O(1) — мгновенно, без перебора всего массива.
Write-Host "[$([datetime]::Now.ToString('HH:mm:ss'))] Построение индекса групп в памяти..." -ForegroundColor Cyan

$groupIndex = @{}
foreach ($g in $allGroups) {
    $groupIndex[$g.DistinguishedName] = $g
}

# HashSet всех DN пользователей — для мгновенной проверки "это юзер или группа?"
# HashSet.Contains() работает за O(1), в отличие от массива где поиск O(n).
$allUserDNs = [System.Collections.Generic.HashSet[string]]::new(
    [string[]](Get-ADUser -Filter * | Select-Object -ExpandProperty DistinguishedName),
    [System.StringComparer]::OrdinalIgnoreCase
)
Write-Host "[$([datetime]::Now.ToString('HH:mm:ss'))] Загружено пользователей: $($allUserDNs.Count)" -ForegroundColor Green

# ─── ШАГ 3: Воркер для каждого Runspace ──────────────────────────────────────
# Этот блок выполняется в каждом параллельном потоке.
# Важно: у него НЕТ доступа к переменным основного скрипта —
# всё нужное передаётся явно через AddArgument().
$WorkerScript = {
    param(
        [string]$GroupDN,
        [string]$GroupName,
        [hashtable]$GroupIndex,
        [object]$UserDNs,
        [datetime]$WhenChanged
    )

    # Рекурсивный обход всей глубины вложенности.
    # HashSet $Visited защищает от бесконечной рекурсии при циклических ссылках.
    function Get-NestedUserCount {
        param(
            [string]$DN,
            [hashtable]$Idx,
            [object]$Users,
            [System.Collections.Generic.HashSet[string]]$Visited
        )
        # Если эту группу уже посещали — прерываем, чтобы не зациклиться
        if (-not $Visited.Add($DN)) { return 0 }

        $count = 0
        $group = $Idx[$DN]
        if (-not $group) { return 0 }

        foreach ($memberDN in $group.Members) {
            if ($Users.Contains($memberDN)) {
                # Это пользователь — считаем
                $count++
            } elseif ($Idx.ContainsKey($memberDN)) {
                # Это вложенная группа — уходим глубже рекурсивно
                $count += Get-NestedUserCount -DN $memberDN -Idx $Idx -Users $Users -Visited $Visited
            }
            # Компьютеры, контакты и прочее — молча пропускаем
        }
        return $count
    }

    $group = $GroupIndex[$GroupDN]
    if (-not $group) {
        return [PSCustomObject]@{
            GroupName    = $GroupName
            DirectGroups = 0
            DirectUsers  = 0
            TotalUsers   = 0
            WhenChanged  = $WhenChanged
        }
    }

    $members      = $group.Members
    $directUsers  = 0
    $directGroups = 0

    foreach ($memberDN in $members) {
        if      ($UserDNs.Contains($memberDN))       { $directUsers++ }
        elseif  ($GroupIndex.ContainsKey($memberDN)) { $directGroups++ }
    }

    $visited    = [System.Collections.Generic.HashSet[string]]::new()
    $totalUsers = Get-NestedUserCount -DN $GroupDN -Idx $GroupIndex -Users $UserDNs -Visited $visited

    return [PSCustomObject]@{
        GroupName    = $GroupName
        DirectGroups = $directGroups
        DirectUsers  = $directUsers
        TotalUsers   = $totalUsers
        WhenChanged  = $WhenChanged
    }
}

# ─── ШАГ 4: Запуск Runspace Pool ─────────────────────────────────────────────
# RunspacePool создаёт пул до $ThrottleLimit параллельных потоков.
# Каждый поток независим и не конкурирует за AD — только считает в памяти.
Write-Host "[$([datetime]::Now.ToString('HH:mm:ss'))] Запуск параллельной обработки (потоков: $ThrottleLimit)..." -ForegroundColor Yellow

$pool = [RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit)
$pool.ApartmentState = 'MTA'   # Multi-Threaded Apartment — обязательно для работы с AD-объектами
$pool.Open()

$jobs = [System.Collections.Generic.List[hashtable]]::new()

foreach ($group in $allGroups) {
    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($WorkerScript)
    [void]$ps.AddArgument($group.DistinguishedName)
    [void]$ps.AddArgument($group.Name)
    [void]$ps.AddArgument($groupIndex)      # весь индекс — в каждый поток
    [void]$ps.AddArgument($allUserDNs)      # HashSet пользователей — в каждый поток
    [void]$ps.AddArgument($group.whenChanged)

    $jobs.Add(@{
        PS     = $ps
        Handle = $ps.BeginInvoke()   # асинхронный запуск — не блокирует основной поток
        Name   = $group.Name
    })
}

# ─── ШАГ 5: Сбор результатов с прогресс-баром ────────────────────────────────
$results   = [System.Collections.Generic.List[object]]::new()
$done      = 0
$startTime = [datetime]::Now

while ($done -lt $jobs.Count) {
    foreach ($job in ($jobs | Where-Object { $null -ne $_.Handle -and $_.Handle.IsCompleted })) {
        try {
            $result = $job.PS.EndInvoke($job.Handle)
            if ($result) { $results.Add($result) }
        } catch {
            # Поток упал — добавляем запись с нулями, чтобы группа не потерялась в отчёте
            $results.Add([PSCustomObject]@{
                GroupName    = $job.Name
                DirectGroups = 0
                DirectUsers  = 0
                TotalUsers   = 0
                WhenChanged  = $null
            })
        } finally {
            $job.PS.Dispose()
            $job.Handle = $null
            $done++
        }
    }

    $elapsed = ([datetime]::Now - $startTime).TotalSeconds
    $rate    = if ($elapsed -gt 0) { [math]::Round($done / $elapsed, 1) } else { 0 }
    $eta     = if ($rate -gt 0)    { [math]::Round(($jobs.Count - $done) / $rate) } else { '?' }

    Write-Progress `
        -Activity "Обработка групп AD" `
        -Status   "$done / $totalGroups  |  $rate гр/сек  |  ETA: ${eta}с" `
        -PercentComplete ([math]::Min(100, ($done / $totalGroups * 100)))

    if ($done -lt $jobs.Count) { Start-Sleep -Milliseconds 300 }
}

Write-Progress -Activity "Обработка групп AD" -Completed
$pool.Close()
$pool.Dispose()

# ─── ШАГ 6: Сортировка и экспорт ─────────────────────────────────────────────
Write-Host "`n[$([datetime]::Now.ToString('HH:mm:ss'))] Формирование отчёта..." -ForegroundColor Cyan

$today = [datetime]::Today

# Явное приведение к [int] в Sort-Object критически важно:
# без него PowerShell сортирует как строки ("10" < "2"),
# с ним — как числа (2 < 10). Это и была причина неправильной сортировки.
$report = $results |
    Sort-Object { [int]$_.TotalUsers } |
    Select-Object `
        @{ N=$col1; E={ $_.GroupName } },
        @{ N=$col2; E={ [int]$_.DirectGroups } },
        @{ N=$col3; E={ [int]$_.DirectUsers } },
        @{ N=$col4; E={ [int]$_.TotalUsers } },
        @{ N=$col5; E={ if ($_.WhenChanged) { $_.WhenChanged.ToString('dd.MM.yyyy HH:mm') } else { 'н/д' } } },
        @{ N=$col6; E={ if ($_.WhenChanged) { [int]($today - $_.WhenChanged.Date).Days } else { 'н/д' } } }

$report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

$elapsed = [math]::Round(([datetime]::Now - $startTime).TotalSeconds)
Write-Host "[$([datetime]::Now.ToString('HH:mm:ss'))] Готово за ${elapsed}с. Отчёт: $OutputPath" -ForegroundColor Green
Write-Host "Всего групп в отчёте: $($report.Count)" -ForegroundColor Cyan