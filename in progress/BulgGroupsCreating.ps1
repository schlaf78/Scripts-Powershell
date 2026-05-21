# Параметры
$OU = "OU=ntfs,OU=groups,OU=<OU NAME>,DC=<DOMAIN>,DC=<DOMAIN>,DC=ru"
$Prefix = "ntfs-o-public-work-groups-"
$TargetParentGroup = "enable-policy-users-disk-r"   # Целевая группа

Write-Host "Введите имена через запятую (пробелы не имеют значения: proj1,proj2 или proj1, proj2)" -ForegroundColor Cyan
$InputNames = Read-Host
$Names = $InputNames -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

if (-not $Names) {
    Write-Host "Вы не ввели ни одного имени!" -ForegroundColor Yellow
    exit
}

# Формируем список групп
$GroupsToCreate = foreach ($Name in $Names) {
    "$Prefix$Name-ro"
    "$Prefix$Name-rw"
}

# Вывод списка
Write-Host "`nБудут созданы следующие группы в OU: ${OU}" -ForegroundColor Cyan
$GroupsToCreate | ForEach-Object { Write-Host " - $_" }

# Подтверждение создания
$Confirm = Read-Host "`nДля создания групп введите 'create'"

$CreatedGroups = @()

if ($Confirm -eq "create") {
    foreach ($Group in $GroupsToCreate) {
        try {
            # Проверка существования
            $exists = Get-ADGroup -Filter "SamAccountName -eq '$Group'" -ErrorAction SilentlyContinue
            if ($exists) {
                Write-Host "Группа уже существует: ${Group} (пропуск)" -ForegroundColor Yellow
                continue
            }

            # Создание
            New-ADGroup -Name $Group -SamAccountName $Group -GroupScope Global -GroupCategory Security -Path $OU -ErrorAction Stop
            Write-Host "Создана группа: ${Group}" -ForegroundColor Green
            $CreatedGroups += $Group

            # Добавление в родительскую группу
            try {
                Add-ADGroupMember -Identity $TargetParentGroup -Members $Group -ErrorAction Stop
                Write-Host "Добавлена в ${TargetParentGroup}: ${Group}" -ForegroundColor Green
            } catch {
                Write-Host ("Ошибка при добавлении {0} в {1}: {2}" -f $Group, $TargetParentGroup, $_.Exception.Message) -ForegroundColor Red
            }
        }
        catch {
            Write-Host ("Ошибка при создании {0}: {1}" -f $Group, $_.Exception.Message) -ForegroundColor Red
        }
    }

    # Итог
    if ($CreatedGroups.Count -gt 0) {
        Write-Host "`nИТОГ:" -ForegroundColor Cyan
        Write-Host ("Созданы группы: {0}" -f ($CreatedGroups -join ", ")) -ForegroundColor Green
        Write-Host ("Все они добавлены в {0}" -f $TargetParentGroup) -ForegroundColor Green
    } else {
        Write-Host "`nНовые группы не были созданы." -ForegroundColor Yellow
    }
} else {
    Write-Host "Операция отменена пользователем." -ForegroundColor Yellow
}
