

# Получаем группы AD
$AD_Groups = Get-ADGroup -Filter {Name -like "confluence-1core-*"} -SearchBase "OU=Jira Group,OU=группы,DC=DOMAIN,DC=DOMAIN,DC=ru" -Properties Description

# Создаем массив для хранения результатов
$AD_Group_Members = @()

# Перебираем каждую группу
foreach ($Group in $AD_Groups) {
    Write-Host "Обработка группы: $($Group.Name)"
    
    try {
        # Получаем членов группы
        $Members = Get-ADGroupMember -Identity $Group -ErrorAction Stop

        # Перебираем каждого пользователя в группе
        foreach ($Member in $Members) {
            try {
                # Получаем информацию о пользователе, включая должность (title)
                $User = Get-ADUser -Identity $Member.SamAccountName -Properties enabled, company, department, title -ErrorAction Stop

                # Преобразуем значение Enabled в "Активна" или "Отключена"
                $EnabledStatus = if ($User.Enabled) { "Активна" } else { "Отключена" }

                # Если пользователь найден, добавляем его данные в массив
                $AD_Group_Members += [PSCustomObject]@{
                    GroupName       = $Group.Name
                    GroupDescription = $Group.Description  
                    UserName        = $User.Name
                    SamAccountName  = $User.SamAccountName
                    Enabled         = $EnabledStatus
                    Company         = $User.Company
                    Department      = $User.Department
                    Title           = $User.Title  
                }
            } catch {
                Write-Warning "Ошибка при получении информации о пользователе $($Member.SamAccountName): $_"
            }
        }
    } catch {
        Write-Warning "Ошибка при получении членов группы $($Group.Name): $_"
    }
}

# Проверяем, есть ли данные для экспорта
if ($AD_Group_Members.Count -eq 0) {
    Write-Host "Нет данных для экспорта." -ForegroundColor Yellow
    exit
}

# Запрашиваем путь для экспорта файла
$SavePath = Read-Host "Введите путь для сохранения файла (например, C:\sys\.ps1\.csv\lic_result.csv)"

# Проверяем, введен ли путь
if (-not $SavePath) {
    Write-Host "Путь не указан. Экспорт отменен." -ForegroundColor Red
    exit
}

# Проверяем доступность пути для записи
$SaveDirectory = Split-Path $SavePath -Parent
if (-not (Test-Path -Path $SaveDirectory)) {
    New-Item -ItemType Directory -Path $SaveDirectory | Out-Null
}

# Экспортируем данные в CSV
try {
    $AD_Group_Members | Export-Csv -Path $SavePath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
    Write-Host "Экспорт завершен. Данные сохранены в $SavePath" -ForegroundColor Green
} catch {
    Write-Host "Ошибка при экспорте данных: $_" -ForegroundColor Red
}