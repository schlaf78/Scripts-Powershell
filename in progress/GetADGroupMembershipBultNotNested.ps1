Import-Module ActiveDirectory

# ================== ВВОД ГРУПП ==================
$Groups = @()

Write-Host "Введите имя AD-группы и нажмите ENTER."
Write-Host "Для завершения ввода введите END и нажмите ENTER.`n"

while ($true) {
    $groupName = Read-Host "Group"
    if ($groupName -eq "END") { break }
    if ($groupName) { $Groups += $groupName }
}

if ($Groups.Count -eq 0) {
    Write-Warning "Группы не указаны. Выход."
    return
}

# ================== EXCEL ==================
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)

$row = 1

foreach ($group in $Groups) {

    # ---- Заголовок группы ----
    $sheet.Cells.Item($row, 1).Value = "GROUP: $group"
    $sheet.Cells.Item($row, 1).Font.Bold = $true
    $row++

    # ---- Заголовки столбцов ----
    $sheet.Cells.Item($row, 1).Value = "Display Name"
    $sheet.Cells.Item($row, 2).Value = "SamAccountName"
    $sheet.Cells.Item($row, 3).Value = "Status"
    $sheet.Range("A$row:C$row").Font.Bold = $true
    $row++

    try {
        $users = Get-ADGroupMember -Identity $group -Recursive |
                 Where-Object { $_.objectClass -eq "user" } |
                 Get-ADUser -Properties DisplayName, SamAccountName, Enabled |
                 Sort-Object DisplayName
    }
    catch {
        $sheet.Cells.Item($row, 1).Value = "ERROR: group not found"
        $row += 2
        continue
    }

    foreach ($user in $users) {

        $sheet.Cells.Item($row, 1).Value = $user.DisplayName
        $sheet.Cells.Item($row, 2).Value = $user.SamAccountName

        if ($user.Enabled -eq $false) {
            $sheet.Cells.Item($row, 3).Value = "DISABLED"

            # подсветка строки
            $sheet.Range("A$row:C$row").Interior.ColorIndex = 15
        }
        else {
            $sheet.Cells.Item($row, 3).Value = "ENABLED"
        }

        $row++
    }

    # Пустая строка между группами
    $row++
}

# Автоширина колонок
$sheet.Columns.AutoFit()

# Сохранение
$path = "$env:USERPROFILE\Desktop\AD_Groups_Users.xlsx"
$workbook.SaveAs($path)
$workbook.Close()
$excel.Quit()

Write-Host "`nГотово. Файл создан:"
Write-Host $path
