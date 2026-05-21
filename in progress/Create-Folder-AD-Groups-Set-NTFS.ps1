Import-Module ActiveDirectory

# ------------------------------------------------
# LOG FUNCTION
# ------------------------------------------------

function Write-Log {

    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    switch ($Level) {

        "INFO"  { Write-Host "[INFO ] $Message" -ForegroundColor Cyan }
        "OK"    { Write-Host "[ OK  ] $Message" -ForegroundColor Green }
        "WARN"  { Write-Host "[WARN ] $Message" -ForegroundColor Yellow }
        "ERROR" { Write-Host "[FAIL ] $Message" -ForegroundColor Red }

    }

}

Write-Host ""
Write-Host "===== NTFS Folder + AD Groups Tool =====" -ForegroundColor Magenta
Write-Host ""

# ------------------------------------------------
# INPUT
# ------------------------------------------------

$FolderName = Read-Host "Введите имя папки"
$BasePath = Read-Host "Введите путь где создать папку"
$GroupBaseName = Read-Host "Введите базовое имя группы"

$FullFolderPath = Join-Path $BasePath $FolderName

# ------------------------------------------------
# CREATE FOLDER
# ------------------------------------------------

Write-Log "Этап 1/4 — создание папки"

if (Test-Path $FullFolderPath) {

    Write-Log "Папка уже существует: $FullFolderPath" "WARN"

}
else {

    try {

        New-Item -Path $FullFolderPath -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Log "Папка создана: $FullFolderPath" "OK"

    }
    catch {

        Write-Log "Ошибка создания папки $_" "ERROR"

    }

}

# ------------------------------------------------
# GROUP NAMES
# ------------------------------------------------

$GroupRO = "$GroupBaseName-ro"
$GroupRW = "$GroupBaseName-rw"
$GroupLS = "$GroupBaseName-ls"

# ------------------------------------------------
# OU PATHS
# ------------------------------------------------

$OU_RWRO = "OU=ntfs,OU=groups,OU=<OU NAME>,DC=<DOMAIN>,DC=<DOMAIN>,DC=ru"
$OU_LS = "OU=ls-groups,OU=ntfs,OU=<OU NAME>,OU=<OU NAME>,DC=<DOMAIN>,DC=<DOMAIN>,DC=ru"

# ------------------------------------------------
# CREATE GROUP FUNCTION
# ------------------------------------------------

function Create-ADGroupSafe {

    param(
        [string]$Name,
        [string]$OU
    )

    if (Get-ADGroup -Filter "Name -eq '$Name'" -ErrorAction SilentlyContinue) {

        Write-Log "Группа уже существует: $Name" "WARN"
        return $true

    }

    try {

        New-ADGroup `
            -Name $Name `
            -SamAccountName $Name `
            -GroupScope Global `
            -GroupCategory Security `
            -Path $OU `
            -ErrorAction Stop

        Write-Log "Создана группа $Name" "OK"
        return $true

    }
    catch {

        Write-Log "Ошибка создания группы $Name" "ERROR"
        return $false

    }

}

# ------------------------------------------------
# CREATE GROUPS
# ------------------------------------------------

Write-Log "Этап 2/4 — создание AD групп"

Create-ADGroupSafe $GroupRO $OU_RWRO
Create-ADGroupSafe $GroupRW $OU_RWRO
Create-ADGroupSafe $GroupLS $OU_LS

# ------------------------------------------------
# ADD GROUPS TO PARENT
# ------------------------------------------------

Write-Log "Этап 3/4 — добавление в родительскую группу"

$AddParent = Read-Host "Добавить группы в другую группу? (yes/no)"

if ($AddParent -match "yes|y|да") {

    $ParentGroup = Read-Host "Введите имя родительской группы"

    $ParentCheck = Get-ADGroup -Identity $ParentGroup -ErrorAction SilentlyContinue

    if (!$ParentCheck) {

        Write-Log "Родительская группа не найдена: $ParentGroup" "ERROR"

    }
    else {

        try {

            Add-ADGroupMember -Identity $ParentGroup -Members $GroupRO,$GroupRW,$GroupLS
            Write-Log "Группы добавлены в $ParentGroup" "OK"

        }
        catch {

            Write-Log "Ошибка добавления в группу $ParentGroup" "ERROR"

        }

    }

}
else {

    Write-Log "Добавление в родительскую группу пропущено"

}

# ------------------------------------------------
# NTFS PERMISSIONS
# ------------------------------------------------

Write-Log "Этап 4/4 — назначение NTFS прав"

$ACL = Get-Acl $FullFolderPath

try {

# RW
$ruleRW = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $GroupRW,
    "Modify",
    "ContainerInherit,ObjectInherit",
    "None",
    "Allow"
)

$ACL.AddAccessRule($ruleRW)

Write-Log "Назначены RW права (Modify)" "OK"

# RO
$ruleRO = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $GroupRO,
    "ReadAndExecute",
    "ContainerInherit,ObjectInherit",
    "None",
    "Allow"
)

$ACL.AddAccessRule($ruleRO)

Write-Log "Назначены RO права (Read)" "OK"

# LS
$ruleLS = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $GroupLS,
    "ListDirectory,ReadAttributes,ReadPermissions",
    "None",
    "None",
    "Allow"
)

$ACL.AddAccessRule($ruleLS)

Write-Log "Назначены LS права (List only, This folder only)" "OK"

Set-Acl -Path $FullFolderPath -AclObject $ACL

}
catch {

    Write-Log "Ошибка назначения NTFS прав $_" "ERROR"

}

# ------------------------------------------------
# SUMMARY
# ------------------------------------------------

Write-Host ""
Write-Host "===== RESULT =====" -ForegroundColor Magenta
Write-Host "Folder : $FullFolderPath"
Write-Host "Groups :"
Write-Host "  $GroupRO"
Write-Host "  $GroupRW"
Write-Host "  $GroupLS"
Write-Host ""
Write-Log "Скрипт завершён" "OK"