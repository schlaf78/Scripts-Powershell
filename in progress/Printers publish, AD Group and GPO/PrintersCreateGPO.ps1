#run this script on adm or any with rsat modules

Import-Module GroupPolicy

$InputList = Import-Csv -Path "C:\Users\<USERNAME>\Downloads\printers-export 8.csv"
$OU = "OU=printers,OU=Группы,DC=<DOMAIN>,DC=<DOMAIN>,DC=ru"

foreach ($entry in $InputList) {
    $PrinterName = $entry.PrinterName
    $GroupName   = $entry.GroupName

    # Create AD Group
    if (-not (Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue)) {
        New-ADGroup -Name $GroupName -Path $OU -GroupScope Global -PassThru
    }

    # Create and link GPO
    if (-not (Get-GPO -Name $PrinterName -ErrorAction SilentlyContinue)) {
        New-GPO -Name $PrinterName -ErrorAction Stop
    }

    New-GPLink -Name $PrinterName -Target "DC=<DOMAIN>,DC=<DOMAIN>,DC=ru" -LinkEnabled Yes

    # Set Security Filtering
    $GPO = Get-GPO -Name $PrinterName
    #Remove-GPPermission -Name $PrinterName -TargetName "Everyone" -TargetType Group -ErrorAction SilentlyContinue
    Remove-GPPermission -Name $PrinterName -TargetName "Everyone" -TargetType Group
    Set-GPPermission -Name $PrinterName -TargetName $GroupName -TargetType Group -PermissionLevel GpoApply
    Set-GPPermission -Name $PrinterName -TargetName "Компьютеры домена" -TargetType Group -PermissionLevel GpoApply

    # Delegation
    #Set-GPPermission -Name $PrinterName -Tar
    }