param (
    [string]$FilePath = "\\<FULL PATH>\File.xlsx"
)

Import-Module ActiveDirectory

function Get-GroupUsers($Group, $From) {
    $users = @()
    try {
        foreach ($m in Get-ADGroupMember -Identity $Group) {
            switch ($m.objectClass) {
                'user'  { $users += [PSCustomObject]@{User=$m.SamAccountName; FromGroup=$From} }
                'group' { $users += Get-GroupUsers $m.SamAccountName $Group }
            }
        }
    } catch {
        Write-Warning ("Failed to expand '{0}': {1}" -f $Group, $_)
    }
    return $users
}

$results = @()

foreach ($ace in (Get-Acl $FilePath).Access) {
    $id = $ace.IdentityReference.Value
    $acct = $id.Split('\')[-1]
    $rights = $ace.FileSystemRights

    try {
        $obj = Get-ADObject -Filter { SamAccountName -eq $acct } -Properties objectClass
        if ($obj.objectClass -eq 'group') {
            $results += Get-GroupUsers $acct $id | ForEach-Object {
                $_ | Add-Member NoteProperty Rights $rights -Force; $_
            }
        } else {
            $results += [PSCustomObject]@{User=$acct; FromGroup='Direct'; Rights=$rights}
        }
    } catch {
        Write-Warning "Cannot resolve $id in AD"
    }
}

$results | Sort-Object User | Format-Table -AutoSize

$results | Export-Excel -Path "C:\output\file_permissions.xlsx" -AutoSize -WorksheetName "Permissions"