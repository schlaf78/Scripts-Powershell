$ErrorActionPreference = "Continue"

$root = "\\<FULL FOLDER PATH>"

#$folders = Get-ChildItem $root -Directory
$folders = Get-ChildItem $root | Where-Object { $_.PSIsContainer }
$result = $folders | foreach {
    $path = $_.fullname
    $access = $path | get-acl
    $access.Access | foreach {
        $AccessControlType = $_.AccessControlType.tostring()
        $FileSystemRights = $_.FileSystemRights.tostring() -replace ", Synchronize|DeleteSubdirectoriesAndFiles, "
        $identities = $_ | where {$_.IdentityReference -match "DOMAIN\\"}
        $identities | foreach {
            $sam = $_.IdentityReference.Value -replace "DOMAIN\\"
            switch ((Get-ADObject -Filter {samaccountname -eq $sam}).ObjectClass) {
                {$_ -eq "user"} {
                    $adobject = Get-ADUser -Filter {enabled -eq $true -and samaccountname -eq $sam}
                    Write-Host "user [1] $($adobject.SamAccountName)" -ForegroundColor Yellow
                    If ($adobject) {
                        [pscustomobject]@{Folder=$path;Control=$AccessControlType;Permission=$FileSystemRights;User="$($adobject.name) ($($adobject.SamAccountName))"}
                    }
                    Else {
                        Write-Host "No $sam" -ForegroundColor Yellow
                    }
                    break
                }
                {$_ -eq "group"} {
                    Write-Host "group $sam" -ForegroundColor Yellow
                    Get-ADGroupMember $sam -Recursive | Get-ADUser | where {$_.enabled -eq $true} | foreach {
                        Write-Host "user [2] $($_.SamAccountName)" -ForegroundColor Yellow
                        [pscustomobject]@{Folder=$path;Control=$AccessControlType;Permission=$FileSystemRights;User="$($_.name) ($($_.SamAccountName)) (Group: $sam)"}
                    }
                    break
                }
                Default {Write-Host "Skip $sam : $_" -ForegroundColor Yellow}
            }
        }
    }
}


$result | sort -Unique User | Export-Csv -Delimiter ";" -Encoding UTF8 -Path C:\Reports\lr-dept-ntfs.csv -NoTypeInformation






#$path = "\\holding.melonfashion.ru\share\SPB.Office\Финансово-административный департамент"
#$targets = Get-ChildItem $path -Recurse -Directory -Depth 2


#$targets | foreach {
#    $fullname = $_.FullName -replace "\\FQDN path\share"
#    $owner = (get-acl $_.FullName -ErrorAction SilentlyContinue).Owner
#    [PSCustomObject]@{fullname=$fullname;owner=$owner}
#} | Export-Csv C:\Users\<user name>\Desktop\fad.csv -Delimiter ";" -Encoding UTF8 -NoTypeInformation