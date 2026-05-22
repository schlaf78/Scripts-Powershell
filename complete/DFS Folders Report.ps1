Import-Module DFSN

$DomainFQDN = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name

$CsvFile = Join-Path $env:USERPROFILE "Downloads\DFS_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

$Roots = Get-DfsnRoot -Domain $DomainFQDN

$Result = @()

foreach ($Root in $Roots) {

    # DFS root header row (inside DFSPath column only)
    $Result += [PSCustomObject]@{
        DFSPath     = "===== $($Root.Path) ====="
        TargetPath  = ""
        Description = ""
    }

    $Folders = Get-DfsnFolder -Path "$($Root.Path)\*" -ErrorAction SilentlyContinue

    foreach ($Folder in $Folders) {

        $Targets = Get-DfsnFolderTarget -Path $Folder.Path -ErrorAction SilentlyContinue

        if ($Targets) {
            foreach ($Target in $Targets) {

                $Result += [PSCustomObject]@{
                    DFSPath     = $Folder.Path
                    TargetPath  = $Target.TargetPath
                    Description = $Folder.Description
                }
            }
        }
        else {

            $Result += [PSCustomObject]@{
                DFSPath     = $Folder.Path
                TargetPath  = ""
                Description = $Folder.Description
            }
        }
    }

    # blank separator row
    $Result += [PSCustomObject]@{
        DFSPath     = ""
        TargetPath  = ""
        Description = ""
    }
}

# Excel-friendly export
$Result |
    Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8 -Delimiter ";"

Write-Host "Exported: $CsvFile"