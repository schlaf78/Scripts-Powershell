Import-Module ActiveDirectory
Import-Module ImportExcel

# =========================================================
# GLOBAL VARIABLES
# =========================================================

$Global:SelectedDomain = $null

# =========================================================
# COMMON FUNCTIONS
# =========================================================

function Pause-Script {

    Write-Host ""
    Read-Host "Press ENTER to continue"
}

# =========================================================
# XLSX EXPORT FUNCTION
# =========================================================

function Save-ReportToXlsx {

    param (
        [string]$ReportName,
        [array]$ReportData
    )

    try {

        $Date = Get-Date -Format "yyyy-MM-dd"

        $DocumentsPath = [Environment]::GetFolderPath("MyDocuments")

        $FileName = "$ReportName-$Date.xlsx"

        $FullPath = Join-Path $DocumentsPath $FileName

        $ReportData | Export-Excel `
            -Path $FullPath `
            -WorksheetName "Report" `
            -AutoSize `
            -BoldTopRow `
            -FreezeTopRow `
            -TableName "ReportTable" `
            -TableStyle Medium2

        Write-Host ""
        Write-Host "Success! File was saved:" -ForegroundColor Green
        Write-Host $FullPath -ForegroundColor Cyan

    }
    catch {

        Write-Host ""
        Write-Host "Failed to save XLSX report!" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }

    Pause-Script
}

# =========================================================
# TRANSLITERATION FUNCTION (UNICODE SAFE)
# =========================================================

function ConvertTo-Latin {

    param (
        [string]$Text
    )

    $Text = $Text.ToLower()
    $Result = ""

    foreach ($c in $Text.ToCharArray()) {

        switch ([int][char]$c) {

            1072 { $Result += "a"; break }
            1073 { $Result += "b"; break }
            1074 { $Result += "v"; break }
            1075 { $Result += "g"; break }
            1076 { $Result += "d"; break }
            1077 { $Result += "e"; break }
            1105 { $Result += "e"; break }
            1078 { $Result += "zh"; break }
            1079 { $Result += "z"; break }
            1080 { $Result += "i"; break }
            1081 { $Result += "y"; break }
            1082 { $Result += "k"; break }
            1083 { $Result += "l"; break }
            1084 { $Result += "m"; break }
            1085 { $Result += "n"; break }
            1086 { $Result += "o"; break }
            1087 { $Result += "p"; break }
            1088 { $Result += "r"; break }
            1089 { $Result += "s"; break }
            1090 { $Result += "t"; break }
            1091 { $Result += "u"; break }
            1092 { $Result += "f"; break }
            1093 { $Result += "h"; break }
            1094 { $Result += "ts"; break }
            1095 { $Result += "ch"; break }
            1096 { $Result += "sh"; break }
            1097 { $Result += "sch"; break }
            1099 { $Result += "y"; break }
            1101 { $Result += "e"; break }
            1102 { $Result += "yu"; break }
            1103 { $Result += "ya"; break }

            default {

                if ($c -match '[a-z0-9]') {
                    $Result += $c
                }
                else {
                    $Result += "-"
                }
            }
        }
    }

    $Result = $Result -replace '-+', '-'
    $Result = $Result.Trim('-')

    return $Result
}

# =========================================================
# DOMAIN SELECTION
# =========================================================

function Select-Domain {

    Clear-Host

    try {

        $Domains = (Get-ADForest).Domains

        Write-Host "Available domains:" -ForegroundColor Cyan
        Write-Host ""

        for ($i = 0; $i -lt $Domains.Count; $i++) {

            Write-Host ("{0}. {1}" -f ($i + 1), $Domains[$i])
        }

        Write-Host ""

        do {

            $DomainChoice = Read-Host "Select domain by number"

        } while (
            -not (
                $DomainChoice -as [int] -and
                $DomainChoice -ge 1 -and
                $DomainChoice -le $Domains.Count
            )
        )

        $Global:SelectedDomain = $Domains[$DomainChoice - 1]

        Write-Host ""
        Write-Host "Selected domain: $Global:SelectedDomain" -ForegroundColor Green
    }
    catch {

        Write-Host ""
        Write-Host "Failed to retrieve AD domains!" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }

    Pause-Script
}

# =========================================================
# MAIN MENU
# =========================================================

function Show-MainMenu {

    Clear-Host

    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host " AD MANAGEMENT TOOLKIT" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Current domain: $Global:SelectedDomain" -ForegroundColor Yellow
    Write-Host ""

    Write-Host "1. User Management"
    Write-Host "2. Folders Management"
    Write-Host "3. Reports"
    Write-Host "0. Exit"

    Write-Host ""

    $MenuChoice = Read-Host "Select option"

    switch ($MenuChoice) {

        "1" { Show-UserManagementMenu }
        "2" { Show-FoldersManagementMenu }
        "3" { Show-ReportsMenu }
        "0" { exit }

        default {

            Write-Host ""
            Write-Host "Invalid selection!" -ForegroundColor Red

            Pause-Script

            Show-MainMenu
        }
    }
}

# =========================================================
# USER MANAGEMENT MENU
# =========================================================

function Show-UserManagementMenu {

    Clear-Host

    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host " USER MANAGEMENT" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Current domain: $Global:SelectedDomain" -ForegroundColor Yellow
    Write-Host ""

    Write-Host "1. Create User (Regular Account)"
    Write-Host "2. Create User (ADM Account)"
    Write-Host "3. Create User (Support Account)"
    Write-Host "0. Back"

    Write-Host ""

    $MenuChoice = Read-Host "Select option"

    switch ($MenuChoice) {

        "1" {

            Write-Host ""
            Write-Host "Function is under development." -ForegroundColor Yellow

            Pause-Script

            Show-UserManagementMenu
        }

        "2" {

            Write-Host ""
            Write-Host "Function is under development." -ForegroundColor Yellow

            Pause-Script

            Show-UserManagementMenu
        }

        "3" {

            Write-Host ""
            Write-Host "Function is under development." -ForegroundColor Yellow

            Pause-Script

            Show-UserManagementMenu
        }

        "0" {

            Show-MainMenu
        }

        default {

            Write-Host ""
            Write-Host "Invalid selection!" -ForegroundColor Red

            Pause-Script

            Show-UserManagementMenu
        }
    }
}

# =========================================================
# FOLDERS MANAGEMENT MENU
# =========================================================

function Show-FoldersManagementMenu {

    Clear-Host

    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host " FOLDERS MANAGEMENT" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Current domain: $Global:SelectedDomain" -ForegroundColor Yellow
    Write-Host ""

    Write-Host "1. Create Folder (Drive R)"
    Write-Host "2. Create Folder (Drive O)"
    Write-Host "0. Back"

    Write-Host ""

    $MenuChoice = Read-Host "Select option"

    switch ($MenuChoice) {

        "1" {

            Write-Host ""
            Write-Host "Function is under development." -ForegroundColor Yellow

            Pause-Script

            Show-FoldersManagementMenu
        }

        "2" {

            Write-Host ""
            Write-Host "Function is under development." -ForegroundColor Yellow

            Pause-Script

            Show-FoldersManagementMenu
        }

        "0" {

            Show-MainMenu
        }

        default {

            Write-Host ""
            Write-Host "Invalid selection!" -ForegroundColor Red

            Pause-Script

            Show-FoldersManagementMenu
        }
    }
}

# =========================================================
# REPORTS MENU
# =========================================================

function Show-ReportsMenu {

    Clear-Host

    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host " REPORTS" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Current domain: $Global:SelectedDomain" -ForegroundColor Yellow
    Write-Host ""

    Write-Host "1. Folder Report"
    Write-Host "0. Back"

    Write-Host ""

    $MenuChoice = Read-Host "Select option"

    switch ($MenuChoice) {

        "1" { Show-FolderReportsMenu }
        "0" { Show-MainMenu }

        default {

            Write-Host ""
            Write-Host "Invalid selection!" -ForegroundColor Red

            Pause-Script

            Show-ReportsMenu
        }
    }
}

# =========================================================
# FOLDER REPORTS MENU
# =========================================================

function Show-FolderReportsMenu {

    Clear-Host

    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host " FOLDER REPORTS" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Current domain: $Global:SelectedDomain" -ForegroundColor Yellow
    Write-Host ""

    Write-Host "1. AD Single Group"
    Write-Host "2. Confluence Group Pack"
    Write-Host "3. Jira Group Pack"
    Write-Host "0. Back"

    Write-Host ""

    $MenuChoice = Read-Host "Select option"

    switch ($MenuChoice) {

        "1" { Get-ADSingleGroupReport }
        "2" { Get-ConfluenceGroupReport }
        "3" { Get-JiraGroupReport }
        "0" { Show-ReportsMenu }

        default {

            Write-Host ""
            Write-Host "Invalid selection!" -ForegroundColor Red

            Pause-Script

            Show-FolderReportsMenu
        }
    }
}

# =========================================================
# REPORT FUNCTIONS
# =========================================================

function Get-ADSingleGroupReport {

    try {

        Clear-Host

        $GroupName = Read-Host "Enter AD Group name"

        Write-Host ""
        Write-Host "Members of group: $GroupName" -ForegroundColor Cyan
        Write-Host ""

        $ReportData = Get-ADGroupMember $GroupName |
            Select-Object Name, SamAccountName

        $ReportData | Format-Table -AutoSize

        Write-Host ""

        $SaveChoice = Read-Host "Do you want to save report as XLSX file? (Y/N)"

        if ($SaveChoice -eq "Y") {

            Save-ReportToXlsx `
                -ReportName "AD-Single-Group-Report" `
                -ReportData $ReportData
        }
        else {

            Pause-Script
        }
    }
    catch {

        Write-Host ""
        Write-Host "Failed to generate report!" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red

        Pause-Script
    }

    Show-FolderReportsMenu
}

function Get-ConfluenceGroupReport {

    try {

        Clear-Host

        $GroupName = Read-Host "Enter Confluence Group name"

        $ReportData = @()

        Write-Host ""

        foreach ($Suffix in @("adm", "rw", "ro")) {

            $CurrentGroup = "$GroupName-$Suffix"

            Write-Host "$CurrentGroup members" -ForegroundColor Cyan

            $GroupMembers = Get-ADGroupMember $CurrentGroup |
                Select-Object @{
                    Name = "Group"
                    Expression = { $CurrentGroup }
                }, Name, SamAccountName

            $GroupMembers | Format-Table -AutoSize

            $ReportData += $GroupMembers

            Write-Host ""
        }

        $SaveChoice = Read-Host "Do you want to save report as XLSX file? (Y/N)"

        if ($SaveChoice -eq "Y") {

            Save-ReportToXlsx `
                -ReportName "Confluence-Groups-Report" `
                -ReportData $ReportData
        }
        else {

            Pause-Script
        }
    }
    catch {

        Write-Host ""
        Write-Host "Failed to generate report!" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red

        Pause-Script
    }

    Show-FolderReportsMenu
}

function Get-JiraGroupReport {

    try {

        Clear-Host

        $GroupName = Read-Host "Enter Jira Group name"

        $ReportData = @()

        Write-Host ""

        foreach ($Suffix in @("adm", "anl", "dev", "usr")) {

            $CurrentGroup = "$GroupName-$Suffix"

            Write-Host "$CurrentGroup members" -ForegroundColor Cyan

            $GroupMembers = Get-ADGroupMember $CurrentGroup |
                Select-Object @{
                    Name = "Group"
                    Expression = { $CurrentGroup }
                }, Name, SamAccountName

            $GroupMembers | Format-Table -AutoSize

            $ReportData += $GroupMembers

            Write-Host ""
        }

        $SaveChoice = Read-Host "Do you want to save report as XLSX file? (Y/N)"

        if ($SaveChoice -eq "Y") {

            Save-ReportToXlsx `
                -ReportName "Jira-Groups-Report" `
                -ReportData $ReportData
        }
        else {

            Pause-Script
        }
    }
    catch {

        Write-Host ""
        Write-Host "Failed to generate report!" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red

        Pause-Script
    }

    Show-FolderReportsMenu
}

# =========================================================
# SCRIPT START
# =========================================================

Select-Domain
Show-MainMenu