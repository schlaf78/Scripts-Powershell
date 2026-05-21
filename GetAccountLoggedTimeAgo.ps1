# Import Active Directory module (if not already loaded)
Import-Module ActiveDirectory

# Define the OU to search
$OU = "OU=<OU NAME>,OU=<OU NAME>,OU=<OU NAME>,DC=DOMAIN,DC=DOMAIN,DC=ru"

# Get the date threshold (30 days ago)
$ThresholdDate = (Get-Date).AddDays(-30)

# Get user accounts from the specified OU
$Users = Get-ADUser -SearchBase $OU -Filter * -Properties LastLogonDate, LockedOut

# Process each user and check last logon & lock status
$Results = $Users | Where-Object {
    $_.LastLogonDate -lt $ThresholdDate -or -not $_.LastLogonDate
} | Select-Object Name, SamAccountName, LastLogonDate, 
    @{Name="DaysSinceLastLogon"; Expression={if ($_.LastLogonDate) {(New-TimeSpan -Start $_.LastLogonDate -End (Get-Date)).Days} else {"Never logged in"}}}, 
    LockedOut

# Define the export path
$ExportPath = "C:\Users\<USERNAME>\Downloads\InactiveUsers.csv"

# Create Excel COM Object
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

# Create a new workbook
$Workbook = $Excel.Workbooks.Add()
$Sheet = $Workbook.Worksheets.Item(1)

# Add headers
$Headers = @("Name", "SamAccountName", "LastLogonDate", "DaysSinceLastLogon", "LockedOut")
for ($i = 0; $i -lt $Headers.Length; $i++) {
    $Sheet.Cells.Item(1, $i + 1) = $Headers[$i]
    $Sheet.Cells.Item(1, $i + 1).Font.Bold = $true
}

# Add user data
$Row = 2
foreach ($User in $Results) {
    $Sheet.Cells.Item($Row, 1) = $User.Name
    $Sheet.Cells.Item($Row, 2) = $User.SamAccountName
    $Sheet.Cells.Item($Row, 3) = $User.LastLogonDate
    $Sheet.Cells.Item($Row, 4) = $User.DaysSinceLastLogon
    $Sheet.Cells.Item($Row, 5) = $User.LockedOut
    $Row++
}

# Auto-fit columns for better readability
$Sheet.Columns.AutoFit()

# Save and close the workbook
$Workbook.SaveAs($ExportPath, 51)  # 51 = Excel Open XML Format (.xlsx)
$Workbook.Close()
$Excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Sheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null

