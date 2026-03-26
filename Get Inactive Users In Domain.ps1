#Script allows to get thwe report about inactive  users for last 90 days  in Windopws  Domain

$days = 90
$time = (Get-Date).AddDays(-$days)

Get-ADUser -Filter {Enabled -eq $true -and LastLogonDate -lt $time} `
    -Properties LastLogonDate |
    Select-Object Name, SamAccountName, LastLogonDate |
    Sort-Object LastLogonDate