Import-Module ActiveDirectory

#Script copies users in nested groups and make them as plain-added to the new group!


# Source and target groups
$SourceGroup = "<SOURCE GROUP>"
$TargetGroup = "<TARGET GROUP>"

# Get all users/computers from source group, resolving nested groups
$members = Get-ADGroupMember -Identity $SourceGroup -Recursive |
           Where-Object { $_.objectClass -in @("user","computer") }

# Add them as direct members into the target group
foreach ($m in $members) {
    try {
        Add-ADGroupMember -Identity $TargetGroup -Members $m -ErrorAction Stop
        Write-Host "Added $($m.SamAccountName) to $TargetGroup"
    }
    catch {
        Write-Warning "Skipping $($m.SamAccountName): $_"
    }
}

