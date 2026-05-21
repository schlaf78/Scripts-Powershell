$gr = Get-ADGroupMember "<SOURCE GROUP NAME>"
Add-ADGroupMember -Identity "CN=<OU NAME>,OU=<OU NAME>,OU=<OU NAME>,DC=DOMAIN,DC=DOMAIN,DC=ru" -Members $gr
