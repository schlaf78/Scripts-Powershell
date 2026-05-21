# Clean Previous Output
Clear-Host

# =========================
# Functions
# =========================

# Active Directory Single Group
Function ActiveDirectorySingleGroup {

    $GroupName = Read-Host -Prompt "Enter AD Group name"

    "`n"
    "Members of group: $GroupName"

    Get-ADGroupMember $GroupName |
        Select-Object Name, SamAccountName
}


# Atlassian Confluence Group Block
Function AtlassianConfluence {

    $ConfluenceGroupInventory = Read-Host -Prompt "Enter Confluence Group name"

    # Administrator Group
    "`n"
    "$ConfluenceGroupInventory-adm members"

    Get-ADGroupMember "${ConfluenceGroupInventory}-adm" |
        Select-Object Name, SamAccountName

    # Read-Write Group
    "`n"
    "$ConfluenceGroupInventory-rw members"

    Get-ADGroupMember "${ConfluenceGroupInventory}-rw" |
        Select-Object Name, SamAccountName

    # Read-Only Group
    "`n"
    "$ConfluenceGroupInventory-ro members"

    Get-ADGroupMember "${ConfluenceGroupInventory}-ro" |
        Select-Object Name, SamAccountName
}


# Atlassian Jira Group Block
Function AtlassianJira {

    $JiraGroupInventory = Read-Host -Prompt "Enter Jira Group name"

    # Administrator Group
    "`n"
    "$JiraGroupInventory-adm members"

    Get-ADGroupMember "${JiraGroupInventory}-adm" |
        Select-Object Name, SamAccountName

    # Analyst Group
    "`n"
    "$JiraGroupInventory-anl members"

    Get-ADGroupMember "${JiraGroupInventory}-anl" |
        Select-Object Name, SamAccountName

    # Developer Group
    "`n"
    "$JiraGroupInventory-dev members"

    Get-ADGroupMember "${JiraGroupInventory}-dev" |
        Select-Object Name, SamAccountName

    # User Group
    "`n"
    "$JiraGroupInventory-usr members"

    Get-ADGroupMember "${JiraGroupInventory}-usr" |
        Select-Object Name, SamAccountName
}


# =========================
# Menu
# =========================

Write-Host "Select Service To Get Group Membership"

$Menu = Read-Host "
(1) Active Directory Single Group
(2) Atlassian Confluence Typical Pack Of Groups
(3) Atlassian Jira Typical Pack Of Groups
"

Switch ($Menu)
{
    "1" { ActiveDirectorySingleGroup }
    "2" { AtlassianConfluence }
    "3" { AtlassianJira }
    Default {
        Write-Host "`nInvalid selection." -ForegroundColor Red
    }
}