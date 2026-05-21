param(
    [string]$UserA,
    [string]$UserB
)



# Требуется модуль ActiveDirectory
Import-Module ActiveDirectory -ErrorAction Stop

function Get-GroupPathsRec {
    param(
        [string]$GroupDN,
        [hashtable]$Cache,
        [System.Collections.Generic.HashSet[string]]$Visited
    )

    if ($Visited.Contains($GroupDN)) { return @() }
    $Visited.Add($GroupDN) | Out-Null

    if (-not $Cache.ContainsKey($GroupDN)) {
        try {
            $g = Get-ADGroup -Identity $GroupDN -Properties MemberOf,Name
        } catch {
            $g = [pscustomobject]@{ Name=$GroupDN; MemberOf=@() }
        }
        $Cache[$GroupDN] = $g
    } else {
        $g = $Cache[$GroupDN]
    }

    $name = $g.Name
    $paths = @()

    if ($g.MemberOf) {
        foreach ($p in $g.MemberOf) {
            $sub = Get-GroupPathsRec -GroupDN $p -Cache $Cache -Visited $Visited
            foreach ($s in $sub) {
                $paths += ,(@($s) + $name)
            }
        }
    }
    if (-not $paths) { $paths += ,(@($name)) }

    $Visited.Remove($GroupDN) | Out-Null
    return $paths
}

function Get-UserGroupPaths {
    param([string]$UserId)

    $u = Get-ADUser -Identity $UserId -Properties MemberOf,Name
    $cache = @{}
    $allPaths = @()
    $groups = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($dn in $u.MemberOf) {
        $visited = New-Object 'System.Collections.Generic.HashSet[string]'
        $paths = Get-GroupPathsRec -GroupDN $dn -Cache $cache -Visited $visited
        foreach ($p in $paths) {
            $allPaths += ,$p
            $p | ForEach-Object { $null = $groups.Add($_) }
        }
    }

    return @{
        Paths=$allPaths
        Groups=$groups
        Name=$u.Name
        Sam=$u.SamAccountName
    }
}

function Build-Tree {
    param([array]$Paths)

    $tree = @{ name=""; children=@() }

    function EnsureChild($parent, $childName) {
        $ex = $parent.children | Where-Object { $_.name -eq $childName }
        if ($ex) { return $ex }
        $node = @{ name=$childName; children=@() }
        $parent.children += ,$node
        return $node
    }

    foreach ($p in $Paths) {
        $parent=$tree
        foreach ($seg in $p) {
            $parent = EnsureChild $parent $seg
        }
    }
    return $tree
}

function Collect-Lines {
    param(
        $Node,
        $Indent,
        [System.Collections.Generic.HashSet[string]]$Common,
        [System.Collections.Generic.HashSet[string]]$Other
    )

    $lines = @()
    foreach ($c in $Node.children) {
        $name = $c.name
        if ($Common.Contains($name)) {
            $color = "Green"
        } elseif ($Other.Contains($name)) {
            $color = "Yellow"
        } else {
            $color = "Red"
        }
        $line = @{ Text=(" " * $Indent + "└─ " + $name); Color=$color }
        $lines += ,$line
        $lines += Collect-Lines -Node $c -Indent ($Indent+3) -Common $Common -Other $Other
    }
    return $lines
}

function Show-SideBySide {
    param($LeftTree, $RightTree, $Common, $LeftGroups, $RightGroups)

    $leftLines  = Collect-Lines -Node $LeftTree  -Indent 0 -Common $Common -Other $RightGroups
    $rightLines = Collect-Lines -Node $RightTree -Indent 0 -Common $Common -Other $LeftGroups

    $max = [Math]::Max($leftLines.Count, $rightLines.Count)
    $width = ($leftLines | ForEach-Object { $_.Text.Length } | Measure-Object -Maximum).Maximum
    if (-not $width) { $width = 30 }

    for ($i=0; $i -lt $max; $i++) {
        $l = if ($i -lt $leftLines.Count) { $leftLines[$i] } else { @{Text="";Color="White"} }
        $r = if ($i -lt $rightLines.Count) { $rightLines[$i] } else { @{Text="";Color="White"} }

        Write-Host ($l.Text.PadRight($width+4)) -ForegroundColor $l.Color -NoNewline
        Write-Host "│   " -NoNewline
        Write-Host $r.Text -ForegroundColor $r.Color
    }
}

# --- main ---

if (-not $UserA) { $UserA = Read-Host "Введите первого пользователя" }
if (-not $UserB) { $UserB = Read-Host "Введите второго пользователя" }

$left  = Get-UserGroupPaths $UserA
$right = Get-UserGroupPaths $UserB

$common = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($g in $left.Groups) { if ($right.Groups.Contains($g)) { $null=$common.Add($g) } }

$leftTree  = Build-Tree $left.Paths
$rightTree = Build-Tree $right.Paths

Write-Host ("Пользователь A: {0} ({1})" -f $left.Name,$left.Sam) -ForegroundColor Cyan -NoNewline
Write-Host (" " * 20 + "Пользователь B: {0} ({1})" -f $right.Name,$right.Sam) -ForegroundColor Cyan
Show-SideBySide -LeftTree $leftTree -RightTree $rightTree -Common $common -LeftGroups $left.Groups -RightGroups $right.Groups
