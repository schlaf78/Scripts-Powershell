function Get-Size {
    param([string]$pth)

    $items = Get-ChildItem -Path $pth -Recurse -File -ErrorAction SilentlyContinue
    $totalSize = 0

    foreach ($item in $items) {
        $totalSize += $item.Length
    }

    return "{0:n2}" -f ($totalSize / 1MB)
}

# Output header
"D:\<FOLDER>" | Out-File -FilePath D:\filesize.txt -Encoding UTF8

# Get direct subfolders
$subfolders = Get-ChildItem -Path "D:\<FOLDER>" -Directory

foreach ($folder in $subfolders) {
    $sizeMB = Get-Size $folder.FullName
    "$($folder.FullName) - $sizeMB MB" | Out-File -FilePath D:\filesize.txt -Append -Encoding UTF8
}
