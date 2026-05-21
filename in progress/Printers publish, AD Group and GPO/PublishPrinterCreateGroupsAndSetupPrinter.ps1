###################################################################
#Order to Run
#0. Run under the administartor ONLY!

#1. Run this script on PRN-4 Locally: Replace CSV file  and change  
#Driver name and Starting Counter, Change Export CSV file 

#2. Run 2nd script with GPO on ADN or any with RSAT

#3. fix GPO after

#4. add ad group to printer
###################################################################

$CsvFilePath = "C:\Users\<USERNAME>\Downloads\printers 8.csv"
$DriverName = "TSC DA220"
#Possible DriverName: "Pantum M6550NW Series", "TSC DA220", "Kyocera ECOSYS M2640idw KX"
$OutputList = @()
$Counter = 238

$Printers = Get-Content -Path $CsvFilePath
$OU = "OU=printers,OU=Группы,DC=<DOMAIN>,DC=<DOMAIN>,DC=ru"

foreach ($PrinterDNSName in $Printers) {
    $PortName = $PrinterDNSName
    $PrinterName = "Printer #$Counter"
    $GroupName = $PrinterName

           # Create AD Group
    if (-not (Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue)) {
        New-ADGroup -Name $GroupName -Path $OU -GroupScope Global -PassThru
    }

    if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
        Add-PrinterPort -Name $PortName -PrinterHostAddress $PrinterDNSName
    }

    if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
        Add-Printer -Name $PrinterName -PortName $PortName -DriverName $DriverName `
            -Shared -ShareName $PrinterName `
            -Location $PrinterDNSName -Comment $PrinterDNSName
    } else {
        Set-Printer -Name $PrinterName -Location $PrinterDNSName -Comment $PrinterDNSName
    }

    $OutputList += [PSCustomObject]@{
        PrinterName = $PrinterName
        DNSName     = $PrinterDNSName
        GroupName   = $GroupName
    }

    $Counter++
    if ($Counter -gt 237) { break }
}

# Export to JSON or CSV
$OutputList | Export-Csv -Path "C:\Users\<USER>\Downloads\printers-export 8.csv" -NoTypeInformation
