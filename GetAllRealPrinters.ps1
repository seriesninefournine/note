Clear-host
$PrintEvent = Get-WinEvent `
-LogName 'Microsoft-Windows-PrintService/Operational' `
-FilterXPath '<QueryList><Query Id="0"><Select>*[System[EventID=307]]</Select></Query></QueryList>'
$lang = $null
if (!$lang) {$lang = (Get-ItemProperty 'HKCU:\Control Panel\Desktop' PreferredUILanguages -ErrorAction SilentlyContinue ).PreferredUILanguages[0]}
if (!$lang) {$lang = (Get-Culture).Name}

$PrintData = foreach ($PrintEventLine in $PrintEvent) {
        if ($lang -eq 'en-US') {
        $Message = $PrintEventLine.Message `
            -replace '^Document \d+\, Print Document owned by ', '' `
            -replace ' on ', '|' `
            -replace ' was printed', '' `
            -replace ' through port ', '|' `
            -replace '\.  Size in bytes.+', '' `
            -replace ' ', ''
    }
    if ($lang -eq 'ru-RU') {
        $Message = $PrintEventLine[0].Message `
            -replace '^Документ \d+\, Печать документа, которым владеет ', '' `
            -replace ' на ', '|' `
            -replace '\, был распечатан', '' `
            -replace ' через порт ', '|' `
            -replace '\.  Размер в байтах.+', '' `
            -replace ' ', ''
    }

    $Message -match '^([a-zA-Z.]+)\|([a-zA-Z0-9\-\\]+)\|([a-zA-Z0-9\-_\(\)]+)\|(.+$)' | Out-Null
    New-Object PSObject -Property @{
                User            = $Matches[1]
        Workstation     = $Matches[2]
        PrinterName     = $Matches[3]
        PrinterPortName = $Matches[4]

    }
}

$PrintData = $PrintData | select * -Unique
$DriverVersionList = Get-WindowsDriver -Online -All | Select-Object -Property OriginalFileName,Version

$PrintData | Foreach {

    Add-Member -InputObject $_ -NotePropertyName "WorkstationAddress" -NotePropertyValue (Resolve-DnsName $_.Workstation -ErrorAction SilentlyContinue).IPAddress
    Add-Member -InputObject $_ -NotePropertyName "PortAddress"        -NotePropertyValue (Get-PrinterPort -Name $_.PrinterPortName).PrinterHostAddress
    Add-Member -InputObject $_ -NotePropertyName "Printerlocation"    -NotePropertyValue (Get-Printer -Name $_.PrinterName).Location
    Add-Member -InputObject $_ -NotePropertyName "PrinterComment"     -NotePropertyValue (Get-Printer -Name $_.PrinterName).Comment
    Add-Member -InputObject $_ -NotePropertyName "PrinterShare"       -NotePropertyValue (Get-Printer -Name $_.PrinterName).ShareName
    Add-Member -InputObject $_ -NotePropertyName "DriverName"         -NotePropertyValue (Get-Printer -Name $_.PrinterName).DriverName
    Add-Member -InputObject $_ -NotePropertyName "DriverVersion"      -NotePropertyValue ($DriverVersionList | Select-Object  -Property OriginalFileName,Version | Where-Object -Property OriginalFileName -Like ((Get-PrinterDriver -name (Get-Printer -Name $_.PrinterName).DriverName | Select-Object InfPath)[0]).InfPath).Version

}

$PrintData | Export-Csv -Path "C:\temp\Printer_export_$env:COMPUTERNAME.csv" -NoTypeInformation -Encoding UTF8 -Force
$PrintData | Out-GridView
