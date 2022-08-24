$PrintEvent = Get-WinEvent `
    -LogName 'Microsoft-Windows-PrintService/Operational' `
    -FilterXPath '<QueryList><Query Id="0"><Select>*[System[EventID=307]]</Select></Query></QueryList>'

$PrintData = foreach ($PrintEventLine in $PrintEvent) {
    if ((GET-WinSystemLocale).Name -eq 'en-US') {
        $Message = $PrintEventLine.Message `
            -replace 'Document \d+, Print Document owned by ', '' `
            -replace ' on ', '|' `
            -replace ' was printed', '' `
            -replace ' through port ', '|' `
            -replace '\.  Size in bytes.+', '' `
            -replace ' ', ''
    }
    if ((GET-WinSystemLocale).Name -eq 'ru-RU') {
        $Message = $PrintEventLine[0].Message `
            -replace '^Документ \d, Печать документа, которым владеет ', '' `
            -replace ' на ', '|' `
            -replace '\, был распечатан', '' `
            -replace ' через порт ', '|' `
            -replace '\.  Размер в байтах.+', '' `
            -replace ' ', ''
    }
        
    #$Message    
    $Message -match '^([a-zA-Z.]+)\|([a-zA-Z0-9\-\\]+)\|([a-zA-Z0-9\-_\(\)]+)\|(.+$)' | Out-Null
    #Resolve-DnsName $Matches[2]
    New-Object PSObject -Property @{
        User            = $Matches[1]
        Workstation     = $Matches[2]
        PrinterName     = $Matches[3]
        PrinterPortName = $Matches[4]
    }
}

$PrintData = $PrintData | Select-Object -Property * -Unique
$DriverVersionList = Get-WmiObject Win32_PnPSignedDriver | Select-Object -Property FriendlyName,DriverVersion

$PrintData | ForEach-Object {

    Add-Member -InputObject $_ -NotePropertyName "WorkstationAddress" -NotePropertyValue (Resolve-DnsName $_.Workstation -ErrorAction SilentlyContinue).IPAddress
    Add-Member -InputObject $_ -NotePropertyName "PortAddress"        -NotePropertyValue (Get-PrinterPort -Name $_.PrinterPortName).PrinterHostAddress
    Add-Member -InputObject $_ -NotePropertyName "Printerlocation"    -NotePropertyValue (Get-Printer -Name $_.PrinterName).Location
    Add-Member -InputObject $_ -NotePropertyName "PrinterComment"     -NotePropertyValue (Get-Printer -Name $_.PrinterName).Comment
    Add-Member -InputObject $_ -NotePropertyName "PrinterShare"       -NotePropertyValue (Get-SmbShare | Where-Object Path -Like $($_.PrinterName+',LocalsplOnly')).name
    Add-Member -InputObject $_ -NotePropertyName "DriverName"         -NotePropertyValue (Get-Printer -Name $_.PrinterName).DriverName
    Add-Member -InputObject $_ -NotePropertyName "DriverVersion"      -NotePropertyValue ($DriverVersionList | Where-Object -Property FriendlyName -like $_.PrinterName).DriverVersion

}

$PrintData | Format-Table -AutoSize
#$PrintData | Export-Csv -Path C:\temp\Printer_expoert.csv -NoTypeInformation -Encoding UTF8 
