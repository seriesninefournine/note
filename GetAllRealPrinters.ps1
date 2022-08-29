Clear-host

$BackUpDriverPath = "c:\temp\printdrivers\"
$StartScript = Get-Date
Write-Host '------------------------------------------------------'
Write-Host '-------------Загрузка логов печати--------------------'

$PrintEvent = Get-WinEvent `
-LogName 'Microsoft-Windows-PrintService/Operational' `
-FilterXPath '<QueryList><Query Id="0"><Select>*[System[EventID=307]]</Select></Query></QueryList>'

Write-Host ''
Write-Host "Прошло: $([math]::Round($(($(Get-Date) - $StartScript).TotalSeconds),2)) секунд"
Write-Host '------------------------------------------------------'
Write-Host '--------------Парсинг логов печати--------------------'
Write-Host ''

$PrintData = foreach ($PrintEventLine in $PrintEvent) {
       if ($PrintEventLine.Message[0] -eq 'D') {
        $PrintEventLine.Message -match '^Document.+owned by ([a-zA-Z\-\.]+) on ([\da-zA-Z\-\._\\]+).+on ([ \da-zA-Zа-яА-Я\-\._\(\)]+) .+port ([\da-zA-Z\-\._]+)\..+$'  | Out-Null          
    }
    if ($PrintEventLine.Message[0] -eq 'Д') {
        $PrintEventLine.Message -match '^Документ.+владеет ([a-zA-Z\-\.]+) на ([\da-zA-Z\-\._\\]+).+на ([ \da-zA-Zа-яА-Я\-\._\(\)]+) .+порт ([\da-zA-Z\-\._]+)\..+$'  | Out-Null 
    }
       
       if (!(Get-Printer -Name $Matches[3] -ErrorAction SilentlyContinue)) {
             continue
       }
       
    New-Object PSObject -Property @{
             User            = $Matches[1]
        Workstation     = $Matches[2] -replace "\\", ''
        PrinterName     = $Matches[3]
        PrinterPortName = $Matches[4]

    }
}
$PrintData = $PrintData | select * -Unique

Write-Host ''
Write-Host "Прошло: $([math]::Round($(($(Get-Date) - $StartScript).TotalSeconds),2)) секунд"
Write-Host '------------------------------------------------------'
Write-Host '-------------Загрузка базы драйверов------------------'

$DriverVersionList = Get-WindowsDriver -Online -All | Select-Object -Property OriginalFileName,Version

Write-Host ''
Write-Host "Прошло: $([math]::Round($(($(Get-Date) - $StartScript).TotalSeconds),2)) секунд"
Write-Host '------------------------------------------------------'
Write-Host '----------Выборка дополнительных данных---------------'

$PrintData | Foreach {

    Add-Member -InputObject $_ -NotePropertyName "WorkstationAddress" -NotePropertyValue (Resolve-DnsName $_.Workstation -ErrorAction SilentlyContinue).IPAddress
    Add-Member -InputObject $_ -NotePropertyName "PortAddress"        -NotePropertyValue (Get-PrinterPort -Name $_.PrinterPortName).PrinterHostAddress
    Add-Member -InputObject $_ -NotePropertyName "Printerlocation"    -NotePropertyValue (Get-Printer -Name $_.PrinterName).Location
    Add-Member -InputObject $_ -NotePropertyName "PrinterComment"     -NotePropertyValue (Get-Printer -Name $_.PrinterName).Comment
    Add-Member -InputObject $_ -NotePropertyName "PrinterShare"       -NotePropertyValue (Get-Printer -Name $_.PrinterName).ShareName
    Add-Member -InputObject $_ -NotePropertyName "DriverName"         -NotePropertyValue (Get-Printer -Name $_.PrinterName).DriverName
    Add-Member -InputObject $_ -NotePropertyName "DriverVersion"      -NotePropertyValue ($DriverVersionList | Where-Object -Property OriginalFileName -Like ((Get-PrinterDriver | where name -Like  ((Get-Printer -Name $_.PrinterName).DriverName  -replace "\[", '`[' -replace ']', '`]') | Select-Object InfPath)[0]).InfPath).Version
       $DriverInf = ($DriverVersionList | Where-Object -Property OriginalFileName -Like ((Get-PrinterDriver | where name -Like  ((Get-Printer -Name $_.PrinterName).DriverName  -replace "\[", '`[' -replace ']', '`]') | Select-Object InfPath)[0]).InfPath).OriginalFileName
    Add-Member -InputObject $_ -NotePropertyName "DriverPath"              -NotePropertyValue ($BackUpDriverPath+((Get-ChildItem($DriverInf)).directory).Name)
       Add-Member -InputObject $_ -NotePropertyName "DriverFullInf"        -NotePropertyValue $DriverInf
       Add-Member -InputObject $_ -NotePropertyName "DriverInf"            -NotePropertyValue ((Get-ChildItem($DriverInf)).Name)
}

Write-Host ''
Write-Host "Прошло: $([math]::Round($(($(Get-Date) - $StartScript).TotalSeconds),2)) секунд"
Write-Host '------------------------------------------------------'
Write-Host '----------------Выгрузка драйверов--------------------'

$PrintData | select DriverFullInf -Unique | Foreach {
       cp ((Get-ChildItem($_.DriverFullInf)).directoryname) $BackUpDriverPath -Recurse -Force
}

$PrintData | select * -ExcludeProperty DriverFullInf | Export-Csv -Path "C:\temp\Printer_export_$env:COMPUTERNAME.csv" -NoTypeInformation -Encoding UTF8 -Force
$PrintData | select * -ExcludeProperty DriverFullInf | Out-GridView
Write-Host ''
Write-Host "Прошло: $([math]::Round($(($(Get-Date) - $StartScript).TotalSeconds),2)) секунд"
Write-Host '------------------------------------------------------'
Write-Host '-----------------Работа завершена---------------------'

