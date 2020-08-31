Try{
    Import-Module powershell-yaml -ErrorAction Stop
}
Catch{
    Write-Host "Unable to import PowerShell-Yaml Module. Trying to install." -ForegroundColor Yellow
    Try{
        #OverRide Default PS TLS settings
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
        Install-Module powershell-yaml -ErrorAction Stop
        }
    Catch{
        Write-Host "Unable to install PowerShell-Yaml Module. `n`
        `nPlease visit `"https://www.powershellgallery.com/packages/powershell-yaml`" and install module." -ForegroundColor Red
        Pause
        quit
        }
}
$YamlArray = @()
$Rules = Get-ChildItem .\*.yml -r | ?{$_.Attributes -ne "Directory"}

Foreach($Rule in $Rules){ 
    Write-Progress -Activity "Parsing $($Rules.Count) files" -CurrentOperation "$($Rule.Name)" -PercentComplete (($Rules.IndexOf($Rule) / $Rules.Count)*100)

    $YamlArrayT = "" | Select Name,Filename,ID,Description,Status,Tags,Product,Service,References
    $YamlRead = get-content $Rule.FullName | ConvertFrom-Yaml -Ordered

    $YamlArrayT.Name = $YamlRead.title
    $YamlArrayT.Filename = $Rule.Name    
    $YamlArrayT.ID = $YamlRead.id
    $YamlArrayT.Description = $YamlRead.description
    $YamlArrayT.status = $YamlRead.status
    $YamlArrayT.tags = $YamlRead.tags -join "`n"
    $YamlArrayT.product = $YamlRead.logsource.product -join "`n"
    $YamlArrayT.service = $YamlRead.logsource.service -join "`n"
    $YamlArrayT.references = $YamlRead.references -join "`n"


    $YamlArray += $YamlArrayT

}

Write-Progress -Activity "Parsing $($Rules.Count) files" -Completed

$YamlArray | Export-Csv -NoTypeInformation -Force .\Rules.csv
#$YamlArray | ConvertTo-Html | Out-File .\Rules.html
