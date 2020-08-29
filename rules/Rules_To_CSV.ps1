
Try{
    Import-Module powershell-yaml -ErrorAction Stop
}
Catch{
    Write-Host "Unable to import PowerShell-Yaml Module. Trying to install." -ForegroundColor Yellow
    Try{
        Install-Module powershell-yaml
        }
    Catch{
        Write-Host "Unable to install PowerShell-Yaml Module. `n`
        `nPlease visit `"https://www.powershellgallery.com/packages/powershell-yaml`" and install module." -ForegroundColor Red
        Pause
        Exit
        }
}
$YamlArray = @()
$Rules = Get-ChildItem .\*.yml -r | ?{$_.Attributes -ne "Directory"}

Foreach($Rule in $Rules){ 
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

$YamlArray | Export-Csv -NoTypeInformation -Force .\Rules.csv
#$YamlArray | ConvertTo-Html | Out-File .\Rules.html

