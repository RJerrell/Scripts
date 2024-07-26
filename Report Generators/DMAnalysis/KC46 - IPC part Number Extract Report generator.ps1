
<#

Title: 
Author: Roger Jerrell
Date Created: Get-Date 
Purpose: 
Description of Operation: 
Description of Use:

#>
cls
$sd = Get-Date
$ErrorActionPreference = "SilentlyContinue"
$error.Clear()
$a = (Get-Host).PrivateData
$a.WarningBackgroundColor = "Yellow"
$a.WarningForegroundColor = "Black"
$rn = "<br/>"
$a.ErrorBackgroundColor = "red"
$a.ErrorForegroundColor = "White"

$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
$env:PSModulePath

#Get-Module -ListAvailable
Import-Module -Name "KC46Common" -Verbose -Force

# *****************************************************************************************************
# /dmodule/content/illustratedPartsCatalog/catalogSeqNumber/itemSequenceNumber/partNumber
$pathToDatamodules = "F:\KC46 Staging\Production\Manuals\IPB\S1000D\SDLLIVE\DMC*.XML"

$dms = gci -Path $pathToDatamodules

$reportpath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$rptBaseName = "IPC Part Number Listing"
$rptSuffix1 = "- Sorted by PN"
$rptSuffix2 = "- Sorted by DMC"

$allParts =  @()
$objsec = $null
$dmXML = New-Object System.Xml.XmlDocument
foreach ($dm in $dms)
{
    $dmXML.Load($dm.FullName)
    $pnS = $dmXML.SelectNodes("/dmodule/content/illustratedPartsCatalog/catalogSeqNumber/itemSequenceNumber/partNumber")
    foreach ($pn in $pnS)
    {
        $objsec = [pscustomobject][ordered]@{Part_Number=$pn;DMC=$dm.Name;}
        $allParts += $objsec
        $objsec = $null        
     }    
}
$prop1 = @{Expression='Part_Number'; Ascending=$true }
$prop2 = @{Expression='DMC'; Ascending=$true }
$allParts.GetEnumerator() | Sort-Object -Property $prop1, $prop2 | Export-Csv "$reportpath\$rptBaseName $rptSuffix1.csv" -NoTypeInformation -Encoding UTF8
$allParts.GetEnumerator() | Sort-Object -Property $prop2, $prop1 | Export-Csv "$reportpath\$rptBaseName $rptSuffix2.csv" -NoTypeInformation -Encoding UTF8

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"