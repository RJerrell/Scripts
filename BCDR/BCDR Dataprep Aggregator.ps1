
<#
Title: 
Author: Roger Jerrell
Date Created: 
Purpose: 
Description of Operation: 
Description of Use: Point the input folder at each IPR folder and get the individual IPR dm listings and then aggregate them all into a single dm listing.

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
Import-Module -Name "KC46Common" -Verbose -Force
# *****************************************************************************************************

[string[]] $PubList   = @("KC46", "ACS","AMM","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WUC","WDM")
[string[]] $ManualS   = @("AMM","IPB","SIMR","SSM","TC","WDM","WUC")
#[string[]] $PubList   = @("ARD", "ACS")
$dmData_Rootpath = "D:\Shared\BCDR\IPR\100 Percent IPR - Nov 2018"
$iprListings = gci -Path $dmData_Rootpath -Filter "KC46_*.XML" -Recurse -File
$doc = New-Object System.Xml.XmlDocument

$doc.CreateProcessingInstruction("true")
$dec = $doc.CreateXmlDeclaration("1.0","UTF-8",$null)
$doc.AppendChild($dec)
$c = $doc.CreateComment("Aggregate XML Listing of all IPRs for KC46 Tanker.  Root folder is " + $dmData_Rootpath)
$doc.AppendChild($c)
$ROOT = $doc.CreateNode("element","dmodule",$null)
$NULL = $doc.AppendChild($ROOT)
$iprXml = New-Object System.Xml.XmlDocument

foreach ($iprListing in $iprListings)
{
    $iprXml.Load($iprListing.FullName)
    $dmItems = $iprXml.SelectNodes("//dmitem")
    foreach ($dmItem in $dmItems)
    {
        $newNode = $doc.ImportNode($dmItem, $true)
        $null = $ROOT.AppendChild($newNode)
    }
}

$doc.DocumentElement.AppendChild($ROOT)
$doc.OuterXml
$doc.Save("$dmData_Rootpath\full_dm_listing.xml")
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"