
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

$pmcDmRefs = @()
$fileNames = @()

$manualPath = "C:\KC46 Staging\DEV\manuals\ABDR\s1000d\SDLLIVE"
$pmcs = gci -Path $manualPath -Filter PMC*.xml | Sort-Object -Descending | Select-Object -First 1
$pmc = $pmcs[0]
$pm = New-Object System.Xml.XmlDocument
$pm.Load("$manualPath\$pmc")

$dmCodes = $pm.SelectNodes("pm/content//dmCode")
#$dmCodes.count
$suffix = "_001-00_SX-US"

foreach ($dmCode in $dmCodes)
{
    $fn = Create-DMCFileNameFromDMCode -dmCode $dmCode
    $fn += "$suffix.xml"
    #$fn
    $pmcDmRefs += $fn
}

$files = gci -Path "$manualPath\dmc*.xml"

foreach ($file in $files)
{   
    $fileNames += $file.Name
}

$fnSorted = $fileNames  | Sort-Object

$pmEntries = $pmcDmRefs |Sort-Object

Compare-Object -ReferenceObject $pmEntries -DifferenceObject $fnSorted | Ft


# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"
