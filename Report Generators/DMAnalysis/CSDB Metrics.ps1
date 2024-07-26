
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
$objs =  @()
$pathToCSDB = "D:\Shared\IDE cd sets\2017-01-20-14-18-01 - Non CDRL January 2017 - Release 5\CSDB\Manuals"

$reportTitleSuffix = (Get-Date -Format "s").Replace(":","-")
$reportTitle = "KC46 Tanker S1000D CSDB Metrics" + $reportTitleSuffix + ".csv"
$outputPath = "C:\KC46 Staging\Scripts\Report Generators\Outputs"
[string[]] $PubList   = @("KC46","ACS","AMM","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WUC","WDM")

$totalDMs = 0
$totalIlls = 0

foreach ($Pub in $PubList)
{
    $path2DM = "$pathToCSDB\$Pub\S1000D\SDLLIVE"
    $path2Illustrations = "$pathToCSDB\$Pub\Illustrations\Illustrations"
    $filesDM = gci -Path $path2DM -Recurse -Filter DMC*.xml
    $filesILL = gci -Path $path2DM -Recurse -Filter *.cgm
    $DocArray = $null
    foreach ($file in $filesDM)
    {
        $DocArray = $null
        $tName = ""
        $iName = ""
        $docType = ""
        $DocArray = Get-TechNameInfoNameFromDMC -fileName $file.FullName
        $FN = $file.Name.ToUpper().Replace(".XML", "")
        $tName = $DocArray[0].Trim()
        $iName = $DocArray[1].Trim()
        $docType = $DocArray[2].Trim()
        $FNArray = Get-FileNameArray $file.Name
        $iCode = $FNArray[7].Substring(0,3)
        $obj = [pscustomobject][ordered]@{Manual=$Pub;DMC=$FN;Type=$docType;TechName=$tName;InfoName=$iName;InfoCode=$iCode;}
        $objs += $obj
        $obj=$null
    }

}

$prop1 = @{Expression='Manual'; Ascending=$true }
$prop2 = @{Expression='DMC'; Ascending=$true }
$objs.GetEnumerator() | Sort-Object -Property $prop1,$prop2  | Export-Csv "$outputPath\$reportTitle" -NoTypeInformation
"The report awaits you: " + $outputPath

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"