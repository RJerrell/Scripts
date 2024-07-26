
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
Import-Module -Name "KC46S1000DRules" -Verbose -Force
Import-Module -Name "KC46DataManagement" -Verbose -Force
# *****************************************************************************************************
$drive = "F:"
$getHistory = $false
$commonRoot = "KC46"

[string[]] $PubList   = @("KC46", "ACS", "AMM" , "ARD", "FIM", "IPB", "LOAPS", "NDT", "SIMR", "SPCC","SSM","SWPM", "TC","WUC","WDM")
if($getHistory)
{
   
   $drive = "F:" 
}
else
{
    $reportNameCSV = "Current Counts of Data Modules - Release 10 - All Manuals.csv"
}

$KC46DataRoot = "$drive\KC46 Staging"
$source_BaseLocation = "$KC46DataRoot\production\Manuals"

$outputPath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$objs = @()

foreach ($pub in $PubList)
{
    $pathToDMs = "$source_BaseLocation\$pub\s1000d\sdllive\dmc*.xml"
    $1stVerCount = 0
    $2ndVerCount = 0
    $files = gci -Path $pathToDMs    
    $fcount = [int] $files.Count
    foreach ($file in $files)
    {
        $dmXML = New-Object System.Xml.XmlDocument
        $dmXML.Load($file.FullName)
        $1stVerificationValue = $dmXML.dmodule.identAndStatusSection.dmStatus.qualityAssurance.firstVerification.verificationType
        $2ndVerificationValue = $dmXML.dmodule.identAndStatusSection.dmStatus.qualityAssurance.secondVerification.verificationType

        if($1stVerificationValue.length -gt 0)
        {
            $1stVerCount += 1
        }
        if($2ndVerificationValue.length -gt 0)
        {
            $2ndVerCount += 1
        }
    }

    $obj = [pscustomobject][ordered]@{Manual=$pub;Total_DataModules=$fcount;firstVerification = $1stVerCount;secondVerification=$2ndVerCount}
    $objs += $obj 
    $obj = $null
}

# -- Sort properties for the report
$prop1 = @{Expression ='Total_DataModules'; Ascending=$false}

# Store the full data report
$objs.GetEnumerator() | Sort-Object -Property $prop1 | Export-Csv "$outputPath\$reportNameCSV" -NoTypeInformation

# *****************************************************************************************************

$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"