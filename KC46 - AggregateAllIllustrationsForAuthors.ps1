<#
Title: 
Author: Roger Jerrell
Date Created: Get-Date 
Purpose: Put all the illustrations for all the books into 1 big folder for the authors convenience.
Description of Operation: Copy all the illustrations to a single destination folder
Description of Use: Authors will use the illustrations in writing BDS manuals that use illustrations in other books.

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
#Import-Module -Name "KC46Common" -Verbose -Force
# *****************************************************************************************************
Function Set-Illustrations
{
    [string[]] $PubList   = @("KC46","ACS","AMM","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WUC","WDM")
    #[string[]] $PubList   = @("KC46","ACS","AMM","ARD")
    $basepath = "C:\KC46 Staging\Production\Manuals"

    $targetPath =  "D:\AllTankerIllustrations"
    if(!(Test-Path -Path $targetPath))
    {
        md $targetPath
    }

    foreach($Pub in $PubList)
    {
        $sourcePath = "$basepath\$Pub\ILLUSTRATIONS\ILLUSTRATIONS\*.*"
        $sourceGraphics = gci -Path $sourcePath -Filter ICN*.* -Exclude *.txt
        foreach($sourceGraphic in $sourceGraphics)
        {
            $name = $sourceGraphic.Name
            if(!([System.IO.File]::Exists("$targetPath\$name")))
            {
                Copy-Item -Path $sourceGraphic.FullName -Destination $targetPath -Force
            }
        }
    }
}

Remove-Item -Path "D:\AllTankerIllustrations" -Recurse -Force -Verbose

Set-Illustrations

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.Days
"Total Hours to complete:`t" + $x.Hours
"Total Minutes to complete:`t" + $x.Minutes
"Total Seconds to complete:`t" + $x.Seconds
"Process completed"