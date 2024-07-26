
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
Import-Module -Name "KC46Common" -Verbose -Force
# *****************************************************************************************************
cls
$objs =  @()
$global:environment = "Production"  # *************   Override to Production  ************#

# Where the source S1000D data is located that will eventually become an IETM
$global:KC46DataRoot = "f:\KC46 Staging"
$commonRoot = "KC46"
# BASIC FOLDER STRUCTURES THAT ARE PRETTY HARD CODED AND EXPECTED TO BE THERE.  IF THESE CHANGE, IT ALL FALLS DOWN.
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"
$outputfolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs"
$reportName = "KC46 - Quality Assurance Report - S1000D.csv"

[string[]] $PubList   = @("KC46", "ACS", "ARD", "AMM", "FIM", "IPB", "LOAPS",  "NDT", "SIMR", "SPCC",  "SRM", "SSM", "SWPM", "WUC", "WDM")
foreach ($pub in $PubList)
   {       
       # $files = gci -Path "$source_BaseLocation\$pub\S1000D\SDLLIVE\DMC*.*"
       $files = gci -Path "$source_BaseLocation\$pub\S1000D\SDLLIVE\DMC*.*"
       
       $fCounter = 1
       
       foreach ($file in $files)
       {
            Write-Progress -Activity “Processing the $pub folder ...” -status “Finding file $file” -percentComplete ($fCounter / $files.count*100)
           
            $FN = $file.Name
            
            $FFN = $file.FullName
            
            $FNShort = $file.Name.Replace(".xml", "")

            $dm = [xml] (Get-Content -Path $FFN)
            # /dmodule/identAndStatusSection/dmAddress/dmIdent/issueInfo
            $issueNumber = $dm.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.issueNumber.ToString()
             
            # /dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle/techName
            $techName = $dm.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName.ToString()
            
            # /dmodule/identAndStatusSection/dmStatus/qualityAssurance/firstVerification
            $firstVerification = [string] $dm.dmodule.identAndStatusSection.dmStatus.qualityAssurance.firstVerification.verificationType
              
            # # /dmodule/identAndStatusSection/dmStatus/qualityAssurance/secondVerification
            $secondVerification = [string] $dm.dmodule.identAndStatusSection.dmStatus.qualityAssurance.secondVerification.verificationType
            
            $obj = [pscustomobject][ordered]@{Manual=$pub;TechName=$techName;DMC=$FNShort;IssueNumber=$issueNumber;firstVerification=$firstVerification;secondVerification=$secondVerification;}
            $objs += $obj
            
            $fCounter ++ 
       }       
   }
    
   if(!(Test-Path -Path $outputfolder))
   {
    md $outputfolder
   }

# Define the sort properties for the report
$prop1 = @{Expression='Manual'; Ascending=$true }
$prop2 = @{Expression='DMC'; Ascending=$true }


# Store the report
$objs.GetEnumerator() | Sort-Object -Property $prop1, $prop2| Export-Csv "$outputfolder\$reportName" -NoTypeInformation

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"
