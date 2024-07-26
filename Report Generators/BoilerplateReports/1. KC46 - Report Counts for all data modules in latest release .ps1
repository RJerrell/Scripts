<#
Title: 
Author: Roger Jerrell
Date Created: Get-Date 
Purpose: 
Description of Operation: 
Description of Use:
#>
cls

# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'

#Get-Module -ListAvailable
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force
# *****************************************************************************************************
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$dmXml = new-object -TypeName S1000D.DataModule_401

$currentRelease = "Release 10"

$currReleasePath = "F:\KC46 Staging\Production\Manuals"

$exportFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs\Verification\Out"

$reportName = "Listing of all data modules - $currentRelease`.xlsx"

$masterList = @() # Carries all the values we need for this report
$ctr = [int] 0
$pmCTR = 0
[string[]] $Manuals   = @("ABDR","AMM","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WUC","WDM")

foreach ($Manual in $Manuals)
{

    $pmXml = New-Object System.Xml.XmlDocument

    # Path to the most current release of the CSDB
    $pms = gci -Path "$currReleasePath\$Manual" -Filter PMC*.XML -File -Recurse

    $pmXml.Load($pms[$pmCTR].FullName)
    $dmRefs = $pmXml.SelectNodes("/pm/content//dmRef")
    $pmEntries = $pmXml.SelectNodes("/pm/content//pmEntry")
    foreach ($dmRef in $dmRefs)
    {

        $pbTask = $dmRef.ParentNode
        
        $fileName = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
        $fileOnDisk = gci -Path "$currReleasePath\$Manual\S1000D\SDLLIVE\$fileName`*.xml"| Sort-Object -Descending | Select-Object -First 1

        $dmXml.ParseDM($fileOnDisk[0].FullName)
        # attrib -R $fileOnDisk[0].FullName
        Set-ItemProperty -Path $fileOnDisk[0].FullName -Name IsReadOnly -Value $false -Force -Verbose

        $cver = $dmXml.DmIssueInfo.issueNumber
        $infoName =  $dmXml.InfoName
        $techName =  $dmXml.TechName

        $type = $dmXml.DmType

        $pbTask = [string] $dmRef.href
        if($pbTask.length -eq 0)
        {
            $pbTask = "00-00-00"
        }
        $masterList += New-Object -TypeName PSObject -Property @{
                    Manual = $Manual;
                    Type = $type;
                    PBTask = $pbTask;
                    DMC = $fileName;
                    TechName=$techName;
                    InfoName = $infoName;                
                
        } | Select Manual,Type,PBTask,DMC,TechName,InfoName
        $ctr ++
        "Processing $ctr"
    }
}

# Export it all
Remove-Item -Path "$exportFolder\$reportName" -Force
$masterList | Export-XLSX -Path "$exportFolder\$reportName" -Header  Manual,Type,PBTask,DMC,TechName,InfoName -WorksheetName $currentRelease -Append
"$exportFolder\$reportName"

# *****************************************************************************************************

$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMilliseconds
"Process completed"