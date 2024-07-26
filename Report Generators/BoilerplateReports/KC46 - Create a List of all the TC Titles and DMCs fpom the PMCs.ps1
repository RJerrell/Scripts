<#
Title: KC46 - Set USAF SecondVerification Values
Author: Roger Jerrell
Date Created: 09/11/2017
Purpose: Set the values for the secondVerification tags to the values supplied by the USAF
Description of Operation: Use a white list of values to drive the values
Description of Use: 
    - Read in an Excel file supplied by the CUSTOMER and set the secondVerification element to the value supplied.
    - If the IssueNumber for a DMC is lower than the Verified value in the spreadsheet, the secondVerification element must be removed.
    - A text log will be kept and furnished to the CUSTOMER for any discrpencies.
#>
cls

$sd = Get-Date

$ErrorActionPreference = "SilentlyContinue"

$error.Clear()

$a = (Get-Host).PrivateData

$a.WarningBackgroundColor = "Yellow"
$a.WarningForegroundColor = "Black"
$a.ErrorBackgroundColor = "red"
$a.ErrorForegroundColor = "White"

$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force

# *****************************************************************************************************

#$PathToData = "F:\KC46 Staging\Production\Manuals"
$PathToData = "c:\KC46 Staging\Production\Manuals"
$exportFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs"
$reportName = "Task Card List Based on the TC PMC for the current revision.xlsx"


[string[]] $PubList   = @("TC")

$taskCardList = @()

$pmXml = New-Object System.Xml.XmlDocument
$dmXml = New-Object System.Xml.XmlDocument

foreach ($Pub in $PubList)
{
    $pathBase = "$PathToData\$Pub\S1000D\SDLLIVE"

    $path = "$PathToData\$Pub\S1000D\SDLLIVE\PMC*.XML"

    $pms = gci -Path $path -File

    foreach ($pm in $pms)
    {
        # /pm/content/pmEntry/pmEntry/pmEntry/dmRef
        $pmXml.Load($pm[0].FullName)
        $dmRefs = $pmXml.SelectNodes("/pm/content//dmRef")
        $pmEntries = $pmXml.SelectNodes("/pm/content//pmEntry")

        foreach ($dmRef in $dmRefs)
        {
            $pbTask = $dmRef.ParentNode
            #/dmodule/identAndStatusSection/dmAddress/dmIdent/issueInfo
            $fileName = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
            $fileOnDisk = gci -Path "$pathBase\$fileName`*.xml"
            $dmXml.Load($fileOnDisk[0].FullName)
            #/dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle/techName

            $techName = $dmXml.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
            $taskCardNumberArray = $techName.Split("-") 
            $taskCardNumber = ""
            for ($i = 0; $i -lt 4; $i++)
            { 
                if($i -lt 3)
                {                
                    $taskCardNumber += [string] $taskCardNumberArray[$i] + "-"
                }
                else
                {
                    $taskCardNumber += [string] $taskCardNumberArray[$i]
                }

            }

            $taskCardList += New-Object -TypeName PSObject -Property @{
                TaskCardNumber = $taskCardNumber
                TechName = $techName                          
                DMC = $fileName
            } | Select DMC, TechName,TaskCardNumber
        }
    }
}
CLS
$taskCardList2 = @()
foreach  ($Pub in $PubList)
{
    $pathBase = "$PathToData\$Pub\S1000D\SDLLIVE"

    $path = "$PathToData\$Pub\S1000D\SDLLIVE\DMC*.XML"

    $DMS = gci -Path $path -File
    foreach ($DM in $DMS)
    {
        #$DMS[0].FullName
        $dmXml.Load($DM.FullName)
        $techName = $dmXml.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
        $taskCardNumberArray = $techName.Split("-") 
        $taskCardNumber = ""
        for ($i = 0; $i -lt 4; $i++)
        { 
            if($i -lt 3)
            {                
                $taskCardNumber += [string] $taskCardNumberArray[$i] + "-"
            }
            else
            {
                $taskCardNumber += [string] $taskCardNumberArray[$i]
            }
        }
        $taskCardList2 += New-Object -TypeName PSObject -Property @{
            TaskCardNumber = $taskCardNumber
            TechName = $techName                          
            DMC = $dm.Name
        } | Select DMC, TechName,TaskCardNumber
    }
}

# Export it 
$taskCardList | Export-XLSX -Path "$exportFolder\$reportName" -Header DMC,TechName,TaskCardNumber -WorksheetName "TaskCards from IETM"
$taskCardList2| Export-XLSX -Path "$exportFolder\All CSDB Task Cards.xlsx" -Header DMC,TechName,TaskCardNumber -WorksheetName "TaskCardsincluding"
"$exportFolder\$reportName"