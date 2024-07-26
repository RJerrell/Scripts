<#
    Title: KC46 - Create a list of the Task Card from S1000D data
    Author: Roger Jerrell
    Date Created: 09/11/2017
    Purpose: Create an Excel spreadsheet listing all the Task Cards by DMC, TechName, InfoName, TaskCardNumber
    Description of Operation: Use a white list of values to drive the values
    Description of Use: 
        - Get the PMC for the Task Card Manual from the RAM drive location ( no the local c: drive location)
        - Process the PMC to get the Task List #1
        - Process the RAM drive files themselves, instead of the PMC and create a 2nd worksheet showing the historical task cards, if any.
        - Create the 2 worksheet report and store it as a true Excel workbook.
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

# Data should come from the RAM Drive !!!
$PathToData = "F:\KC46 Staging\Production\Manuals"

$exportFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs"
$reportName = "Task Card List Based on the TC PMC for the current revision.xlsx"


[string[]] $PubList   = @("TC")

$taskCardList = @()

$pmXml = New-Object System.Xml.XmlDocument
$dmXml = New-Object System.Xml.XmlDocument

# Aggregates the data for the report based on the PMC dmRef entries
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
            $infoName = $dmXml.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName

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
                InfoName = $infoName
                DMC = $fileName
            } | Select DMC,TechName,InfoName,TaskCardNumber
        }
    }
}

$taskCardList2 = @()
# Aggregates the data for the report based on the actual data modules in the RAM drive folder
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
        $infoName = $dmXml.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName
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
        $sNameArray = $dm.Name.Split("_")
        $sName = [string] $sNameArray[0]
        $taskCardList2 += New-Object -TypeName PSObject -Property @{
            TaskCardNumber = $taskCardNumber
            TechName = $techName
            InfoName =  $infoName                         
            DMC = $sName
        } | Select DMC,TechName,InfoName,TaskCardNumber
    }
}

# Export it

Remove-Item -Path  "$exportFolder\$reportName" -Force

$taskCardList  | Sort-Object -Property "TaskCardNumber"  | Export-XLSX -Path "$exportFolder\$reportName" -Header DMC,TechName,InfoName,TaskCardNumber -WorksheetName "TaskCards Based on PMC Entries" -ReplaceSheet
$taskCardList2 | Sort-Object -Property "TaskCardNumber"  | Export-XLSX -Path "$exportFolder\$reportName" -Header DMC,TechName,InfoName,TaskCardNumber -WorksheetName "TaskCards Based on File System" -ReplaceSheet
"$exportFolder\$reportName"