cls
#region Header
    <#
        Title: 
        Author: Roger Jerrell
        Date Created: 04/06/2017 
        Purpose: Creates a deliverable to the USAF representing the entire contents of the CSDB
        Description of Operation: Rolls over the Manuals\$pub file structure and inventories both the SDLLIVE data and the generic files as well.
        Description of Use: Use this script to create the deliverable.  Output is placed in the "C:\KC46 Staging\Scripts\Report Generators\Outputs" folder
    #>
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
#endregion
# *****************************************************************************************************
#region Variable and Constants

# A list of all the folders that are going to be processed.  Not in the list?  It won't be processed.
[string[]] $PubList = @("KC46","ACS","ARD","AMM","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM","TC","WUC", "WDM")

# Where stuff is located
$ddnTemplateRootPath = "F:\KC46 Staging\Scripts\Templates"
$ddnTemplatePath = "$ddnTemplateRootPath\DDN-1KC46-81205-80049-2017-00001.xml"
$ddnCounterPath = "$ddnTemplateRootPath\DDNCounters.txt"
$manualsRootPath = "F:\KC46 Staging\Production\Manuals"

# Template for the element inside the DDN.xml representing each file in the dispatch
$ddnListItemTemp = "<dispatchFileName>#VALUE#</dispatchFileName>"
#endregion
if(!( Test-Path -Path $ddnTemplateRootPath ))
{
    md $ddnTemplateRootPath
}

#region - Create all the entries for the DDN.XML

# What stuff is going to be called
$YR = $sd.Year
$MO = $sd.Month
$DAY = $sd.Day
$ddnOutputFileName = "DDN-1KC46-81205-80049-$YR-#SEQUENCE#.xml"
$ddnOutputPath = "F:\KC46 Staging\Production\Manuals\KC46\S1000D\DDN\$ddnOutputFileName"

$ddn = New-Object System.Xml.XmlDocument
$ddnTemp = New-Object System.Xml.XmlDocument
$ddnTemp.Load($ddnTemplatePath)
$ddn.LoadXml($ddnTemp.OuterXml)
$ddnTemp = $null

# Delivery List element that all of our entries will append to as we cycle through the CSDB
$dnContentNode = $ddn.SelectSingleNode("/ddn/ddnContent")
foreach ($pub in $PubList)
{
    $filesDMGeneric = gci -Path "$manualsRootPath\$pub\S1000D\*.xml" -Exclude DML*.xml| Sort-Object
    $filesDMSDL = gci -Path "$manualsRootPath\$pub\S1000D\SDLLIVE\*.xml" -Exclude DML*.xml| Sort-Object
    $filesILL = gci -Path "F:\KC46 Staging\Production\Manuals\$pub\ILLUSTRATIONS\ILLUSTRATIONS\*.cgm" -Exclude *.txt| Sort-Object

    $newDL = $ddn.CreateElement("deliveryList")
    $comment = ""
    switch ($pub){
    "KC46" {
        $comment = "Common Lists"
    }
    "ACS" {
        $comment = "Aircraft Cross Servicing Guide ($pub)"
    }
    "ARD" {
        $comment = "Aircraft Recovery Document ($pub)"
    }
    "AMM" {
        $comment = "Aircraft Maintenance Manual ($pub)"
    }
    "FIM" {
        $comment = "Fault Isolation Manual ($pub)"
    }
    "IPB" {
        $comment = "Illustrated Parts Book ($pub)"
    }
    "LOAPS" {
        $comment = "List of Applicable Publications ($pub)"
    }
    "NDT" {
        $comment = "Nondestructive Test Manual ($pub)"
    }
    "SIMR" {
        $comment = "Standard Inspection Maintenance Repairs Manual ($pub)"
    }
    "SPCC" {
        $comment = "System Peculiar Corrosion Control ($pub)"
    }
    "SSM" {
        $comment = "System Schematic Manual ($pub)"
    }
    "SWPM" {
        $comment = "Standard Wiring Practices Manual ($pub)"
    }
    "TC" {
        $comment = "Task Cards ($pub)"
    }
    "WUC" {
        $comment = "Work Unit Code Manual ($pub)"
    }
    "WDM" {
        $comment = "Wiring Diagram Manual ($pub)"
    }
}
# @("KC46","ACS","ARD","AMM","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM","TC","WUC", "WDM")
    $newComment = $ddn.CreateComment($comment)

    $null = $newDL.AppendChild($newComment)

    foreach ($file in $filesDMGeneric)
    {
        $newEle = $ddn.CreateElement("deliveryListItem")
        $newEle.InnerXml = $ddnListItemTemp.Replace("#VALUE#", $file.FullName.Replace("F:\KC46 Staging\Production\Manuals", ""))
        #$newEle.OuterXml
        $null = $newDL.AppendChild($newEle)
    }

    foreach ($file in $filesDMSDL)
    {
        $newEle = $ddn.CreateElement("deliveryListItem")
        $newEle.InnerXml = $ddnListItemTemp.Replace("#VALUE#", $file.FullName.Replace("F:\KC46 Staging\Production\Manuals", ""))
        #$newEle.OuterXml
        $null = $newDL.AppendChild($newEle)
    }

    foreach ($file in $filesILL)
    {
        $newEle = $ddn.CreateElement("deliveryListItem")
        $newEle.InnerXml = $ddnListItemTemp.Replace("#VALUE#", $file.FullName.Replace("F:\KC46 Staging\Production\Manuals", ""))
        #$newEle.OuterXml
        $null = $newDL.AppendChild($newEle)
    }

    $null = $dnContentNode.AppendChild($newDL)
    $pub
}
#endregion

#region - Establish the next sequence number to be used

# Create a 5-digit left-zero-padded sequence number
attrib -R $ddnCounterPath
$seqNum = ""
$sequenceNumber = ""
$sr = New-Object System.IO.StreamReader($ddnCounterPath)
$seqNumStr = $sr.ReadLine()
$sr.Close()
$sr.Dispose()

$seqNum = [int] $seqNumStr
$sequenceNumber = $seqNum.ToString("00000")
$sequenceNumber
$newNumber = [int] $seqNumStr + 1

$sw = [System.IO.StreamWriter] $ddnCounterPath
$sw.WriteLine($newNumber)
$sw.Close()
$sw.Dispose()

$ddnOutputNewPath = $ddnOutputPath.Replace("#SEQUENCE#", $sequenceNumber)
$ddnOutputNewPath
attrib +R $ddnCounterPath
#endregion

#region - Set values inside the header of the outgoing DDN to match your new sequence 

# /ddn/identAndStatusSection/ddnAddress/ddnAddressItems/issueDate
$ddn.ddn.identAndStatusSection.ddnAddress.ddnIdent.ddnCode.yearOfDataIssue = $YR
$ddn.ddn.identAndStatusSection.ddnAddress.ddnIdent.ddnCode.seqNumber = $sequenceNumber

$MM = [INT] $MO
$DD = [INT] $DAY
$ddn.ddn.identAndStatusSection.ddnAddress.ddnAddressItems.issueDate.year = $YR
$ddn.ddn.identAndStatusSection.ddnAddress.ddnAddressItems.issueDate.month = $MM.ToString("00")
$ddn.ddn.identAndStatusSection.ddnAddress.ddnAddressItems.issueDate.day = $DD.ToString("00")

#endregion
Save-PrettyXML -FName $ddnOutputNewPath -xmlDoc $ddn
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"