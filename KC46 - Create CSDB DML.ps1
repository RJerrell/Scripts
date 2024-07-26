<#
    Title: KC46 - Create CSDB DML.ps1
    Author: Roger Jerrell
    Date Created: 04/12/2017
    Purpose: The script provides the correct DML for the current release
    Description of Operation: Aggregates the DMLs for the current replease into a single listing
    Description of Use: Packages into the outbound IETM package for delivery to the USAF
#>
CLS
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
Import-Module -Name "KC46S1000DRules" -Verbose -Force
Import-Module -Name "KC46DataManagement" -Verbose -Force
# *****************************************************************************************************
#region Variable and Constants
# A list of all the folders that are going to be processed.  Not in the list?  It won't be processed.
$commonRoot = "KC46"

# RESOURCE LOCATIONS
$releseSetToProcess = "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set\Release 7 DMLs\DML*.xml"
$dmlTemplateRootPath = "C:\KC46 Staging\Scripts\Templates"

$dmlTemplatePath = "$dmlTemplateRootPath\DML-1KC46-81205-P-2017-00001_001-00_SX-US.xml"
$dmlCounterPath = "C:\KC46 Staging\Scripts\Templates\DMLCounters.txt"
$manualsRootPath = "C:\KC46 Staging\Production\Manuals"
$dmEntryType = "p"
$dmEntryTypeUpper = $dmEntryType.ToUpper()

# Template for the element inside the dml.xml representing each file in the dispatch
$dmlListItemTemp = "<dispatchFileName>#VALUE#</dispatchFileName>"
#endregion

#region - Create all the entries for the dml.XML

# What stuff is going to be called
$YR = $sd.Year
$MO = $sd.Month
$DAY = $sd.Day

$dmlOutputFileName = "DML-1KC46-81205-$dmEntryTypeUpper-$YR-#SEQUENCE#_001-00_SX-US.xml"
$dmlOutputPath = "C:\KC46 Staging\Production\Manuals\KC46\S1000D\DML\$dmlOutputFileName"
$dmlFinal = New-Object System.Xml.XmlDocument
$dmlTemp = New-Object System.Xml.XmlDocument
$dmlTemp.Load($dmlTemplatePath)
$dmlFinal.LoadXml($dmlTemp.OuterXml)
$dmlTemp = $null
$dmlFinal_ContentNode = $dmlFinal.SelectSingleNode("/dml/dmlContent")
$dmlEntryList = ""
$dmls = gci -Path $releseSetToProcess  | Sort-Object -Property Length

$bookLetters = Get-Bookletters

$comment = ""
foreach ($dml in $dmls)
{
    $bookParts = $dml.Name.Split("-")
    $book = $bookParts[2].Substring($bookParts[2].Length - 1, 1)
    $pub = ($bookLetters.GetEnumerator() | ?{$_.Value -eq $book}).Name

    switch ($pub){
    "KC46" {
        $comment = "Common Lists"
        break
    }
    "ACS" {
        $comment = "Aircraft Cross Servicing Guide ($pub)"
        break    
        }
    "ARD" {
        $comment = "Aircraft Recovery Document ($pub)"
        break    
    }
    "AMM" {
        $comment = "Aircraft Maintenance Manual ($pub)"
        break    
    }
    "FIM" {
        $comment = "Fault Isolation Manual ($pub)"
        break    
    }
    "IPB" {
        $comment = "Illustrated Parts Book ($pub)"
        break    
    }
    "LOAPS" {
        $comment = "List of Applicable Publications ($pub)"
        break    
    }
    "NDT" {
        $comment = "Nondestructive Test Manual ($pub)"
        break    
    }
    "SIMR" {
        $comment = "Standard Inspection Maintenance Repairs Manual ($pub)"
        break    
    }
    "SPCC" {
        $comment = "System Peculiar Corrosion Control ($pub)"
        break    
    }
    "SSM" {
        $comment = "System Schematic Manual ($pub)"
        break    
    }
    "SWPM" {
        $comment = "Standard Wiring Practices Manual ($pub)"
        break    
    }
    "TC" {
        $comment = "Task Cards ($pub)"
        break    
    }
    "WUC" {
        $comment = "Work Unit Code Manual ($pub)"
        break    
    }
    "WDM" {
        $comment = "Wiring Diagram Manual ($pub)"
        break
    }
}
    
    $dmlXML = New-Object System.Xml.XmlDocument
    $dmlXML.Load($dml.FullName)
    $gutsOfTheDML = $dmlXML.SelectNodes("/dml/dmlContent/dmEntry")
    $dmlTypeNodePath = $dmlXML.SelectSingleNode("/dml/identAndStatusSection/dmlAddress/dmlIdent/dmlCode").dmlType

    $dmlEntryList += "<!--`t ***`t ***`t ***`t$comment`t *** `t *** `t *** `t -->" + "`r`n"

    foreach($n in $gutsOfTheDML)
    {    
        $dmlEntryList += $n.OuterXml + "`r`n"
    }   
}

$dmlFinal_ContentNode.InnerXml = $dmlEntryList

#endregion
#region - Establish the next sequence number to be used

# Create a 5-digit left-zero-padded sequence number

attrib -R $dmlCounterPath

$seqNum = ""
$sequenceNumber = ""
$sr = New-Object System.IO.StreamReader($dmlCounterPath)
$seqNumStr = $sr.ReadLine()
$sr.Close()
$sr.Dispose()

$seqNum = [int] $seqNumStr
$sequenceNumber = $seqNum.ToString("00000")
$sequenceNumber
$newNumber = [int] $seqNumStr + 1

$sw = [System.IO.StreamWriter] $dmlCounterPath

$sw.WriteLine($newNumber)

$sw.Close()

$sw.Dispose()

#$sequenceNumber = "00001"

$dmlOutputNewPath = $dmlOutputPath.Replace("#SEQUENCE#", $sequenceNumber)
$dmlOutputNewPath

attrib +R $dmlCounterPath

#endregion
#region - Set values inside the header of the outgoing dml to match your new sequence 

$MM = [INT] $MO
$DD = [INT] $DAY
# PATH TO DMLType attribute value
$dmlFinal.dml.identAndStatusSection.dmlAddress.dmlIdent.dmlCode.dmlType = $dmEntryType
$dmlFinal.dml.identAndStatusSection.dmlAddress.dmlIdent.dmlCode.seqNumber = $sequenceNumber
$dmlFinal.dml.identAndStatusSection.dmlAddress.dmlIdent.dmlCode.yearOfDataIssue = $YR.ToString()
$dmlFinal.dml.identAndStatusSection.dmlAddress.dmlAddressItems.issueDate.year = $YR.ToString()
$dmlFinal.dml.identAndStatusSection.dmlAddress.dmlAddressItems.issueDate.month = $MM.ToString("00")
$dmlFinal.dml.identAndStatusSection.dmlAddress.dmlAddressItems.issueDate.day = $DD.ToString("00")

#endregion
Save-PrettyXML -FName $dmlOutputNewPath -xmlDoc $dmlFinal

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"