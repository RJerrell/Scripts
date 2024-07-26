<#
    Title: KC46 - Is this DMC referenced by any other DMC
    Author: Roger Jerrell
    Date Created: 1010/2017 
    Purpose: Create a report that determines if a given data module is refernced by any other data module(s)
    Description of Operation: Takes input of the DMC in question and performs search across the current CSDB
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
$FNameToFind = "F:\KC46 Staging\Production\Manuals\AMM\S1000D\SDLLIVE\DMC-1KC46-A-28-05-0000-08A0A-310A-A_001-00_SX-US.xml"


                                                                        
$dmxml = Get-DM -pathToDM $FNameToFind

$nameToMatch = Create-DMCFileNameFromDMCode -dmCode $dmxml.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode

$fileList = @()
$librayToSearchPath = "F:\KC46 Staging\Production\Manuals\TC\S1000D\SDLLIVE"
$librayToSearch     =  gci -Path $librayToSearchPath -Filter "DMC*.XML"
$dmXml = New-Object System.Xml.XmlDocument

foreach ($dmFile in $librayToSearch)
{
    $dmXml.Load($dmFile.FullName)
    $refs = $dmXml.dmodule.content.refs
    if($refs.dmRef.Count -gt 0)
    {
        $dmRefs = $refs.dmRef
        foreach ($dmRef in $dmRefs)
        {
            $dmFileName = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
            if($dmFileName -eq $nameToMatch)
            {
                $fileList += $dmFileName
                "match"
            }
            else
            {
                "No matches"
            }
        }
    }
}
$fileList
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"