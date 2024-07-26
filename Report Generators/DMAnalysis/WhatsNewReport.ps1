<#
Title: WhatsNewReport.ps1
Author: Roger Jerrell
Date Created: 04/28/2017
Purpose: Processes each highlights data module into a metrics reports
Description of Operation: 
 - Takes in each highlights data module, 1 per manual, and creates a csv report
Description of Use:
 - Rename the report to the desired name ($reportName)
 - Confirm the path to the inputs ($pathToHighlightsDM)
 - Run

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

#Get-Module -ListAvailable
Import-Module -Name "KC46Common" -Verbose -Force
# *****************************************************************************************************
$dm_allnamesList = ""
$objs = @()
$pathToHighlightsDM = "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set\Highlights"
$hdmFiles = gci -Path $pathToHighlightsDM -Filter *K*.xml |Sort-Object
# Report variables
$outputPath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$reportName = "Release 10 - Highlights Data Modules Breakdown.csv"

foreach ($hdmFile in $hdmFiles)
{   
    $xml = New-Object System.Xml.XmlDocument
    $xml.Load($hdmFile.Fullname)
    # /dmodule/content/description/leveledpara/table/tgroup/tbody/row
    $tableRows = $xml.SelectNodes("/dmodule/content/description//table/tgroup/tbody/row")
    foreach ($tableRow in $tableRows)
    {
        $state = ""
        # Get the No technical change rows from the highlights dm
        $paraCount = $tableRow.entry[9].ChildNodes.Count
        #$tableRow.entry[9].para.Trim()
        if($paraCount-eq 1)
        {
            if($tableRow.entry[9].para.Trim() -eq "Reissued the data module without technical changes.")
            {
                $state = "NTC"
                $dmNode = $tableRow.entry[1].para.dmRef
                $dmc = Get-FilenameFromDMRef -dmRef $dmNode -filePref "DMC"
                $obj = [pscustomobject][ordered]@{Manual=$tableRow.entry.para[0];DMC=$dmc;TechName=$tableRow.entry.para[1].InnerText;State=$state}
                $objs += $obj
                $obj=$null
            }
        }
        if($tableRow.entry[9].para[$paraCount-1] -contains "Reissued the data module without technical changes.")
        {
            $state = "NTC"
            $dmNode = $tableRow.entry[1].para.dmRef
            $dmc = Get-FilenameFromDMRef -dmRef $dmNode -filePref "DMC"
            $obj = [pscustomobject][ordered]@{Manual=$tableRow.entry.para[0];DMC=$dmc;TechName=$tableRow.entry.para[1].InnerText;State=$state}
            $objs += $obj
            $obj=$null

            #break
        }
        # Get the new ones
       
        if($tableRow.entry[5].para -eq "001")
        {
            $state = "New"
            $dmNode = $tableRow.entry[1].para.dmRef
            $dmc = Get-FilenameFromDMRef -dmRef $dmNode -filePref "DMC"
            $obj = [pscustomobject][ordered]@{Manual=$tableRow.entry.para[0];DMC=$dmc;TechName=$tableRow.entry.para[1].InnerText;State=$state}
            $objs += $obj
            $obj=$null
           # break
        }

        elseif($tableRow.entry[5].para-gt "001")
        {
            $state = "Changed"
            $dmNode = $tableRow.entry[1].para.dmRef
            $dmc = Get-FilenameFromDMRef -dmRef $dmNode -filePref "DMC"
            $obj = [pscustomobject][ordered]@{Manual=$tableRow.entry.para[0];DMC=$dmc;TechName=$tableRow.entry.para[1].InnerText;State=$state}
            $objs += $obj
            $obj=$null
        }
        else
        {
            $dmNode = $tableRow.entry[1].para.dmRef
            $dmc = Get-FilenameFromDMRef -dmRef $dmNode -filePref "DMC"
            $dmc
            "none of the above"
        }
    }
}
# Define the sort properties for the report
$prop1 = @{Expression='Manual'; Ascending=$true }
$prop2 = @{Expression='State'; Ascending=$true }
$prop3 = @{Expression='DMC'; Ascending=$true }


# Store the report
$objs.GetEnumerator() | Sort-Object -Property $prop1, $prop2,$prop3 | Export-Csv "$outputPath\$reportName" -NoTypeInformation 
"$outputPath\$reportName"
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"