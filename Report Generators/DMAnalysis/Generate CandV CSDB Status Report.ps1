<#
This script produces 2 reports-- 1 CSV and 1 XML with the former being a being CSV listing for spreadsheet 
users and the other a shortlist used as an input to other programs / scripts.

CSV Report location - "C:\KC46 Staging\scripts\Report Generators\Outputs\PMC DMRef Master LongList2.CSV"
XML Report Location - "C:\KC46 Staging\scripts\Report Generators\Outputs\PMC DMRef Master ShortList2.xml"

Uncomment line $bookList  to run this report for all manuals listed on that line!
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

$environment = "Production"
$outputReportPath = "$env:temp\CandVstatus-CSDB.xml"
$KC46DataRoot = "C:\KC46 Staging"
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"

if(Test-path -Path $outputReportPath)
{
    Remove-Item -Path $outputReportPath -Force
}

[string[]] $bookList   = @("AMM")
#[string[]] $bookList   = @("AMM","ARD","BCLM","FIM", "IPB", "NDT","SRM","SSM","WDM")

$masterList = @()
$masterList += "PBTask|DMC|IssueNumber_Current|IssueNumber_Verified|VerficationMethod"

foreach( $book in $bookList)
{  
    $pmcColl = gci -Path "$source_BaseLocation\$book\S1000D\SDLLIVE\PMC*.xml"    
    foreach ($pmc in $pmcColl)
    {
        $pm = New-Object System.Xml.XmlDocument
        $pm.Load( $pmc.FullName )

        $dmcCollection = $pm.SelectNodes("//content//dmRef")
     
        $lastPB = ""
        $pborTask= "blank"
        for ($i = 0; $i -lt $dmcCollection.Count; $i++)
            {
            $dmc = Get-FilenameFromDMRef -dmRef $dmcCollection[$i] -filePref "DMC"
            $dmcFullName = $dmc + ".xml"
            $title = $dmcCollection[$i].title
            $pborTask = $dmcCollection[$i].href 
            
            if($pborTask -like "PAGEBLOCK*")
            {
                $lastPB = $pborTask
            }

            $fileName =  "$source_BaseLocation\$book\S1000D\SDLLIVE\$dmcFullName"
            $fi = gci -Path $fileName
            $dm = New-Object System.Xml.XmlDocument
            $dm.Load($fileName)
            $issueNumber = $dm.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.issueNumber
            # "PBTask|DMC|IssueNumber_Current|IssueNumber_Verified|VerficationMethod"
            $masterList += $pborTask +  "|" + $dmc + "|" + $issueNumber
        }
    }
}

$masterList | Export-Clixml $outputReportPath -Force
$error
$R = gci -Path "$ENV:TEMP\candv*" | Sort-Object
$R
"Complete!"