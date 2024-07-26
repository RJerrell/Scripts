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
$ErrorActionPreference = "Stop"
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
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force 
# *****************************************************************************************************
$rel5Path = "R:\2017-01-20-14-18-01 - Non CDRL January 2017 - Release 5\CSDB\Manuals\AMM\S1000D\SDLLIVE"
$rel6Path = "R:\2017-06-06-07-18-23 - Non CDRL June 2017 - Release 6\CSDB\DVD\AMM\S1000D\SDLLIVE"
$rel7Path = "R:\2017-09-18-14-39-31 - Non CDRL Sept 2017 - Release 7\CSDB\DVD\AMM\S1000D\SDLLIVE"

$pathToTheGOOOOPile = "\\nw.nos.boeing.com\newgen-tanker\Spt_TrngSys\SptSys\Tech Pubs\Validation\C&V_workbook\DMCertVerWorksheet.xlsx"
$pathToUSAFTemplate = "C:\KC46 Staging\Scripts\Report Generators\Outputs\Verification\Out\Release 7-8\BOEING Verification - Releases 5 thru 8 - INTERNAL BOEING.xlsx"

if(Test-Path -Path "$env:TMP\DMCertVerWorksheet.xlsx")
{
    Remove-Item -Path "$env:TMP\DMCertVerWorksheet.xlsx" -Force
}

Copy-Item -Path $pathToTheGOOOOPile -Destination $env:TMP -Force

$pathToTheGOOOOPile = "$env:TMP\DMCertVerWorksheet.xlsx"
$GOOOWorkSheetName = "StatusWorksheet"
$baseelineRows = Import-XLSX -Path $pathToUSAFTemplate  -Sheet "Verifications" # -Verbose
$ttandooList = @()
foreach($baseelineRow in $baseelineRows)
{
    if($baseelineRow.BOE_CertificationType -eq "ttandoo")
    {
        $ttandooList += $baseelineRow.PBTask
    }
}
$ttandooList2 = $ttandooList | Sort-Object

$excel = new-object -comobject Excel.Application
$workbook = $excel.Workbooks.Open($pathToTheGOOOOPile)
$ws = $Workbook.worksheets | where-object {$_.Name -eq $GOOOWorkSheetName}
$CTR=1
foreach($ttandooItem in $ttandooList2)
{    
    $ttandooItem
    
    for ($i = 1; $i -lt 2500; $i++)
    { 
        if($ttandooItem -eq $ws.Columns[13].Cells[$i].Value2)
        {
            $ws.Columns[1].Cells[$i].Value2 = 1
            $CTR
            $CTR++
            break
        }
    }    
}

$ws.SaveAs("C:\KC46 Staging\Scripts\Report Generators\Outputs\Verification\Out\1.xlsx")
$workbook.Close()
$excel.Quit()

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Minutes to complete:`t" + $x.TotalMinutes
"Report now available at this location:`r`n$exportFolder\$reportName"