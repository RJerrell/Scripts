
<#

Title: 
Author: Roger Jerrell
Date Created: Get-Date 
Purpose: 
Description of Operation: 
Description of Use:

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

$env:PSModulePath = "C:\Program Files\WindowsPowerShell\Modules;C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
                                                                                  
# *****************************************************************************************************

$ChangedItems = @()
$ChangedItems += "WS Row|WS PBTask|WS DMC|CAS DMC|WS Techname|CAS Techname|WS Infoname|CAS Infoname"
$basepath = "C:\KC46 Staging\Scripts\C&V"
$CASMapping = Import-Csv  -Path "$basepath\AMM-KC_s1000D_to_ATA_mapping.csv"  -Verbose
cls

#Install-Module -Name ImportExcel -Verbose -Force
# Create an Excel object AND get the C&V spreadsheet from ROD and BOB into memory

$Excel = New-Excel -Path "$basepath\DMCertVerWorksheet.xlsx"
$WorkSheet = $Excel | Get-WorkSheet -Name "StatusWorksheet"
for ($i = 2; $i -lt $WorkSheet.Cells.Rows ; $i++)
{   
    if($WorkSheet.Cells[$i,12].Value.Length -gt 0)
    {
        #cls
        $WS_PBTASK = $WorkSheet.Cells[$i,12].Text
        if($WS_PBTASK.Length -gt 0)
        {
            for ($z = 0; $z -lt $CASMapping.Count; $z++)
            { 
                $casAMTOWS_PBTASKue = "Task " + $CASMapping[$z].'ATA-AMTOSS'
                
                # Perform the lookup and compare the CAS map value to the spreadsheet value for the PB/TASK
                if($casAMTOWS_PBTASKue.ToUpper() -eq $WS_PBTASK.ToUpper())
                {
                    <# Before we change the DMC value, let's log it so we can actually go back and see what changed. #>
                    if($WorkSheet.Cells[$i,13].Text.ToUpper() -ne $CASMapping[$z].'S1000D-DMC'.ToUpper())
                    {                        
                        $ChangedItems += $i.ToString() +  '|' + $WS_PBTASK + '|' + $WorkSheet.Cells[$i,13].Text + '|' + $CASMapping[$z].'S1000D-DMC' + '|' + $WorkSheet.Cells[$i,17].Text + '|' + $CASMapping[$z].'TECH-NAME' + '|' + $WorkSheet.Cells[$i,18].Text + '|' + $CASMapping[$z].'INFO-NAME'
                        $null = $WorkSheet.SetValue($i,13,$CASMapping[$z].'S1000D-DMC')
                        # After the change :  $WorkSheet.Cells[$i,13].Text
                        break
                    }                    
                }                
            }                       
        }
    }
}
$Excel | sAVE-Excel -Save -Path "$basepath\ABC.XLSX"
$Excel | Export-Excel -Path "$basepath\ABC.XLSX" -PassThru

$ChangedItems | Export-Excel "$basepath\DMCertVerWorksheet - Revision Report.xlsx" -WorkSheetname "Revision Report"  -Show -FreezeTopRow $true

# *****************************************************************************************************

$ed = Get-Date

$x = $ed.Subtract($sd)

"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"