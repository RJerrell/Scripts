<# 
    Get all the task cards (TC)

    Get the dmCode off of each Task Card
    
    Turn each dmCode element into a "link" for Juanita, meaning, we are going to convert
    each dmCode to a dmRef
    
    Then, we'll turn that list of Task Card dmRefs into an Excel for her use.
#>
CLS
$ErrorActionPreference = "Stop"
$error.Clear()
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
# *****************************************************************************************************
$parserDMC  = New-object -TypeName S1000D.DataModule_401
$parserDML  = New-Object -TypeName S1000D.DataModuleList
$parserCOM = New-Object -TypeName S1000D.CommonFunctions
$parserPM    = New-object -TypeName S1000D.PublicationModule_401
$pathtoTCReports = "C:\KC46 Staging\Scripts\Report Generators\Outputs\TaskCards"
$pathToTaskCards = "C:\KC46 Staging\Production\Manuals\TC\S1000D\SDLLIVE"
$pmc = gci -Path $pathToTaskCards -Filter "pmc*.xml" | Sort-Object -Descending | Select-Object -First 1
$pmc[0].FullName
$reportName = "KC46 - Task Card Inventory - $startTime.xlsx"

$dmc = gci -Path $pathToTaskCards -Filter "dmc*.xml" | Sort-Object
$dmc[0].FullName
$parserPM.ParsePM($pmc[0].FullName)
$rows = @()
foreach ($dmRef in $parserPM.DmRefs)
{
    $techName= ""
    $dmc = ($dmRef.dmRefIdent.'#comment' ).Replace(": " , "-").Trim()
    $tcDataModule = gci -Path $pathToTaskCards -Filter "$dmc`*.xml" | Sort-Object -Descending | Select-Object -First 1

    $parserDMC.ParseDM($tcDataModule[0].FullName)

    $techName = $parserDMC.TechName

    $dr = $dmRef.OuterXml
    $rows += New-Object -TypeName PSObject -Property @{
                dmc = $dmc;
                techname = $techName;
                dmref = $dr;
    } | Select dmc, techname, dmref  
}

 $rows | Export-XLSX -Path "$pathtoTCReports\$reportName" -Header dmc, techname, dmref  -WorksheetName "Task Card Refs"
 "$pathtoTCReports\$reportName"

# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"					
					