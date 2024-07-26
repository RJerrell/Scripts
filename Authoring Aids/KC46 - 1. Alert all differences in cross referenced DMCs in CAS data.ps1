cls
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'

Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force 
# *****************************************************************************************************
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose

$parserDM = new-object -TypeName S1000D.DataModule_401
$parserPM = new-object -TypeName S1000D.PublicationModule_401
$parserCommon = New-Object -TypeName S1000D.CommonFunctions

$rptTimeStamp = $sd.ToShortDateString().Replace("/","-")
$exportFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs\BoilerPlateReports\AuthoringDiscrepancies"
$reportName = "BDS - CAS  Report.csv"

$basePathToManuals = "F:\KC46 Staging\Production\Manuals"
$Discrepancies = @()
$targetDMRefs = @() 
#$ammDMCList = New-Object System.Collections.Generic.List[String]

[string[]] $srcDMCPubList   = @("ABDR","ACS","ASIP","LOAPS","SIMR","SPCC","WUC")
[string[]] $targetPubList   = @("KC46","ABDR","ACS","AMM","ARD","ASIP","FIM","IPB","LOAPS","NDT","SIMR","SPCC","TC","WUC","SSM","SWPM","WDM") | Sort-Object
$a = $targetPubList | where {$srcDMCPubList -notcontains $_}
$a

foreach ($pub in $a)
{
$pub
    # Get the PMC for the target book
    $pathToData = "$basePathToManuals\$pub\S1000D\SDLLIVE"
    $module = gci -Recurse -path $pathToData -Filter PMC*.XML  | Sort-Object -Descending | Select-Object -First 1 # gET BOTH THE PMC AND DMC FILES
    $parserPM.ParsePM($module[0].FullName)
    $pmName = $module[0].Name
    foreach ($dmRef in $parserPM.DmRefs)
    {
        # What book is the dmRef pointing to from within the PMC itself?
        $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"        
        $targetBook = $parserCommon.GetAssociatedBook($dmc) 
        $targetDMRefs += $dmc
    }
}

foreach ($pub in $srcDMCPubList)
{
    # Get the PMC for the target book
    $pathtoSrcData = "$basePathToManuals\$pub\S1000D\SDLLIVE"

    $srcPubModule = gci -Recurse -path $pathtoSrcData -Filter PMC*.XML  | Sort-Object -Descending | Select-Object -First 1 # gET BOTH THE PMC AND DMC FILES
    $parserPM.ParsePM($srcPubModule[0].FullName)
    $pmName = $srcPubModule[0].Name
    foreach ($dmRef in $parserPM.DmRefs)
    {
        # What book is the dmRef pointing to from within the PMC itself ?
        $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
        $dmBook = $parserCommon.GetAssociatedBook($dmc) 
        if($pub -ne $dmBook)
        {
            if(!$targetDMRefs.Contains($dmc))
            {
                $Discrepancies += New-Object -TypeName PSObject -Property @{
                Book = $pub;
                DMC = $pmName;
                TargetBook=$targetBook;
                TargetDMC=$dmc;
            } | Select Book, DMC,TargetBook,TargetDMC}
            else
            {
                $targetbook = $parserCommon.GetAssociatedBook($dmc) 
                $pathToDm = "$basePathToManuals\$targetbook\S1000D\SDLLIVE\$dmc`*.xml"
                $targetDM = gci -path $pathToDm
                $parserDM.ParseDM($targetDM[0].FullName)
                foreach ($dmr in $parserDM.DmRefs)
                {
                    $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
                    $book = $parserCommon.GetAssociatedBook($dmc) 
                    $pathToDm = "$basePathToManuals\$book\S1000D\SDLLIVE\$dmc`*.xml"
                    $targetDM = gci -path $pathToDm
                    $parserDM.ParseDM($targetDM[0].FullName)
                    foreach ($dmr in $parserDM.DmRefs)
                    {
                        $dmc2 = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
                        $book2 = $parserCommon.GetAssociatedBook($dmc) 
                        if(!$targetDMRefs.Contains($dmc))
                        {
                                $Discrepancies += New-Object -TypeName PSObject -Property @{
                                Book = $book;
                                DMC = $dmc;
                                TargetBook=$book2;
                                TargetDMC=$dmc2;
                         } | Select Book, DMC,TargetBook,TargetDMC}
                    }
                }
            } 
        }     
    }
}

$Discrepancies | fl

"Executive summary Report :`t" + $sd
if(!(Test-Path -Path "$exportFolder"))
{ md "$exportFolder"}
#Remove-Item -Path "$exportFolder\$reportName" -Force -ErrorAction SilentlyContinue
If($Discrepancies.Count -gt 0)
{
    $Discrepancies    | Export-XLSX -Path "$exportFolder\$reportName" -Header  Book, DMC,TargetBook,TargetDMC -WorksheetName "Discrepancies"
    "Report now available at this location:`r`n$exportFolder\$reportName"
}
else
{
    "No discrepencies"
}
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"$PSCommandPath `r`n" +  "Process completed"
