<#
This script produces 2 reports-- 1 CSV and 1 XML with the former being a boring CSV listing for spreadsheet 
users and the other a shortlist used as an input to other programs / scripts.

CSV Report location - "C:\KC46 Staging\scripts\Report Generators\Outputs\PMC DMRef Master LongList.CSV"
XML Report Location - "C:\KC46 Staging\scripts\Report Generators\Outputs\PMC DMRef Master ShortList.xml"

            Uncomment line 20 to run this report for all manuals listed on that line!
#>
cls
$error.clear()
$dt = (Get-Date -Format yyyy-MM-dd-HH-mm-ss)

$environment = "Production"
$outputReportPath1 = "C:\KC46 Staging\Scripts\Report Generators\Outputs\PBTask to DMC Map.CSV"
$outputReportPath2 = "C:\KC46 Staging\Scripts\Report Generators\Outputs\PBTask to DMC Map.xml"
$KC46DataRoot = "C:\KC46 Staging"
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"
#Book|Type|Book|PB Status|Revision|Rev Type|Rev Date|Highlight|Navigation Tree Title|Task Title|Pageblock|C&V|PB/Task|DMC|IC Review|DM Type|Review	Type|TechName|InfoName|Ver List|Cert Complete|Pblock|Support Equipment/Notes|C|V|USAF Rep|ETOPS?|Daily Plng|Proc Avail|Cert Interim Comp|Cert Comp|Ver Interim Comp|Ver Comp|AC Type|CAS Incorp|USAF Printed (P) - Delivered(Date)|Tasks Too Difficult To Do|Chapter|Section|Subject|Info1|Comb1|Info2|Comb2|PB	Info3|Prep Time|Est Duration (Adtnl)|Post Time|Ver Type|Rationale|AMC Recommendation|Boeing Review of AF Recommendation|Done|Boeing Recommendation|Original Ver Type|Original Ver List|Original Rationale|CertTeamRecord
if(Test-path -Path $outputReportPath1)
{
    Remove-Item -Path $outputReportPath1 -Force
}

if(Test-path -Path $outputReportPath2)
{
    Remove-Item -Path $outputReportPath2 -Force
}

[string[]] $bookList   = @("AMM")

$masterListLong = @()
$masterListShort = @()
$masterListLong += "BookType|Book|DMC|IssueNumber|PB|Task|FigID|FigTitle|GraphicID|ICN|GraphicTitle|MIC|CH|SE|SU|DC|DCV|IC|ICV|ILC|ICR"
$masterListShort += "DMC|Task"
$book = "AMM"
    $pm = [xml] (Get-Content -Path "C:\KC46 Staging\Production\Manuals\AMM\S1000D\SDLLIVE\PMC-1KC46-81205-A0000-00.xml")
    $dmcCollection = $pm.SelectNodes("//content//dmRef")
     
    $lastPB = ""
    $pborTask= "blank"
    for ($i = 0; $i -lt $dmcCollection.Count; $i++)
        { 
        [STRING] $modelIdentCode = $dmcCollection[$i].dmRefIdent.dmCode.modelIdentCode
        [STRING] $systemDiffCode = $dmcCollection[$i].dmRefIdent.dmCode.systemDiffCode
        [STRING] $systemCode = $dmcCollection[$i].dmRefIdent.dmCode.systemCode
        [STRING] $subSystemCode = $dmcCollection[$i].dmRefIdent.dmCode.subSystemCode
        [STRING] $subSubSystemCode = $dmcCollection[$i].dmRefIdent.dmCode.subSubSystemCode
        [STRING] $assyCode = $dmcCollection[$i].dmRefIdent.dmCode.assyCode
        [STRING] $disassyCode = $dmcCollection[$i].dmRefIdent.dmCode.disassyCode
        [STRING] $disassyCodeVariant = $dmcCollection[$i].dmRefIdent.dmCode.disassyCodeVariant
        [STRING] $infoCode = $dmcCollection[$i].dmRefIdent.dmCode.infoCode
        [STRING] $infoCodeVariant = $dmcCollection[$i].dmRefIdent.dmCode.infoCodeVariant
        [STRING] $itemLocationCode = $dmcCollection[$i].dmRefIdent.dmCode.itemLocationCode
        [STRING] $ch = $systemCode
        [STRING] $se = $subSystemCode+$subSubSystemCode
        [STRING] $su = $assyCode
            
        $title = $dmcCollection[$i].title
            
        $pborTask = $dmcCollection[$i].href
            
        if($pborTask -like "PAGEBLOCK*")
        {
            $lastPB = $pborTask
        }

        $fileName =  "$source_BaseLocation\$book\S1000D\SDLLIVE\DMC-$modelIdentCode-$systemDiffCode-$systemCode-$subSystemCode$subSubSystemCode-$assyCode-$disassyCode$disassyCodeVariant-$infoCode$infoCodeVariant-$itemLocationCode.xml"
        $fi = gci -Path $fileName
        $masterListShort += $fi.Name +  "|" + $pborTask
    }



$masterListShort | Export-Clixml $outputReportPath2 -Force
"Complete!"