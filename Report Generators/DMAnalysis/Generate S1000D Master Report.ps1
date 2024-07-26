<#
This script produces 2 reports-- 1 CSV and 1 XML with the former being a boring CSV listing for spreadsheet 
users and the other a shortlist used as an input to other programs / scripts.

CSV Report location - "C:\KC46 Staging\scripts\Report Generators\Outputs\PMC DMRef Master LongList.CSV"
XML Report Location - "C:\KC46 Staging\scripts\Report Generators\Outputs\PMC DMRef Master ShortList.xml"

            Uncomment line 20 to run this report for all manuals listed on that line!
#>
cls
$error.clear()
$err
$environment = "Production"
$outputReportPath1 = "C:\KC46 Staging\scripts\Report Generators\Outputs\PMC DMRef Master LongList.CSV"
$outputReportPath2 = "C:\KC46 Staging\scripts\Report Generators\Outputs\PMC DMRef Master ShortList.xml"
$KC46DataRoot = "F:\KC46 Staging"
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"
#Book|Type|Book|PB Status|Revision|Rev Type|Rev Date|Highlight|Navigation Tree Title|Task Title|Pageblock|C&V|PB/Task|DMC|IC Review|DM Type|Review	Type|TechName|InfoName|Ver List|Cert Complete|Pblock|Support Equipment/Notes|C|V|USAF Rep|ETOPS?|Daily Plng|Proc Avail|Cert Interim Comp|Cert Comp|Ver Interim Comp|Ver Comp|AC Type|CAS Incorp|USAF Printed (P) - Delivered(Date)|Tasks Too Difficult To Do|Chapter|Section|Subject|Info1|Comb1|Info2|Comb2|PB	Info3|Prep Time|Est Duration (Adtnl)|Post Time|Ver Type|Rationale|AMC Recommendation|Boeing Review of AF Recommendation|Done|Boeing Recommendation|Original Ver Type|Original Ver List|Original Rationale|CertTeamRecord

[string[]] $bookList   = @("KC46","AMM","ARD","FIM","IPB","NDT","SSM","SWPM","TC","WDM" ) | Sort-Object

$masterListLong = @()
$masterListShort = @()
$masterListLong += "BookType|Book|DMC|IssueNumber|PB|Task|FigID|FigTitle|GraphicID|ICN|GraphicTitle|MIC|CH|SE|SU|DC|DCV|IC|ICV|ILC|ICR|ModuleType|firstVerification|secondVerification|IssueNumber_VerfiedByUSAF"
$masterListShort += "Book|DMC|IssueNumber|PB|Task|ModuleType|firstVerification|secondVerification|IssueNumber_VerfiedByUSAF"

$BookType = ""
$BookType1 = "Maintenance"
$BookType2 = "Flight"

foreach( $book in $bookList)
    {
    
    if($book -eq "BCLM")
    {
        $BookType = $BookType2
    }
    else
    {
        $BookType = $BookType1
    }

    if($book -ne "IPB")
    {
        $pmcColl = gci -Path "$source_BaseLocation\$book\S1000D\SDLLIVE\PMC*.xml"
    }
    else
    {
        $pmcColl = gci -Path "f:\KC46 Staging\Production\Manuals\IPB\s1000d\SDLLIVE\PMC-1KC46-81205-P*.xml"
    }

    foreach ($pmc in $pmcColl)
    {
        $pm = [xml] (Get-Content -Path $pmc.FullName)
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

            $fileName =  "$source_BaseLocation\$book\S1000D\SDLLIVE\DMC-$modelIdentCode-$systemDiffCode-$systemCode-$subSystemCode$subSubSystemCode-$assyCode-$disassyCode$disassyCodeVariant-$infoCode$infoCodeVariant-$itemLocationCode"
            $wcfn = $fileName + "`*.xml"
            
            Try
            {
                $fi = gci -Path $wcfn -Force -Recurse -Verbose 
                Try
                {
                    $FN = $fi[0].FullName
                    $SN = $fi[0].Name
                }
                Catch
                {
                   "Oops:`t" + $FN
                   break
                }
                
                #$FN
                $dm = New-Object System.Xml.XmlDocument
                $dm.Load($FN) 

                $firstVerification = $dm.dmodule.identAndStatusSection.dmStatus.qualityAssurance.firstVerification.verificationType
                $secondVerification = $dm.dmodule.identAndStatusSection.dmStatus.qualityAssurance.secondVerification.verificationType
                $issueNumber = $dm.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.issueNumber
                $figRefs = $dm.SelectNodes("//figure")
                
                $mpNode = $dm.dmodule.content.procedure.mainprocedure
                if($mpNode.ChildNodes.Count -gt 0)
                {
                    $masterListShort += "$book|" + $SN +  "|" + $issueNumber + "|" + $lastPB + "|" + $pborTask + "|" + "procdure" + "|" + $firstVerification + "|" +  $secondVerification +  "|"                 
                    foreach($fig  in $figRefs)
                    {
                        $graphics = $fig.SelectNodes("graphic")                    
                        foreach($g  in $graphics)
                        {
                            $masterListLong += "$book|" + $SN +  "|" +  $issueNumber + "|" + $lastPB + "|" + $pborTask + "|" + $fig.id + "|" + $fig.title + "|" + $g.id+ "|" + $g.infoEntityIdent + "|" + $g.title + "|" + $modelIdentCode + "|" + $ch + "|" + $se + "|" + $su + "|" + $disassyCode + "|" + $disassyCodeVariant + "|" + $infoCode + "|" + $infoCodeVariant + "|" + $itemLocationCode + "|" + ([int] $infoCode.Substring(0,1) * 100) + "|" + "procedure" + "|" + $firstVerification + "|" +  $secondVerification+  "|"
                        }
                    }
                }
                else
                {
                    $masterListShort += "$book|" + $SN +  "|" + $issueNumber + "|" + $lastPB + "|" + $pborTask + "|" + "description" + "|" + $firstVerification + "|" +  $secondVerification +  "|"
                }

            }
            Catch
            {
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                $FailedItem
                
            }
        }
    }
}

$masterListLong | Out-File $outputReportPath1 -Force
$masterListShort | Export-Clixml $outputReportPath2 -Force
$masterListShort | Out-File "C:\KC46 Staging\Scripts\Report Generators\Outputs\PMC DMRef Master ShortList.csv"

$error