<#
This script produces 2 reports-- 1 CSV and 1 XML with the former being a boring CSV listing for spreadsheet 
users and the other a shortlist used as an input to other programs / scripts.

CSV Report location - "C:\KC46 Staging\scripts\Report Generators\Outputs\PMC DMRef Master LongList.CSV"
XML Report Location - "C:\KC46 Staging\scripts\Report Generators\Outputs\PMC DMRef Master ShortList.xml"

            Uncomment line 20 to run this report for all manuals listed on that line!
#>
cls

# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'

$ErrorActionPreference = "Stop"
$error.Clear()
Import-Module -Name "KC46Common" -Verbose -Force

Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force 
# *****************************************************************************************************
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$parserDM = new-object -TypeName S1000D.DataModule_401
$parserPM = new-object -TypeName S1000D.PublicationModule_401

$environment = "Production"
$outputReportPath1 = "C:\KC46 Staging\scripts\Report Generators\Outputs\S1000D Data Metrics.csv"
$outputReportPath2 = "C:\KC46 Staging\scripts\Report Generators\Outputs\S1000D Data Metrics.xml"
$KC46DataRoot = "F:\KC46 Staging"
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"
#Book|Type|Book|PB Status|Revision|Rev Type|Rev Date|Highlight|Navigation Tree Title|Task Title|Pageblock|C&V|PB/Task|DMC|IC Review|DM Type|Review	Type|TechName|InfoName|Ver List|Cert Complete|Pblock|Support Equipment/Notes|C|V|USAF Rep|ETOPS?|Daily Plng|Proc Avail|Cert Interim Comp|Cert Comp|Ver Interim Comp|Ver Comp|AC Type|CAS Incorp|USAF Printed (P) - Delivered(Date)|Tasks Too Difficult To Do|Chapter|Section|Subject|Info1|Comb1|Info2|Comb2|PB	Info3|Prep Time|Est Duration (Adtnl)|Post Time|Ver Type|Rationale|AMC Recommendation|Boeing Review of AF Recommendation|Done|Boeing Recommendation|Original Ver Type|Original Ver List|Original Rationale|CertTeamRecord

[string[]] $bookList   = @("AMM", "ARD", "FIM", "NDT", "SSM", "WDM")

$masterListShort = @()
$masterListShort += "Book|DMC|TechName|InfoName|IssueNumber|PB|Task|CH|SE|SU|DMType|BDS_COC_Count|CAS_COC_Count"
foreach( $book in $bookList)
    {    
    
    if($book -ne "IPB")
    {
        $pmcColl = gci -Path "$source_BaseLocation\$book\S1000D\SDLLIVE\PMC*.xml"
    }
    else
    {
        $pmcColl = gci -Path "F:\KC46 Staging\Production\Manuals\IPB\s1000d\SDLLIVE\PMC-1KC46-81205-P*.xml"
    }

    foreach ($pmc in $pmcColl)
    {
        #$pm = [xml] (Get-Content -Path $pmc.FullName)
        $parserPM.ParsePM($pmc.FullName)
        $dmcCollection = $parserPM.DmRefs
      
        $lastPB = ""
        $pborTask= "blank"
        for ($i = 0; $i -lt $dmcCollection.Count; $i++)
        { 
            [STRING] $modelIdentCode = $parserPM.DmRefs[$i].dmRefIdent.dmCode.modelIdentCode
            [STRING] $systemDiffCode = $parserPM.DmRefs[$i].dmRefIdent.dmCode.systemDiffCode
            [STRING] $systemCode = $parserPM.DmRefs[$i].dmRefIdent.dmCode.systemCode
            [STRING] $subSystemCode = $parserPM.DmRefs[$i].dmRefIdent.dmCode.subSystemCode
            [STRING] $subSubSystemCode = $parserPM.DmRefs[$i].dmRefIdent.dmCode.subSubSystemCode
            [STRING] $assyCode = $parserPM.DmRefs[$i].dmRefIdent.dmCode.assyCode
            [STRING] $disassyCode = $parserPM.DmRefs[$i].dmRefIdent.dmCode.disassyCode
            [STRING] $disassyCodeVariant = $parserPM.DmRefs[$i].dmRefIdent.dmCode.disassyCodeVariant
            [STRING] $infoCode = $parserPM.DmRefs[$i].dmRefIdent.dmCode.infoCode
            [STRING] $infoCodeVariant = $parserPM.DmRefs[$i].dmRefIdent.dmCode.infoCodeVariant
            [STRING] $itemLocationCode = $parserPM.DmRefs[$i].dmRefIdent.dmCode.itemLocationCode
            [STRING] $ch = $systemCode
            [STRING] $se = $subSystemCode+$subSubSystemCode
            [STRING] $su = $assyCode
            [STRING] $techName = $parserPM.DmRefs[$i].title
            
            $title = $parserPM.DmRefs[$i].title
            
            $pborTask = $parserPM.DmRefs[$i].href
            
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
                #$dm = New-Object System.Xml.XmlDocument
                #$dm.Load($FN) 

                $parserDM.ParseDM($FN)

                $issueNumber = $parserDM.IssueInfo.issueNumber
                # /dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle/techName
                $techName = $parserDM.TechName
                $infoName = $parserDM.InfoName
                $dm = $parserDM.Dmodule
                $BDS_COCs = ($dm.DocumentElement.SelectNodes("/dmodule/content//*[starts-with(@id,`"KC46`") or starts-with(@internalRefId,`"acr`") or starts-with(@id,`"acr`") ]")).Count
                $CAS_COCs = ($dm.DocumentElement.SelectNodes("/dmodule/content//*[@authorityName=`"COC`" and not(starts-with(@internalRefId,`"acr`")) and not(starts-with(@id,`"acr`")) ]")).Count

                $mpNode = $dm.dmodule.content.procedure.mainprocedure

                if($mpNode.ChildNodes.Count -gt 0)
                {
                    # "Book|DMC|IssueNumber|PB|Task|CH|SE|SU|DMType|BDS_COC_Count|CAS_COC_Count"
                    $masterListShort += "$book|" + $SN  +  "|" + $techName +  "|" + $infoName +  "|" + $issueNumber + "|" + $lastPB + "|" + $pborTask+ "|" +  $ch + "|" + $se + "|" + $su + "|" + "procdure" + "|" + $BDS_COCs + "|" +  $CAS_COCs
                }
                else
                {
                     $masterListShort += "$book|" + $SN  +  "|" + $techName +  "|" + $infoName +  "|" + $issueNumber + "|" + $lastPB + "|" + $pborTask+ "|" +  $ch + "|" + $se + "|" + $su + "|" + "description" + "|" + $BDS_COCs + "|" +  $CAS_COCs
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

$masterListShort | Export-Clixml $outputReportPath2 -Force
$masterListShort | Out-File $outputReportPath1

$error