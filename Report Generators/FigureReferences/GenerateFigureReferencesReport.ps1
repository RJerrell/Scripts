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
$env:PSModulePath

Import-Module -Name "KC46Common" -Verbose -Force

$environment = "Production"
$path = "C:\KC46 Staging\Scripts\Report Generators\FigureReferences"
$KC46DataRoot = "C:\KC46 Staging"
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"

[string[]] $PubList   = @("AMM")
$fiveTwentyWithoutA720 = @()
$all520s = @()
$masterList = @("")
$masterList += "TechName`tDMC`tPB`tTask`tFigID`tFigTitle`tGraphicID`tICN`tGraphicTitle`tMIC`tCH`tSE`tSU`tDC`tDCV`tIC`tICV`tILC`tICR"
$BookType = ""
$BookType1 = "Maintenance"
$BookType2 = "Flight"

foreach( $pub in $PubList)
    {
    
    if($pub -eq "BCLM")
    {
        $BookType = $BookType2
    }
    else
    {
        $BookType = $BookType1
    }

    $pmcColl = gci -Path "$source_BaseLocation\$pub\S1000D\SDLLIVE\PMC*.xml"

    foreach ($pmc in $pmcColl)
    {
        $pm = [xml] (Get-Content -Path $pmc.FullName)
        $dmcCollection = $pm.SelectNodes("//dmRef")
     
        $lastPB = ""
        $pborTask= "blank"

        foreach( $dmc in $dmcCollection )
        {   
            [STRING] $modelIdentCode = $dmc.dmRefIdent.dmCode.modelIdentCode
            [STRING] $systemDiffCode = $dmc.dmRefIdent.dmCode.systemDiffCode
            [STRING] $systemCode = $dmc.dmRefIdent.dmCode.systemCode
            [STRING] $subSystemCode = $dmc.dmRefIdent.dmCode.subSystemCode
            [STRING] $subSubSystemCode = $dmc.dmRefIdent.dmCode.subSubSystemCode
            [STRING] $assyCode = $dmc.dmRefIdent.dmCode.assyCode
            [STRING] $disassyCode = $dmc.dmRefIdent.dmCode.disassyCode
            [STRING] $disassyCodeVariant = $dmc.dmRefIdent.dmCode.disassyCodeVariant
            [STRING] $infoCode = $dmc.dmRefIdent.dmCode.infoCode
            [STRING] $infoCodeVariant = $dmc.dmRefIdent.dmCode.infoCodeVariant
            [STRING] $itemLocationCode = $dmc.dmRefIdent.dmCode.itemLocationCode
            [STRING] $ch = $systemCode
            [STRING] $se = $subSystemCode+$subSubSystemCode
            [STRING] $su = $assyCode
            
            $title = $dmc.Attributes[0].Value
            $pborTask = $dmc.Attributes[1].Value
            
            if($pborTask.Contains("PAGEBLOCK"))
            {
                $lastPB = $pborTask
            }
            $fileName =  "$source_BaseLocation\$pub\S1000D\SDLLIVE\DMC-$modelIdentCode-$systemDiffCode-$systemCode-$subSystemCode$subSubSystemCode-$assyCode-$disassyCode$disassyCodeVariant-$infoCode$infoCodeVariant-$itemLocationCode.xml"
            
            $fi = gci -Path $fileName
            $shortName = $fi.Name

            if($infoCode -eq "520")
            {
                $infoCode
                $TitleArray520 = (Get-TechNameInfoNameFromDMC -fileName $fileName)
                
                $all520s += "$shortName|" + $TitleArray520[0] + "|" + $TitleArray520[1]
                $mateBase = "$source_BaseLocation\$pub\S1000D\SDLLIVE\DMC-$modelIdentCode-$systemDiffCode-$systemCode-$subSystemCode$subSubSystemCode-*720A*"
                $mateFiles = gci -Path $mateBase
                if($mateFiles.Count -eq 0)
                {                        
                    $fiveTwentyWithoutA720 += "$shortName|" +  $TitleArray520[0] + "|" + $TitleArray520[1]
                }
                elseif($mateFiles.Count -gt 0)
                {                         
                    foreach ($mateFile in $mateFiles)
                    {
                        $dmcNode = Create-DMCCodeElementFromDMCFileName -dmCode $mateFile.Name
                        $dmcNode.dmCode.infoCode
                        $TitleArray720    = (Get-TechNameInfoNameFromDMC -fileName $mateFile.fullname)
                            
                        $shortTitlePos520 = $TitleArray520[0].LastIndexOf("Removal")
                        $shortTitlePos720 = $TitleArray720[0].LastIndexOf("Inst")
                            
                        $titlePrefix520   = $TitleArray520[0].Substring(0,$shortTitlePos520)
                        $titlePrefix720   = $TitleArray720[0].Substring(0,$shortTitlePos720)

                        if($titlePrefix520 -eq $titlePrefix720)
                        {
                            if (Get-Content $fileName | Select-String -Pattern "ICN-81205-K")
                            {
                                $dm520 = [xml](Get-Content -Path $fileName)
                                $dm720 = [xml](Get-Content -Path $mateFile.FullName)

                                $figRefs520 = $dm520.SelectNodes("//figure")
                                $figRefs720 = $dm720.SelectNodes("//figure")
                                $temp520Figs = @()
                                $temp720Figs = @()
                                foreach($fig  in $figRefs520)
                                {
                                    $temp520Figs += $fig.id
                                }
                                foreach($fig  in $figRefs720)
                                {
                                    $temp720Figs += $fig.id
                                }
                                $areEqual = @(Compare-Object $temp520Figs $temp720Figs).Length -eq 0
                                $areEqual
                                if( $areEqual -eq $false )
                                {
                                    foreach($fig  in $figRefs520)
                                    {
                                        $graphics = $fig.SelectNodes("graphic")
                                        foreach($g  in $graphics)
                                        {
                                            $masterList += $TitleArray520[0] + "`t" + $fi.Name + "`t" + $lastPB + "`t" + $pborTask + "`t" + $fig.id + "`t" + $fig.title + "`t" + $g.id+ "`t" + $g.infoEntityIdent + "`t" + $g.title + "`t" + $modelIdentCode + "`t" + $ch + "`t" + $se + "`t" + $su + "`t" + $disassyCode + "`t" + $disassyCodeVariant + "`t" + $infoCode + "`t" + $infoCodeVariant + "`t" + $itemLocationCode + "`t" + ([int] $infoCode.Substring(1,1) * 100)
                                        }
                                    }
                                    if($figRefs720.Count -gt 0)
                                    {
                                        foreach($fig  in $figRefs720)
                                        {
                                            $graphics = $fig.SelectNodes("graphic")                     
                                            foreach($g  in $graphics)
                                            {
                                                $masterList += $TitleArray720[0] + "`t" + $mateFile.Name + "`t" + "" + "`t" + "" + "`t" + $fig.id + "`t" + $fig.title + "`t" + $g.id+ "`t" + $g.infoEntityIdent + "`t" + $g.title + "`t" + $dmcNode.dmCode.modelIdentCode + "`t" + "" + "`t" + "" + "`t" + "" + "`t" + $dmcNode.dmCode.disassyCode + "`t" + $dmcNode.dmCode.disassyCodeVariant + "`t" + $dmcNode.dmCode.infoCode + "`t" + $dmcNode.dmCode.infoCodeVariant + "`t" + $dmcNode.dmCode.itemLocationCode + "`t" + ([int] ($dmcNode.dmCode.infoCode).Substring(1,1) * 100)
                                            }
                                        }
                                    }
                                    else
                                    {
                                        $masterList += $TitleArray720[0] + "`t" + $mateFile.Name + "`t" + "" + "`t" + "" + "`t" + "No graphics" + "`t" + "" + "`t" + "" + "`t" + "" + "`t" + "" + "`t" + $modelIdentCode + "`t" + $ch + "`t" + $se + "`t" + $su + "`t" + $disassyCode + "`t" + $disassyCodeVariant + "`t" + $infoCode + "`t" + $infoCodeVariant + "`t" + $itemLocationCode + "`t" + ([int] $infoCode.Substring(1,1) * 100)
                                    }
                                }
                            }
                        }

                    }                                              
                }                    
            }
            else
            {
                #$infoCode
            }

        }
    }
}

$masterList | Out-File "C:\KC46 Staging\Scripts\Report Generators\Outputs\Figure References.CSV"
$fiveTwentyWithoutA720 | Sort-Object | Out-File "C:\KC46 Staging\Scripts\Report Generators\Outputs\Figure References - 520 WITHOUT 720.CSV"
$all520s | Sort-Object | Out-File "C:\KC46 Staging\Scripts\Report Generators\Outputs\Figure References - All 520s.CSV"