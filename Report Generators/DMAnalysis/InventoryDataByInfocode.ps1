
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
$ErrorActionPreference = "SilentlyContinue"
$error.Clear()
$a = (Get-Host).PrivateData
$a.WarningBackgroundColor = "Yellow"
$a.WarningForegroundColor = "Black"
$rn = "<br/>"
$a.ErrorBackgroundColor = "red"
$a.ErrorForegroundColor = "White"
# *****************************************************************************************************
"Start"
$basePath = "F:\KC46 Staging\Production\Manuals"

[string[]] $PubList = @("KC46","AMM","ARD", "FIM", "IPB", "LOAPS", "MOM", "NDI", "NDT", "SPCC", "SRM", "SSM", "SWPM", "WDM")

$hash = @{}

foreach( $pub in $PubList)
{
    $pmcName = $pmcHash.Get_Item($pub)
    $pmcPath = "$basePath\$pub\S1000D\SDLLIVE\$pmcName"
    $PMCS = gci -Path "$basePath\$pub\S1000D\SDLLIVE"

    $pmc = New-Object System.Xml.XmlDocument
    $pmc.Load($pmcPath)
    $dmcCollection = $pmc.SelectNodes("//dmRef");
    foreach( $dmc in $dmcCollection )
    {
        $filePref = "DMC";
        # <dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="23" subSystemCode="5" subSubSystemCode="1" assyCode="0600" disassyCode="01" disassyCodeVariant="A0A" infoCode="010" infoCodeVariant="A" itemLocationCode="A" />
        $modelIdentCode = $dmc.dmRefIdent.dmCode.modelIdentCode
        $systemDiffCode = $dmc.dmRefIdent.dmCode.systemDiffCode
        $systemCode = $dmc.dmRefIdent.dmCode.systemCode
        $subSystemCode = $dmc.dmRefIdent.dmCode.subSystemCode
        $subSubSystemCode = $dmc.dmRefIdent.dmCode.subSubSystemCode
        $assyCode = $dmc.dmRefIdent.dmCode.assyCode
        $disassyCode = $dmc.dmRefIdent.dmCode.disassyCode
        $disassyCodeVariant = $dmc.dmRefIdent.dmCode.disassyCodeVariant
        $infoCode = $dmc.dmRefIdent.dmCode.infoCode
        $infoCodeVariant = $dmc.dmRefIdent.dmCode.infoCodeVariant
        $itemLocationCode = $dmc.dmRefIdent.dmCode.itemLocationCode 

        if($hash.ContainsKey($infoCode))
        {
            $hash.Set_Item($infoCode, $hash.Get_Item($infoCode) + 1)
                
        }
        else
        {
            $hash.Add("$infoCode", 1)
        }
    }
        
    $OutputTable = $hash.GetEnumerator() |  
    % { 
        New-Object PSObject -Property ([ordered] @{InfoCode = $_.Name;Count = [string] $_.Value} )    
    }
        
    $OutputTable | Export-CSV "C:\KC46 Staging\Scripts\Report Generators\Outputs\$pub.csv" -NoTypeInformation

    $hash.Clear()
}  
"Complete"

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"
