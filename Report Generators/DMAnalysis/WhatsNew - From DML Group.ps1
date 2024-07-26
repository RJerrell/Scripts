
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

$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"

Import-Module -Name "KC46Common" -Verbose -Force

# *****************************************************************************************************
$pubBasePath = "C:\KC46 Staging\production\Manuals"
$dmlBasePath = "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set"
$ipbPMCPath = "$pubBasePath\IPB\s1000d\SDLLive\PMC-1KC46-81205-P*.xml"
$pathToDMLs = "$dmlBasePath\DML*.XML"

$reportpath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$rptBaseName = "Highlights"
$rptSuffix = "- Release 6 to Release 7"
$nFiles=  @()
$cFiles=  @()
$dFiles=  @()
$allFiles =  @()
$dmls = gci -Path $pathToDMLs
foreach ($dml in $dmls)
{
    #region Prefixes
    if($dml.Name.StartsWith("DML-1KC46-AAA0A"))
    {
        $pub = "AMM"
    }
    elseif($dml.Name.StartsWith("DML-1KC46-AAA0B"))
    {
        $pub = "BCLM"
    }
    elseif($dml.Name.StartsWith("DML-1KC46-AAA0F"))
    {
        $pub = "FIM"
    }
    elseif($dml.Name.StartsWith("DML-1KC46-AAA0R"))
    {
        $pub = "SSM"
    }
    elseif($dml.Name.StartsWith("DML-1KC46-AAA0T"))
    {
        $pub = "TC"
    }
    elseif($dml.Name.StartsWith("DML-1KC46-AAA0V"))
    {
        $pub = "NDT"
    }
    elseif($dml.Name.StartsWith("DML-1KC46-AAA0W"))
    {
        $pub = "WDM"
    }
    elseif($dml.Name.StartsWith("DML-1KC46-AAA0Z"))
    {
        $pub = "KC46"
    }
    elseif($dml.Name.StartsWith("DML-1KC46-AAA0S"))
    {
        $pub = "SRM"
    }
    elseif($dml.Name.StartsWith("DML-1KC46-81205"))
    {
        $pub = "IPB"
    }    
    
    <#
        $path2PMC = "C:\KC46 Staging\production\Manuals\$pub\S1000D\SDLLIVE\PMC*.xml"

        $pmcS = gci -Path $path2PMC | Sort-Object -Descending | Select -First 1

        $dml.Name
    
        $pm = New-Object System.Xml.XmlDocument
    
        if($pmcS.Count -eq 1)
        {
            $pm = $pm.Load($pmcS[0].FullName)
        }
    #>

    #endregion

    $dmlXML = New-Object System.Xml.XmlDocument
    $dmlXML.Load($dml.FullName)
    $dmEntries = $dmlXML.SelectNodes("/dml/dmlContent/dmEntry")
    "$pub`t " + $dmEntries.Count
    foreach ($dmEntry in $dmEntries)
    {
        $dmc = $dmEntry.dmRef.dmRefIdent.dmCode
        $modelIdentCode = [string] $dmc.modelIdentCode
        $systemDiffCode = [string]  $dmc.systemDiffCode
        $systemCode = [string]  $dmc.systemCode
        $subSystemCode = [string]  $dmc.subSystemCode
        $subSubSystemCode = [string]  $dmc.subSubSystemCode
        $assyCode = [string]  $dmc.assyCode
        $disassyCode = [string]  $dmc.disassyCode
        $disassyCodeVariant = [string]  $dmc.disassyCodeVariant
        $infoCode = [string]  $dmc.infoCode
        $infoCodeVariant = [string]  $dmc.infoCodeVariant
        $itemLocationCode = [string]  $dmc.itemLocationCode
        $ch = [string]  $systemCode
        $se = [string]  $subSystemCode+$subSubSystemCode
        $su = [string]  $assyCode           
        $dataModuleCode =  Get-FilenameFromDMRef -dmRef $dmEntry.dmRef -filePref "DMC"
        $techName = [string]  $dmEntry.dmRef.dmRefAddressItems.dmTitle.techName
        $InfoName = [string]  $dmEntry.dmRef.dmRefAddressItems.dmTitle.infoName
        
        $type = [string]  $dmEntry.dmEntryType
        <#
            The report above only tells us the information provided by the DML files we recieve.
            We need, during the C&V and IPR processes, to know the PB/Task information from ATA 2000
            So, we have to take each dataModule, go get its PMC, find the PB/TASK information in it, and append it to each entry
        #>
        
        $objsec = [pscustomobject][ordered]@{Type=$type;Manual=$pub;DMC=$dataModuleCode;TechName=$techName;InfoName=$InfoName;}
        if($type -eq "n")
        {
            $type = "NEW"            
            $nFiles += $objsec
        }
        elseif($type -eq "c")
        {
            $type = "MOD"            
            $cFiles += $objsec
        }
        elseif($type -eq "d")
        {
            $type = "DEL"            
            $dFiles += $objsec            
        }
        else
        {
            "something really bad happened"
            exit
        }
        
        $allFiles += $objsec
        $objsec    = $null        
    }
}


$prop1 = @{Expression='Manual'; Ascending=$true }
$prop2 = @{Expression='DMC'; Ascending=$true }

$allFiles.GetEnumerator() | Sort-Object -Property $prop1, $prop2 | Export-Csv "$reportpath\$rptBaseName $rptSuffix.csv" -NoTypeInformation -Encoding UTF8
$nFiles.GetEnumerator() | Sort-Object -Property $prop1, $prop2   | Export-Csv "$reportpath\$rptBaseName $rptSuffix - New.csv" -NoTypeInformation -Encoding UTF8
$cFiles.GetEnumerator() | Sort-Object -Property $prop1, $prop2   | Export-Csv "$reportpath\$rptBaseName $rptSuffix - Changed.csv" -NoTypeInformation -Encoding UTF8
$dFiles.GetEnumerator() | Sort-Object -Property $prop1, $prop2   | Export-Csv "$reportpath\$rptBaseName $rptSuffix - Deleted.csv" -NoTypeInformation -Encoding UTF8

"The reports awaits you: " + "`r`n$reportpath\$rptBaseName $rptSuffix.csv`r`n"
"$reportpath\\$rptBaseName $rptSuffix - New.csv `r`n"
"$reportpath\\$rptBaseName $rptSuffix - Changed.csv `r`n"
"$reportpath\\$rptBaseName $rptSuffix - Deleted.csv `r`n"

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"
