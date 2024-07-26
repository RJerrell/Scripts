cls

# CHANGE THIS LOCATION FOR THE REPORT FILE
$reportPath = "C:\TEMP\VALIDATEDMREFS.CSV"
<# Set the basic path to the data with the data in a folder structure like this
 
 C:\KC46 Staging\Production\Manuals\AMM\S1000D\S1000D (DMC and PMC files here
                                   \ARD\S1000D\S1000D (DMC and PMC files here
                                   \FIM\S1000D\S1000D (DMC and PMC files here
                                   \NDT\S1000D\S1000D (DMC and PMC files here
                                   \SRM\S1000D\S1000D (DMC and PMC files here
                                   \SSM\S1000D\S1000D (DMC and PMC files here                                            
                                   \WDM\S1000D\S1000D (DMC and PMC files here
#>

$basePath = "C:\KC46 Staging\BetaTR\Manuals"
# [string[]] $PubList = @("AMM","ARD", "FIM", "IPB", "LOAPS", "MOM", "NDI", "NDT", "SPCC", "SRM", "SSM", "SWPM", "WDM")
[string[]] $PubList = @("AMM","ARD", "FIM", "NDT", "SRM","SSM", "WDM")
#[string[]] $PubList = @("AMM", "FIM")
    $pmcHash = @{
    "KC46"="PMC-1KC46-81205-Z0000-00.xml";
    "AMM"="PMC-1KC46-81205-A0000-00.xml";
    "ARD"="PMC-1KC46-81205-E0000-00";
    "FIM"="PMC-1KC46-81205-F0000-00.xml";
    "IPB"="PMC-1KC46-81205-P0006-00.XML";
    "LOAPS"="PMC-1KC46-81205-M1000.XML";
    "MOM"="PMC-1KC46-81205-N0000.XML";
    "NDI"="PMC-1KC46-81205-J0000.XML";
    "NDT"="PMC-1KC46-81205-V0000-00.xml";
    "SPCC"="PMC-1KC46-81205-K0000.XML";
    "SRM"="PMC-1KC46-81205-S0000-00.xml";
    "SSM"="PMC-1KC46-81205-R0000-00.xml";
    "SWPM"="PMC-1KC46-81205-U0000-01.xml";
    "WDM"="PMC-1KC46-81205-W0000-00.xml";
    }

    foreach( $pub in $PubList)
    {
    $a1=@()
    $b1=@()

    $pmcName = $pmcHash.Get_Item($pub)
    $pmcPath = "$basePath\$pub\S1000D\S1000D\$pmcName"
    $pmc = [XML] (Get-Content -Path $pmcPath)
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

        [bool] $existsOnDisk = $false
        # DMC-1KC46-A-00-00-0000-01A0A-018A-A.xml
        
        $fileName =  "$basePath\$pub\S1000D\S1000D\$filePref-$modelIdentCode-$systemDiffCode-$systemCode-$subSystemCode$subSubSystemCode-$assyCode-$disassyCode$disassyCodeVariant-$infoCode$infoCodeVariant-$itemLocationCode.xml"
        $fileName2 =  "$filePref-$modelIdentCode-$systemDiffCode-$systemCode-$subSystemCode$subSubSystemCode-$assyCode-$disassyCode$disassyCodeVariant-$infoCode$infoCodeVariant-$itemLocationCode.xml"
        $a1 += $fileName2
        
        
        $existsOnDisk = (Test-Path -Path $fileName)
         if($existsOnDisk -eq $false)
        {       
            "Missing data module :`t$fileName"
        }

        # "$existsOnDisk :`t$fileName"
    }
    
    <# 
        Now, we will process each S1000D folder and compare the dmCode to the PMC for that
        manual to see if there are files
        that exist on disk that are not referenced.
    #>
    $filePath = "$basePath\$pub\S1000D\S1000D\DMC*.XML"
    $files = Get-ChildItem -Path $filePath
    foreach($file in $files)
    {
        $b1 += $file.Name
    }
    " ********************************  START OF $pub REPORT          **********************************`n"
        "TOTAL PMC ENTRIES`t`t`t`t: " + $a1.Count
        "TOTAL DATA MODULES ON DISK IN CSDB`t`t`t`t: " + $b1.Count
        " DIFFERENCE REPORT BETWEEN THE 2 LISTS: PMC Entries versus CSDB entries"
        "<= indicates a reference in the PMC that does not have an equivalent file on disk`n=> indicates a file on disk that is not referenced in the PCM`n"    

    Compare-Object $a1 $b1  | Format-Table
    Compare-Object $a1 $b1  | Export-Csv $reportPath
   
		 "TOTAL PMC ENTRIES`t`t`t`t: " + $a1.Count
        "TOTAL DATA MODULES ON DISK IN CSDB`t`t`t`t: " + $b1.Count
        " DIFFERENCE REPORT BETWEEN THE 2 LISTS: PMC Entries versus CSDB entries"
        "<= indicates a reference in the PMC that does not have an equivalent file on disk`n=> indicates a file on disk that is not referenced in the PCM`n"

    " ********************************    END OF $pub REPORT          **********************************`n"
}