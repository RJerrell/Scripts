cls
cls
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$dmParser = new-object -TypeName S1000D.DataModule_401
$pmParser = new-object -TypeName S1000D.PublicationModule_401


# CHANGE THIS LOCATION FOR THE REPORT FILE
$reportPath = "C:\TEMP\VALIDATEDMREFS.CSV"
$basePath = "C:\KC46 Staging\Production\Manuals"

 [string[]] $PubList = @("AMM","ARD", "FIM", "IPB", "LOAPS", "MOM", "NDI", "NDT", "SPCC", "SRM", "SSM", "SWPM", "WUC")
[string[]] $PubList = @("ABDR","LOAPS","SIMR","WUC")

$a1=@() # Bad PMC Entries
$b1=@() # Good PMC Entries
$c1=@() # Bad DMC Entries
$d1=@() # Good DMC Entries
$e1=@() # Bad ICN Entries
$f1=@() # Good ICN Entries

foreach( $pub in $PubList)
{
    $manualBasePath = "$basePath\$pub"
    $manualDocs = "$manualBasePath\S1000D\SDLLIVE"
    $manulIllustrations = "$manualBasePath\Illustrations\Illustrations"
    $pmcPath = "$manualDocs\pmc*.xml"
    $pmcFiles = gci -Path $pmcPath | Sort-Object -Descending | Select-Object -First 1
    $pmc = New-Object System.Xml.XmlDocument
    $pmc.Load($pmcFiles[0].Fullname)
    $dmRefs = $pmc.SelectNodes("//dmRef");
    foreach( $dmRef in $dmRefs )
    {
        $filePref = "DMC";
        $RefFilename = Get-FilenameFromDMRef -dmRef $dmRef -filePref $filePref
        $bookReferenced = Get-DocTypeFromDMC -dc $RefFilename
        

        $fs = gci -path "$manualDocs\$RefFilename`*" |Sort-Object -Descending |Select-Object -First 1
        
        if($fs.Count -eq 0)
        {       
            $a1 += "$manualDocs\$RefFilename"
        }
        else
        {
            $b1 += "$manualDocs\$RefFilename"
        }
    }
    if($a1.Count -eq 0)
    {
        $bookReferenced + ": PMC is good"
    }
    <# 
        Now, we will process each referenced dm and determine if the quality of each.
    #>
}

foreach ($fileName in $b1)
{
$Pub = ""
    $dms = gci -path "$fileName`*" |Sort-Object -Descending |Select-Object -First 1
    $dmc = New-Object System.Xml.XmlDocument
    $dmc.Load($dms[0].FullName)
    $dmRefs = $dmc.SelectNodes("/dmodule/content//dmRef");
    $icns = $dmc.SelectNodes("/dmodule/content//graphic");
    foreach( $dmRef in $dmRefs )
    {
        $RefFilename = Get-FilenameFromDMRef -dmRef $dmRef -filePref $filePref
        $bookReferenced = Get-DocTypeFromDMC -dc $RefFilename
        $Pub = Get-DocTypeFromDMC -dc $fileName
        $fullFilePath2ToRefDoc = "$basePath\$bookReferenced\S1000D\SDLLIVE"
        $fs = gci -path "$fullFilePath2ToRefDoc\$RefFilename`*" |Sort-Object -Descending |Select-Object -First 1
        if($fs.Count -eq 0)
        {       
            $c1 += $Pub + "`t"  + $fileName + "`t" + $bookReferenced + "`t" + "$fullFilePath2ToRefDoc\$RefFilename"
        }
        else
        {
            $d1 += $Pub + "`t"  + $fileName + "`t" + $bookReferenced + "`t" + "$fullFilePath2ToRefDoc\$RefFilename"
        }
    }
    foreach ($icn in $icns)
    {
        $icnFile = $icn.infoEntityIdent + ".CGM"
        if(!(Test-Path -Path "$manulIllustrations\$icnFile"))
        {
            $e1 += $Pub + "`t"  + $fileName + "`t" + $bookReferenced + "`t" + "$manulIllustrations\$icnFile"
        }
        else
        {
            $f1 += $Pub + "`t"  + $fileName + "`t" + $bookReferenced + "`t" + "$manulIllustrations\$icnFile" 
        }
    }
    
}
$a1 | Ft
$c1 | Ft
$e1 | Ft

