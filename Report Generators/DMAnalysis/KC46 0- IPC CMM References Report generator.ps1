<#
Author: Roger Jerrell
Date Created: Creates a listing of all the CMMs that are referenced from within the IPB
Purpose: Provides a unique list of all the CMM references in the IPB
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
$pathToDatamodules = "F:\KC46 Staging\Production\Manuals\IPB\S1000D\sdllive\DMC*.XML"
$dms = gci -Path $pathToDatamodules

$reportpath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$rptBaseName = "IPC CMM Reference Listing"
$rptSuffix1 = "- Sorted by CMM"
$rptSuffix2 = "- Sorted by PN"

$allParts =  @()
$allCMMs= @()
$objsec = $null
$dmXML = New-Object System.Xml.XmlDocument
foreach ($dm in $dms)
{
    $dmXML.Load($dm.FullName)
    $cmmRefs = $dmXML.SelectNodes("/dmodule/content/illustratedPartsCatalog/catalogSeqNumber/itemSequenceNumber/genericPartDataGroup/genericPartData[@genericPartDataName='COMPONENT MAINT MANUAL REF:']")
    foreach ($cmmRef in $cmmRefs)
    {
        $pn = $cmmRef.ParentNode.ParentNode.PartNumber
        $cmm = $cmmRef.InnerText
        $objsec = [pscustomobject][ordered]@{PN=$pn;CMM=$cmm;DMC=$dm.Name;}
        $allParts += $objsec
        if($allCMMs -notcontains $cmm)
        {
            $allCMMs += $cmm
        }
        $objsec = $null        
     }    
}
$prop1 = @{Expression='CMM'; Ascending=$true }
$prop2 = @{Expression='PN'; Ascending=$true }

$allParts.Count
#$allParts = $allParts | Sort-Object -Property RefValue -Unique
$allParts.Count
$allCMMs.Count
$allParts.GetEnumerator() | Sort-Object -Property $prop1, $prop2 | Export-Csv "$reportpath\$rptBaseName $rptSuffix1.csv" -NoTypeInformation -Encoding UTF8
$allParts.GetEnumerator() | Sort-Object -Property $prop2, $prop1 | Export-Csv "$reportpath\$rptBaseName $rptSuffix2.csv" -NoTypeInformation -Encoding UTF8
"$reportpath\$rptBaseName $rptSuffix1.csv"
"$reportpath\$rptBaseName $rptSuffix2.csv"
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"