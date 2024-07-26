cls
<#
    
    
    ADD CODE TO GET ALL THE PMC files in the KC46 CONFIGURATION MANUALS


#>

$commonRoot = "KC46"

[string[]] $PubList   = @( "ACS", "ARD", "AMM", "FIM", "IPB", "LOAPS", "NDT", "SIMR", "SRM", "SSM", "SWPM", "WUC", "WDM")    
$PubHash   = @{}
$pathRoot = "c:\kc46 staging\production\manuals"
$pmcList = @()
foreach ($pub in $PubList)
{
    $fullpath = "$pathRoot\$pub\s1000d\s1000d\PMC*.XML"
    $pmc = gci -Path $fullpath
    #$pmcName = $pmcList.Add( $pmc[0].FullName)
    $pmcList += $pmc[0].FullName
    $PubHash.Add($pub,$pmc[0].FullName)
}
$pmcList

$xsltPath = "C:\KC46 Staging\scripts\Report Generators\PageBlock to DMC Report - All Manuals-CSV.xsl"
$xslt = New-Object System.Xml.Xsl.XslCompiledTransform
$settings = $xslt.OutputSettings
$settings.OutputMethod
$xslt.Load($xsltPath)
$counter = 0
$global:startTime = (Get-Date -Format yyyy-MM-dd-HH-mm-ss)
$sortedHash = $PubHash.GetEnumerator() |Sort-Object -Property Name

foreach ($pub in $sortedHash)
{
    $doc = New-Object -TypeName  System.Xml.XmlTextReader -ArgumentList $sortedHash[$counter].Value
    $acronym = $sortedHash[$counter].Key
    $outputPath = "c:\temp\$startTime - $acronym.csv"
    $XmlWriter = New-Object System.XMl.XmlTextWriter($outputPath,$null)
    $xslt.Transform($doc , $XmlWriter)
    $counter ++
}