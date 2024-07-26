$total = 0
$totalDMPerBook = @{}
$totalIllustrationsPerBook = @{}
$basePath = "C:\KC46 Staging\Production\Manuals"

[string[]] $PubList   = @("KC46", "ACS", "ARD", "AMM", "FIM", "IPB", "LOAPS",    "NDT", "SIMR", "SRM", "SSM", "SWPM", "WUC", "WDM")
foreach( $Pub in $PubList)
{
    $dmFiles = gci -Path "$basePath\$Pub\S1000D\S1000D\*MC*.xML"
    
    $totalDMPerBook.Add($Pub , $dmFiles.Length)
    $grFiles = gci -Path "$basePath\$Pub\Illustrations\Illustrations\*.*"
    $totalIllustrationsPerBook.Add($Pub , $grFiles.Length)
}