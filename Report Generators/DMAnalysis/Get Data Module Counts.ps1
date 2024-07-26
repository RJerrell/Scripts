$environ = "Production"
$basePath = "f:\KC46 Staging\$environ\Manuals"
[string[]] $PubList = @("KC46", "AMM","ARD", "FIM", "IPB", "LOAPS", "MOM", "NDI", "NDT", "SPCC",  "SSM", "SWPM", "WDM", "WUC")

foreach( $pub in $PubList)
{
    $filesDM = Get-ChildItem -Path "$basePath\$pub\S1000D\SDLLIVE" 
    $filesIll = Get-ChildItem -Path "$basePath\$pub\Illustrations\Illustrations"
    
    "$pub`tData Modules`t" + $filesDM.Count
    "$pub`tGraphics`t" + $filesIll.Count
}