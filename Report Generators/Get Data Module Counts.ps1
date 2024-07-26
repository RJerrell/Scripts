$environ = "Production"
$basePath = "\\kc46-lc-sdl\c$\KC46 Staging\$environ\Manuals"

[string[]] $PubList = @("KC46", "AMM","ARD", "FIM", "IPB", "LOAPS", "NDT", "SSM", "SWPM", "WDM", "WUC")

foreach( $pub in $PubList)
{
    $filesDM = Get-ChildItem -Path "$basePath\$pub\S1000D\S1000D" 
    $filesIll = Get-ChildItem -Path "$basePath\$pub\Illustrations\Illustrations"
    
    "$pub`tData Modules`t" + $filesDM.Count
    "$pub`tGraphics`t" + $filesIll.Count
}