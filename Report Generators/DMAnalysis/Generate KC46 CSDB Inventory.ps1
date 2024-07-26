cls

$icnArray = @()
$dmcArray = @()
[string[]] $PubList   = @("KC46", "ABDR", "ACS", "AMM", "ARD", "ASIP", "FIM", "IPB", "LOAPS", "NDT", "SIMR",  "SSM", "SRM", "SWPM", "TC", "WUC", "WDM")

foreach ($Pub in $PubList)
{
    $icns = gci -Path "C:\KC46 Staging\Production\Manuals\$pub\ILLUSTRATIONS\ILLUSTRATIONS\*.cgm"
    $dmcs = gci -Path "C:\KC46 Staging\Production\Manuals\$pub\S1000D\SDLLIVE\*.XML"

    foreach ($icn in $icns)
    {
        if($icnArray.Contains($icn.Name))
        {
            $x = "asdf"
        }
        else
        {
            $arrayEntry = "$pub," + $icn.Name
            $icnArray += $arrayEntry
        }
    }
        
    foreach ($dmc in $dmcs)
    {
        if($dmcArray.Contains($dmc.Name))
        {
            $x = "asdf"
        }
        else
        {
            $arrayEntry = "$pub," + $dmc.Name
            $dmcArray += $arrayEntry
        }
    }
}

$icnArray | Out-File "C:\KC46 Staging\Scripts\Report Generators\Outputs\ICN Inventory.csv"
$dmcArray | Out-File "C:\KC46 Staging\Scripts\Report Generators\Outputs\DMC Inventory.csv"