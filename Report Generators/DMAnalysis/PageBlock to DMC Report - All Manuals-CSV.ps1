#Path to PMCS
$basePath = "C:\KC46 Staging\Production\Manuals"
[string[]] $PubList = @( "AMM", "ARD","FIM","NDT", "SSM", "WDM" )



$report = @()

foreach($pub in $PubList)
{
    $dms = Get-ChildItem -Path "$basePath\$pub" -Filter DMC*.XML -Recurse |Sort-Object -Property Name
    foreach($dm IN $dms)
    {
        $dmc = $dm.Name
        $d = $dmc.Remove($dmc.Length - 4, 4)
       #$d
        $dm.Name
        $xml = [xml] (Get-Content -Path $dm.FullName)
        $ctype = "unknown"
        $cNode = $xml.dmodule.content
        $dNodes = $cNode.SelectNodes("//description").Count
        $pNodes = $cNode.SelectNodes("//procedure").Count
        
        if($pNodes -gt 0)
        {
            $ctype = "procedure"
        }

        if($dNodes -gt 0)
        {
            $ctype = "description"
        }

        $tName  = $xml.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
        $iName = $xml.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName
        $report += "$d`t$cType`t$tName`t$iName"
    }

}
$report