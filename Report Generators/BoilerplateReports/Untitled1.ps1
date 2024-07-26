$root = "C:\S1000D Publishing Workspace\customers\MITAC\sets\Maintenance\data\WUC"
$files = gci -Path $root -Filter DMC*.XML
$good = @()
$list = @()
foreach ($file in $files)
{
    $bps = $file.Name.Split("_")
    $bn = $bps[0]
    $gf = gci -Path $root -Filter $bn`*.xml |Sort-Object -Descending | Select-Object -First 1
    if(! $good.Contains($bn))
    {
        $good += $bn
        $list += $gf.Name
    }

}
cls
$list | Sort-Object | %{$_}
