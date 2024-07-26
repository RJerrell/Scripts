cls
$environment = "Production"
$path = "C:\KC46 Staging\Production\Manuals"
$KC46DataRoot = "C:\KC46 Staging"
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"
[string[]] $PubList   = @(  "AMM" )

foreach($pub in $PubList)
{
    $files = gci -Path "C:\KC46 Staging\Production\Manuals\AMM\S1000D\S1000D\DMC*.xml"
    foreach($file in $files)
    {
        $dmc = [xml] (Get-Content -Path $file.FullName)
        $file.FullName
    }
}