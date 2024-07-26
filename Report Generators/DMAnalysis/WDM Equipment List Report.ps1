cls
$outputPath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$pm = New-Object System.Xml.XmlDocument
$pm.Load("C:\KC46 Staging\Production\Manuals\WDM\S1000D\SDLLIVE\PMC-1KC46-81205-W0000-00.xml")
$objs = @()
# /pm/content/pmEntry/pmEntry//dmRef
$eList = $pm.SelectSingleNode("/pm/content/pmEntry/pmEntry[./pmEntryTitle='91-00 - EQUIPMENT LIST']")
foreach($ref in $elist.dmRef)
{
    if($ref.title.StartsWith("D"))
    {
        $obj = [pscustomobject][ordered]@{DMTitle=$ref.title;DMC=$ref.dmRefIdent.'#comment'.Replace(": ","-")}
        $objs += $obj
        $obj=$null

    }
}
$prop1 = @{Expression ='DMTitle'; Ascending=$true }

$objs.GetEnumerator() | Sort-Object -Property $prop1 | Export-Csv "$outputPath\WDM EquipmentList.csv" -NoTypeInformation