$path = "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DML\*.xml"
$publist =@("COMMON" , "AMM-KC", "BCLM-KC", "FIM-KC", "NDT", "SRM", "SSM-KC", "WDM-KC", "TC-KC")

foreach ($pub in $publist)
{  
    $pubPath = "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DDN-1KC46-AAAZZ-81205-2016-00003\$pub\$pub" +  "_DataModules\SDLLIVE\DML*.XML"
    $dmls = gci -Path $pubPath
    foreach ($dml in $dmls)
    {
        $dmlDoc = [xml](Get-Content -Path  $dml.FullName)
        # $dmEntries = $dmlDoc.SelectNodes("/dml/dmlContent/dmEntry[./@dmEntryType!='c' and ./@dmEntryType!='n' and ./@dmEntryType!='d']")
        $dmEntriesN = $dmlDoc.SelectNodes("/dml/dmlContent/dmEntry[./@dmEntryType='n']")
        $dmEntriesC = $dmlDoc.SelectNodes("/dml/dmlContent/dmEntry[./@dmEntryType='c']")
        $dmEntriesD = $dmlDoc.SelectNodes("/dml/dmlContent/dmEntry[./@dmEntryType='d']")

        $dml.FullName
        "$pub`t Total NEW DM Entries: `t"   + $dmEntriesN.Count
        "$pub`t Total CHG DM Entries: `t"   + $dmEntriesC.Count
        "$pub`t Total DEL DM Entries: `t"   + $dmEntriesD.Count
        "$pub`t Total Current CSDB DM Count: `t"   + $dmEntriesD.Count
    }
}