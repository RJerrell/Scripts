cls

$outputPath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$reportName = "Data Modules Search Report"
$stringToLookup = @("lockwire")
$objs =  @()

$Publist = @("AMM")
$pbTaskXML = New-Object System.Xml.XmlDocument
$pbTaskXMLpath = "C:\KC46 Staging\Scripts\Report Generators\Outputs\PBTask to DMC Map.xml"
$pbTaskXML.Load($pbTaskXMLpath)

foreach ($pub in $Publist)
{
    #Limited to Installation Procedures
    $inputPath = "C:\KC46 Staging\production\Manuals\$pub\S1000D\SDLLIVE\dmc-*"
    $files = gci -Path  $inputPath -Verbose
    foreach ($file in $files)
    {
        $fn = $file.Name
        $rows = $pbTaskXML.Objs.S
        $pbTask = ""
        $save = $false
        foreach ($row in $rows)
        {
            if($row.Contains($fn))
            {
                $rArray = $row.Split("|")
                $pbTask = $rArray[1].ToString()
            }
        }
        $list = ""
        $n  = 5
        foreach($string in $stringToLookup)
        {
            $re = "(.{0,$n})(" + [Regex]::Escape($string) + ")(.{0,$n})"
            $result = (Get-Content $file.FullName) -match $re
            if ($result.Length -gt 0)
            {
                for ($i = 0; $i -lt $result.Length; $i++)
                { 
                    $list += $result[$i] + "|"
                }
                $save = $true
            }
        }
        if($list -ne "")
        {        
            $list = $list.Substring(0, $list.Length -1)
        }
        if($save)
        {
            $FN = $file.Name
            $DMC = [xml] (Get-Content -Path $file -Raw)
            $techName = $DMC.DocumentElement.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
            $infoName = $DMC.DocumentElement.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName
            $obj = [pscustomobject][ordered]@{DMC=$FN;TechName=$techName;InfoName=$infoName;Contains=$list;PBTask=$pbTask;}
            $objs += $obj
            $obj=$null
        }
    }
}
# Define the sort properties for the report
$prop1 = @{Expression='DMC'; Ascending=$true }
$prop2 = @{Expression='TechName'; Ascending=$true }


# Store the report
$objs.GetEnumerator() | Sort-Object -Property $prop1 | Export-Csv "$outputPath\$reportName - Airline to Operator Sorted by DMC.csv" -NoTypeInformation 
$objs.GetEnumerator() | Sort-Object -Property $prop2 | Export-Csv "$outputPath\$reportName - Airline to Operator Sorted by TechName.csv" -NoTypeInformation