cls

$outputPath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$reportName = "Data Modules Search Report - Support Equipment Listing for Richard Trach"
$stringToLookup = @("762-17831-1",
"G28035-2",
"RMA100001-1",
"5TSFFT1T2849-9",
"C28002-9",
"2MIT143T0390",
"MIT140T0395",
"MIT140T0397",
"MIT143T0390",
"5TSFFT1T2820-2",
"5TSFFT1T2910-4",
"5TSFFT1T2616-1",
"5TSFFT1T2814-4",
"TSFFT1T2621",
"FPME453T1272",
"MIT143T0390-1",
"MIT842-349866",
"MIT842-349210",
"MA280004-1",
"3ME143T0390",
"2ME143T0390",
"ME143T0390",
"OHME143T0390",
"ME842-349214",
"OHME842-349510",
"OHME842-349710",
"MIT842-349521",
"F72849-1",
"ST2580-177",
"ST2580-381A-15",
"ST895A-3",
"ST8709H-X",
"ST8744",
"G26004-65",
"G26004-66",
"F72959-35",
"F72959-33")
$seNumsNotReferenced = @("762-17831-1",
"G28035-2",
"RMA100001-1",
"5TSFFT1T2849-9",
"C28002-9",
"2MIT143T0390",
"MIT140T0395",
"MIT140T0397",
"MIT143T0390",
"5TSFFT1T2820-2",
"5TSFFT1T2910-4",
"5TSFFT1T2616-1",
"5TSFFT1T2814-4",
"TSFFT1T2621",
"FPME453T1272",
"MIT143T0390-1",
"MIT842-349866",
"MIT842-349210",
"MA280004-1",
"3ME143T0390",
"2ME143T0390",
"ME143T0390",
"OHME143T0390",
"ME842-349214",
"OHME842-349510",
"OHME842-349710",
"MIT842-349521",
"F72849-1",
"ST2580-177",
"ST2580-381A-15",
"ST895A-3",
"ST8709H-X",
"ST8744",
"G26004-65",
"G26004-66",
"F72959-35",
"F72959-33")
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
        foreach ($row in $rows)
        {
            if($row.Contains($fn))
            {
                $rArray = $row.Split("|")
                $pbTask = $rArray[1].ToString()
            }
        }

        $sr = New-Object System.IO.StreamReader($file.FullName)
        $c = $sr.ReadToEnd()
        $sr.Dispose()
        $save = $false
        $list = ""
        
        foreach ($string in $stringToLookup)
        {
            if($c -match $string)
            {
                $save = $true
                $list += $string + "|"
                $seNumsNotReferenced.Remove($string)
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
$objs.GetEnumerator() | Sort-Object -Property $prop1 | Export-Csv "$outputPath\$reportName - Sorted by DMC.csv" -NoTypeInformation 
$objs.GetEnumerator() | Sort-Object -Property $prop2 | Export-Csv "$outputPath\$reportName - Sorted by TechName.csv" -NoTypeInformation

$seNumsNotReferenced | %{$_} |fl