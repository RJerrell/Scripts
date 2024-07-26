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
    $inputPath = "C:\KC46 Staging\production\Manuals\$pub\S1000D\SDLLIVE\DMC*.xml"
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
                break
            }
        }
        $list = ""
        $n  = 5
        $targetDmRef = [xml]@"
                    <dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="1" subSubSystemCode="0" assyCode="2300" disassyCode="04" disassyCodeVariant="A0A" infoCode="921" infoCodeVariant="A" itemLocationCode="A"/>
"@
        foreach($string in $stringToLookup)
        {
            $re = "(.{0,$n})(" + [Regex]::Escape($string) + ")(.{0,$n})"
            $result = (Get-Content $file.FullName) -match $re
            if ($result.Length -gt 0)
            {
                # Does the file have a reference to

                # Load the file into an XmlDocument
                $dm = New-Object System.Xml.XmlDocument
                $dm.Load($file.FullName)
                $dmRefs = $dm.SelectNodes("//dmCode")
                $documentHasCorrectDMRef = $false
                foreach ($dmRef in $dmRefs)
                {
                cls
                   $currentDmRef = [xml]@"
<dmCode modelIdentCode="" systemDiffCode="" systemCode="" subSystemCode="" subSubSystemCode="" assyCode="" disassyCode="" disassyCodeVariant="" infoCode="" infoCodeVariant="" itemLocationCode=""/>
"@
                    foreach ($att in $dmRef.Attributes)
                    {
                        $currentDmRef.dmCode.modelIdentCode = $dmRef.modelIdentCode
                        $currentDmRef.dmCode.systemDiffCode = $dmRef.systemDiffCode
                        $currentDmRef.dmCode.systemCode = $dmRef.systemCode
                        $currentDmRef.dmCode.subSystemCode = $dmRef.subSystemCode
                        $currentDmRef.dmCode.subSubSystemCode = $dmRef.subSubSystemCode
                        $currentDmRef.dmCode.assyCode = $dmRef.assyCode
                        $currentDmRef.dmCode.disassyCode = $dmRef.disassyCode
                        $currentDmRef.dmCode.disassyCodeVariant = $dmRef.disassyCodeVariant
                        $currentDmRef.dmCode.infoCode = $dmRef.infoCode
                        $currentDmRef.dmCode.infoCodeVariant = $dmRef.infoCodeVariant
                        $currentDmRef.dmCode.itemLocationCode = $dmRef.itemLocationCode
                    }
                    $same = Compare-Object -ReferenceObject $currentDmRef.InnerXml -DifferenceObject $targetDmRef.InnerXml
                    $currentDmRef.InnerXml
                    $targetDmRef.InnerXml

                    $same -eq $null

                    if($same -eq $null)
                    {
                        $documentHasCorrectDMRef = $true
                        break
                    }
                }
                if($documentHasCorrectDMRef -eq $false)
                {
                    for ($i = 0; $i -lt $result.Length; $i++)
                    { 
                        $list += $result[$i] + "|"
                    }
                    $save = $true
                }
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
$objs.GetEnumerator() | Sort-Object -Property $prop1 | Export-Csv "$outputPath\$reportName - Lockwires with a DMRef Sorted by DMC.csv" -NoTypeInformation 
$objs.GetEnumerator() | Sort-Object -Property $prop2 | Export-Csv "$outputPath\$reportName - Lockwires with a DMRef Sorted by TechName.csv" -NoTypeInformation