cls
<#
Title: 
Author: Roger Jerrell
Date Created: 
Purpose: 
Description of Operation: 
Description of Use: 
#>

$sd = Get-Date
$ErrorActionPreference = "SilentlyContinue"
$error.Clear()
$a = (Get-Host).PrivateData
$a.WarningBackgroundColor = "Yellow"
$a.WarningForegroundColor = "Black"
$rn = "<br/>"
$a.ErrorBackgroundColor = "red"
$a.ErrorForegroundColor = "White"

$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
$env:PSModulePath

#Get-Module -ListAvailable
Import-Module -Name "KC46Common" -Verbose -Force
# *****************************************************************************************************
$outputPath = "F:\KC46 Staging\scripts\Report Generators\Outputs"
$reportName = "Data Modules Search Report"
$stringToLookup = @("AIPC")
$objs =  @()

$Publist = @("AMM")
$pbTaskXML = New-Object System.Xml.XmlDocument
$pbTaskXMLpath = "F:\KC46 Staging\Scripts\Report Generators\Outputs\PBTask to DMC Map.xml"
$pbTaskXML.Load($pbTaskXMLpath)
function Get-CHSESUFromIPCRefArray
{
    Param([string] $IPCRefString)
    return $IPCRefString.Split("-")
}

foreach ($pub in $Publist)
{
    #Limited to Installation Procedures
    $inputPath = "F:\KC46 Staging\production\Manuals\$pub\S1000D\SDLLIVE\DMC*.xml"
    $files = gci -Path  $inputPath -Verbose
    $pub
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
        #$file.Name
        foreach($string in $stringToLookup)
        {
            # Test the file to see if contains an AIPC reference
            $sr = New-Object System.IO.StreamReader($file.Fullname)
            $c = $sr.ReadToEnd()
            $sr.Close()
            $sr.Dispose()
            
            $matches.Clear()
            
            $rex = "<externalPubTitle>"

            $result = $c -match $rex
            
            # $matches
            if($result)
            {
                $xmlDoc = New-object System.Xml.XmlDocument
                $xmlDoc.Load($file.Fullname)
                $extpubRefs = $xmlDoc.SelectNodes("//externalPubTitle")

                # Joseph
                foreach($match in $matches)
                {
                    $mArray = $match[0].Split(" ")
                    $mCHSESU = $mArray[1].Split("-")

                    # Determine the CH-SE-SU of the current dm based on the file name
                    $fileParts =  Get-FileNameArray $file.FullName
                    $ch = $fileParts[3]
                    if($ch.Contains("N"))
                    {
                        $ch = $ch -replace 'N'
                    }
                    $se = $fileParts[4]
                    $su = $fileParts[5].Substring(0,2)
                    #$su
                    $mch = $mCHSESU[0].Trim()
                    $mse = $mCHSESU[1].Trim()
                    $msu = $mCHSESU[2].Trim()
                    if(($ch -ne $mch) -or ($se -ne $mse) -or ($su -ne $msu) )
                    {
                        $val = $match.Values[0]
                        $list += $match.Values[0] + "|"
                        $val
                    } 
                }
            }
            else
            {
                $list += "No IPC Reference|"
            }
        }

        if($list -ne "")
        {        
            $list = $list.Substring(0, $list.Length -1)
            $save = $true
        }

        if($save)
        {
            $FN = $file.Name
            $DMC = New-Object System.Xml.XmlDocument
            $DMC.Load($file.FullName)

            $techName = $DMC.DocumentElement.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
            $infoName = $DMC.DocumentElement.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName

            $obj = [pscustomobject][ordered]@{Manual=$pub;DMC=$FN;TechName=$techName;InfoName=$infoName;Contains=$list;PBTask=$pbTask;}
            $objs += $obj
            $obj=$null
        }
    }
}

# Define the sort properties for the report
$prop1 = @{Expression='DMC'; Ascending=$true }
$prop2 = @{Expression='TechName'; Ascending=$true }

# Store the report
$objs.GetEnumerator() | Sort-Object -Property $prop1 | Export-Csv "$outputPath\$reportName - All Manuals to IPC Correlation Report Sorted by DMC.csv" -NoTypeInformation 
$objs.GetEnumerator() | Sort-Object -Property $prop2 | Export-Csv "$outputPath\$reportName - All Manuals to IPC Correlation Report Sorted by TechName.csv" -NoTypeInformation

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"