<#

Title: 
Author: Roger Jerrell
Date Created: Get-Date 
Purpose: 
Description of Operation: 
Description of Use:

#>
cls
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

[string[]] $PubList   = @("ABDR","ASIP", "ACS","LOAPS","SIMR","SPCC")
[string[]] $PubList   = @("ABDR")
$basePath = "C:\KC46 Staging\Production\Manuals"
$strToFind = "?>"

$strToAdd  = @"
<!DOCTYPE dmodule [
<!NOTATION CGM SYSTEM "image/cgm">
<!NOTATION PNG SYSTEM "image/png">
]>
"@
foreach ($Pub in $PubList)
{
    $path1 = "$basePath\$Pub\S1000D"
    $path2 = "$basePath\$Pub\S1000D\SDLLIVE"
    $Pub
    $files1 = gci -Path $path1 -Filter DMC*.xml
    $files2 = gci -Path $path2 -Filter DMC*.xml
    foreach ($file in $files1)
    {
        $sr = New-Object System.IO.StreamReader($file.FullName)
        $cString = $sr.ReadToEnd()
        $sr.Close()
        $sr.Dispose()

        if($cString -notcontains "<!DOCTYPE dmodule")
        {
            $idx = $cString.IndexOf($strToFind) + 2
            $idx2 = $cString.IndexOf('<dmodule')
            $cString = $cString.Remove($idx, ($idx2-1) -$idx)
            $cString = $cString.Insert($idx+1,$strToAdd)
        }

        $dmXML = New-Object System.Xml.XmlDocument
        $dmXML.LoadXml($cString)

        Save-PrettyXML -FName $file.FullName -xmlDoc $dmXML
    }
    foreach ($file in $files2)
    {
        $sr = New-Object System.IO.StreamReader($file.FullName)
        $cString = $sr.ReadToEnd()
        $sr.Close()
        $sr.Dispose()

        if($cString -notcontains "<!DOCTYPE dmodule")
        {
            $idx = $cString.IndexOf($strToFind) + 2
            $idx2 = $cString.IndexOf('<dmodule')
            $cString = $cString.Remove($idx, ($idx2-1) -$idx)
            $cString = $cString.Insert($idx+1,$strToAdd)
        }

        $dmXML = New-Object System.Xml.XmlDocument
        $dmXML.LoadXml($cString)

        Save-PrettyXML -FName $file.FullName -xmlDoc $dmXML
    }
}
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"