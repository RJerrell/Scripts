
CLS
$ErrorActionPreference = "Stop"
$error.Clear()
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\KC46 Staging\Scripts\Modules;"
Import-Module -Name "KC46Common" -Verbose -Force					
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
# ********************************************************************
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$dmXml = new-object -TypeName S1000D.DataModule_401

$array = @()

$files1 = gci -Path "F:\KC46 Staging\Production\Manuals\AMM\S1000D\SDLLIVE" -Filter DMC*.XML
$files2 = gci -Path "F:\KC46 Staging\Production\Manuals\FIM\S1000D\SDLLIVE\" -Filter DMC*.XML
$files3 = gci -Path "F:\KC46 Staging\Production\Manuals\TC\S1000D\SDLLIVE\" -Filter DMC*.XML

$dmc2Find = "DMC-1KC46-A-29-21-0100-02A0A-010A-A"

foreach ($file in $files1)
{
    $dmXml.ParseDM($file.Fullname)
    foreach($dmRef in $dmXml.Refs.dmRef)
    {
        $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
        if($dmc2Find -eq $dmc)
        {
            $dmXml.ParseDM($file.FullName)
            $tn = $dmXml.TechName
            $array += "$tn `t " + $file.Name
        }
    }    
}

foreach ($file in $files2)
{
    $dmXml.ParseDM($file.Fullname)
    foreach($dmRef in $dmXml.Refs.dmRef)
    {
        $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
        if($dmc2Find -eq $dmc)
        {
            $dmXml.ParseDM($file.FullName)
            $tn = $dmXml.TechName
            $array += "$tn `t " + $file.Name
        }
    }    
}

foreach ($file in $files3)
{
    $dmXml.ParseDM($file.Fullname)
    foreach($dmRef in $dmXml.Refs.dmRef)
    {
        $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
        if($dmc2Find -eq $dmc)
        {
            $dmXml.ParseDM($file.FullName)
            $tn = $dmXml.TechName
            $array += "$tn `t " + $file.Name
        }
    }    
}

# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"					
					