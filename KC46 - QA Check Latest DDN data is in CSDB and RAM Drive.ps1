CLS
$ErrorActionPreference = "Stop"
$error.Clear()
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\KC46 Staging\Scripts\Modules;"
Import-Module -Name "KC46Common" -Verbose -Force					
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
# ********************************************************************
$pathTolatestDDN = "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DDN_1KC46-81205-81205-2018-P0001"

$pathToRamDrive = "F:\KC46 Staging\Production\Manuals"
$pathToCSDB = "C:\KC46 Staging\Production\FullCSDB"

$ddnDMCs = gci -Path $pathTolatestDDN -Recurse -Filter "DMC*.*"
$notFoundList = @()
$notFoundList_RAM = @()
foreach ($ddnDMC in $ddnDMCs)
{
    $book = Get-DocTypeFromDMC -dc  $ddnDMC.Name
    $sName = $ddnDMC.Name
    [bool] $found = Test-Path -Path "$pathToCSDB\$book\S1000D\SDLLIVE\$sName"
    if($found)
    {
       
    }
    else
    { 

        "File Not found in CSDB: " +  $ddnDMC.FullName
        $notFoundList += $ddnDMC.FullName
        Exit 

    }
   

   [bool] $found = Test-Path -Path "$pathToRamDrive\$book\S1000D\SDLLIVE\$sName"
       if($found)
    {
       
    }
    else
    { 
       
        "File Not found in RAM Drive: " +  $ddnDMC.FullName 
        $notFoundList_RAM += $ddnDMC.FullName
        Exit

    }

}
    
    
    
    "CSDB IS OKAY"
    "RAM DRIVE IS OKAY"


# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"					
					