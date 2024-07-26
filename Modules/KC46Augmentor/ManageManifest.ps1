CLS
$ErrorActionPreference = "Stop"
$error.Clear()
#$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\KC46 Staging\Scripts\Modules;"
#Import-Module -Name "KC46Common"  -Force					
#Import-Module -Name "KC46Augmentor"  -Force

# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
# ********************************************************************

$manifestPath = "C:\KC46 Staging\Scripts\Modules\KC46Augmentor\KC46Augmentor.psd1"

if(!(Test-Path -Path $manifestPath))
{
    New-ModuleManifest -Path $manifestPath  -ModuleVersion "10.0" -Author "Roger A. Jerrell 1641883"
}

Test-ModuleManifest -Path $manifestPath

# ********************************************************************

$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"