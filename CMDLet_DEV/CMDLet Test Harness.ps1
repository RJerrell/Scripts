
<#

Title: 
Author: Roger Jerrell
Date Created: Get-Date 
Purpose: 
Description of Operation: 
Description of Use:

#>

$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
Import-Module "C:\Junk\CMDLET\S1000D401\S1000D401\bin\Debug\S1000D401.dll" -Force -Verbose
$dmClass = $null
$dmClass = New-Object Boeing.BDS.S1000D401.S1000D401DM

# *****************************************************************************************************
$dm = $null

Function Get-DataModule
{
    Param([string] $path)
    
    $dm = $dmClass.GetDM($path)
}

$publist = @("ACS", "ARD", "AMM", "FIM", "IPC", "LOAPS", "NDT", "SIMR", "SPCC", "SRM","SSM", "WDM", "WUC")
$publist = @("ACS", "ARD", "AMM")

Get-DataModule -path "C:\KC46 Staging\Production\Manuals\IPB\S1000D\SDLLIVE\DMC-1KC46-A-11-21-1600-0400P-941A-A_001-00_SX-US.XML"

$ids     = $dmClass.Get_IdentAndStatusSection()
$address = $dmClass.Get_Address()
$content = $dmClass.Get_Content()
$cType   = $dmClass.ContentType("abc")
$dmc     = $dmClass.Get_DMCode()

"Done"

# *****************************************************************************************************
