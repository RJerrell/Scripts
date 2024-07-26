
CLS
$ErrorActionPreference = "Stop"
$error.Clear()
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "KC46Augmentor" -Verbose -Force					

# Add your code here
$BasePathTmp = $BasePath.Replace("F:", "C:")

# BDS Authored manuals
[string[]] $PubList1   = @("KC46","ABDR","ACS","ASIP","LOAPS","NDTS","SIMR","SPCC","WUC") | Sort-Object

# CAS and CDG Manuals
[string[]] $PubList2   = @("AMM","ARD","FIM","IPB","NDT","TC","SSM","SWPM", "WDM")

foreach($pub in $PubList1)
{
    Update-1Folder -basepath "C:\KC46 Staging\production\Manuals\$pub\S1000D\SDLLIVE" -forceAugmentation $true -verbose $true
}

foreach($pub in $PubList2)
{
    Update-1Folder -basepath "C:\KC46 Staging\production\Manuals\$pub\S1000D\SDLLIVE" -forceAugmentation $true -verbose $true
}

# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"					
					