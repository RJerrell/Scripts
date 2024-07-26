CLS
$ErrorActionPreference = "Stop"
$error.Clear()
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'

Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "KC46CSDBManager" -Verbose -Force
Import-Module -Name "KC46Augmentor" -Verbose -Force
Import-Module -Name "KC46S1000DRules" -Verbose -Force
# ********************************************************************
$pathToRAMDriveCSDB = "F:\KC46 Staging\Production\Manuals"
[string[]] $PubList   = @("KC46","ABDR","ACS","AMM","ARD","ASIP","FIM","IPB","LOAPS","NDT","SIMR","SPCC","TC","WUC","SSM","SWPM") | Sort-Object
foreach ($Pub in $PubList)
{
    $pathToBook = "$pathToRAMDriveCSDB\$Pub\S1000D\SDLLIVE"
    Update-1Folder -basepath $pathToBook
}

$pathToCSDB_ON_C_DRIVE = $pathToRAMDriveCSDB.Replace("F:", "C:")
foreach ($Pub in $PubList)
{
    $pathToBook = "$pathToRAMDriveCSDB\$Pub\S1000D\SDLLIVE"
    $files = gci -Path $pathToBook -Filter DMC*.xml
    
    foreach ($file in $files)
    {
        Copy-Item -Path $file.FullName -Destination "$pathToCSDB_ON_C_DRIVE\$Pub\\S1000D\SDLLIVE\" -ErrorAction Stop -Force
    }
}

# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"					
					