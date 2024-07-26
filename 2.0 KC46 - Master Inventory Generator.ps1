cls
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
Import-Module -Name "PSExcel" -Verbose -Force
$parserDM = new-object -TypeName S1000D.DataModule_401
$parserPM = new-object -TypeName S1000D.PublicationModule_401
$parserCommon = New-Object -TypeName S1000D.CommonFunctions
# **************************************************************************

# Objective: Determine if any data module has ever had the Techname or InfoName changed since its 
# initial issuance.

$uFolder = "C:\KC46 Staging\Production\Archives\Source\UnpackHere"
$ddns = Get-ChildItem -Path $uFolder -Directory -Name -Filter DDN-1KC46-AAAZZ* |Sort-Object


#[string[]] $ddns   = @("DDN-1KC46-AAAZZ-81205-2016-00001","DDN-1KC46-AAAZZ-81205-2016-00003","DDN-1KC46-AAAZZ-81205-2016-00004 - Revised","DDN-1KC46-AAAZZ-81205-2016-00005","DDN-1KC46-AAAZZ-81205-2017-00001","DDN-1KC46-AAAZZ-81205-2017-00002","DDN-1KC46-AAAZZ-81205-2017-00003","DDN-1KC46-AAAZZ-81205-2018-00001", "DDN-1KC46-AAAZZ-81205-2018-00002") | Sort-Object

$masterlist = @()

$reportFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs\BoilerPlateReports\MasterInventory"

$reportName = "KC46 Tanker CAS MasterInventory.xlsx"
if(!(Test-Path -Path $reportFolder ))
{
    mkdir $reportFolder
}
else
{
    Remove-Item -Path $reportFolder -Force -Recurse    
    mkdir $reportFolder
}
foreach ($ddn in $ddns)
{
    $files = Get-ChildItem -Path "$uFolder\$ddn" -Recurse -Filter *MC* | ?{$_.DirectoryName.ToUpper().Contains("SDLLIVE")}
    foreach ($file in $files)
    {
        $parts = $file.FullName.Split("\")
        $masterList += New-Object -TypeName PSObject -Property @{
            DMC = $file.Name;
            DocType= $parts[7];
            DDN= $ddn;                
        } | Select-Object DMC,DocType,DDN

    }


}
$prop1 = @{Expression='DMC'; Ascending=$true }
$prop2 = @{Expression='DocType'; Ascending=$true }
$prop3 = @{Expression='DDN'; Ascending=$true }

Remove-Item -Path "$reportFolder\$reportName" -Force -ErrorAction SilentlyContinue
$masterList  | Sort-Object -Property $prop1, $prop2,$prop3  | Export-XLSX -Path "$reportFolder\$reportName" -Header  DMC,DocType,DDN -WorksheetName "TankerSource"


# **************************************************************************
