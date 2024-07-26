cls
$sd = Get-Date
$ErrorActionPreference = "Stop"
$error.Clear()

# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$parserDML = new-object -TypeName S1000D.DataModuleList
$parserDM = new-object -TypeName S1000D.DataModule_401
$parserPM = new-object -TypeName S1000D.PublicationModule_401
$new = @()
$deletes = @()
$changes = @()
$pmDMListing = @()
for ($i = 10; $i -lt 12; $i++)
{ 
    $dmls = gci -Path "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set\Release $i DMLs" -Filter DML-1KC46-AAA0A-P*.* | Sort-Object
    foreach ($dml in $dmls)
    {
        $parserDML.ParsePM($dml.FullName)
        $e = $parserDML.DMEntries
        $lookupPath = "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DDN-1KC46-AAAZZ-81205-2018-00002\AMM-KC\AMM-KC_DataModules\SDLLIVE"
        $csdbPath = ""
        foreach ($entry in $e)
        {          
            if($entry.dmEntryType -eq "n")
            {
                $fname = Get-FilenameFromDMRef -dmRef $entry.dmRef -filePref "DMC"
                $rfs = gci -Path "$lookupPath\$fname`*" -File
                $parserDM.ParseDM($rfs[0].FullName)
                $parserDM.IssueType
                $y = $parserDM.IdentAndStatusSection.dmAddress.dmAddressItems.issueDate.year
                $m = $parserDM.IdentAndStatusSection.dmAddress.dmAddressItems.issueDate.month
                $d = $parserDM.IdentAndStatusSection.dmAddress.dmAddressItems.issueDate.day
                $tn = $parserDM.TechName
                $in = $parserDM.InfoName
                $dt = $y.ToString() + $m.ToString() + $d.ToString()
                if($parserDM.IssueType -eq "new")
                {              
                    if(!($new.Contains($fname)))
                    {
                        #$new += $fname + "`t|`t" + $dml.Name
                        $new += New-Object -TypeName PSObject -Property @{
                                FName = $fname;
                                TechName = $tn;
                                InfoName = $in;
                                IssueDate = $dt;
                                DML = $dml.Name
                                
                        } | Select FName,TechName,InfoName,IssueDate,DML
                    }     
                }           
            }
        }        
    }
}

#$deletes | Sort-Object -Descending

Remove-Item "C:\KC46 Staging\Scripts\Report Generators\Outputs\REPORT.XLSX" -ErrorAction SilentlyContinue
$prop1 = @{Expression='FName'; Ascending=$true }
$new   |  Sort-Object -Property $prop1 | Export-XLSX -Path "C:\KC46 Staging\Scripts\Report Generators\Outputs\REPORT.XLSX" -Header FName,TechName,InfoName,IssueDate,DML -WorksheetName "New Data Modules"

$parserPM.ParsePM("F:\KC46 Staging\Production\Manuals\AMM\S1000D\SDLLIVE\PMC-1KC46-81205-A0000-00_009-00_SX-US.xml")
$dmRefs = $parserPM.DmRefs
foreach($dmr in $dmRefs)
{
    $dmc = Get-FilenameFromDMRef -dmRef $dmr -filePref "DMC"
    if(!($pmDMListing.Contains($dmc)))
    {
        $pmDMListing += $dmc
    }
}

foreach ($item in $new)
{
    if($pmDMListing.Contains( $item.Fname ))
    {
        #"yes"
    }
    else
    {
        $item.Fname
    }
}
