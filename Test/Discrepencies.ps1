cls
$sd = Get-Date
$ErrorActionPreference = "Stop"
$error.Clear()

# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'

Import-Module -Name "KC46Common" -Verbose -Force
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose

# *****************************************************************************************************
$parserDMC = new-object -TypeName S1000D.DataModule_401
$parserCOMM = New-Object -TypeName S1000D.CommonFunctions
$pathToData = "C:\KC46 Staging\Production\Manuals\AMM\S1000D\SDLLIVE"
[string[]] $dmcs = @("DMC-1KC46-A-06-24-0000-02A0A-911A-A","DMC-1KC46-A-06-30-0000-10A0A-900A-A","DMC-1KC46-A-06-30-0000-11A0A-900A-A","DMC-1KC46-A-20-10-0900-07A0A-520A-A","DMC-1KC46-A-20-10-2700-02A0A-720A-A","DMC-1KC46-A-21-21-0000-02A0A-042A-A","DMC-1KC46-A-21-21-1000-02A0A-042A-A","DMC-1KC46-A-21-33-0000-01A0A-042A-A","DMC-1KC46-A-22-12-0000-01A0A-042A-A","DMC-1KC46-A-23-25-3100-02A0A-520A-A","DMC-1KC46-A-25-24-1600-06A0A-520A-A","DMC-1KC46-A-25-25-0300-03A0A-280A-A","DMC-1KC46-A-25-31-0000-05A0A-300A-A","DMC-1KC46-A-25-52-1000-03A0A-685A-A","DMC-1KC46-A-25-52-1000-04A0A-685A-A","DMC-1KC46-A-25-61-0300-03A0A-720A-A","DMC-1KC46-A-25-61-0300-04A0A-720A-A","DMC-1KC46-A-25-65-1500-04A0A-920A-A","DMC-1KC46-A-25-65-1500-06A0A-920A-A","DMC-1KC46-A-25-91-0500-02A0A-310A-A","DMC-1KC46-A-27-00-0100-03A0A-300A-A","DMC-1KC46-A-28-11-0000-12A0A-685A-A","DMC-1KC46-A-28-25-0000-02A0A-042A-A","DMC-1KC46-A-32-11-1300-03A0A-280A-A","DMC-1KC46-A-32-21-1100-04A0A-361A-A","DMC-1KC46-A-35-31-0100-05A0A-300A-A","DMC-1KC46-A-47-32-0000-01A0A-042A-A","DMC-1KC46-A-49-11-0000-06A0A-020A-A","DMC-1KC46-A-51-00-6100-03A0A-280A-A","DMC-1KC46-A-51-21-0400-03A0A-685A-A","DMC-1KC46-A-51-31-0100-03A0A-258A-A","DMC-1KC46-A-52-09-0200-06A0A-520A-A","DMC-1KC46-A-52-48-1700-04A0A-020A-A","DMC-1KC46-A-52-48-1700-05A0A-020A-A","DMC-1KC46-A-56-21-0000-01A0A-042A-A","DMC-1KC46-A-N70-24-0500-03A0A-720A-A","DMC-1KC46-A-N71-11-0000-03A0A-270A-A","DMC-1KC46-A-00-00-0000-03A0F-430A-A")
$records = @()
foreach ($dmc in $dmcs)
{
    $files = gci -Path $pathToData -Filter "$dmc`*"
    foreach($file in $files)
    {
        $parserDMC.ParseDM($file.FullName)

        $records += New-Object -TypeName PSObject -Property @{
                    Book       = "AMM";
                    DMC = $file.Name;
                    Type     = $parserDMC.DmType;                                      
                    EMODBuild   = $parserDMC.EmodBuild.ToString().Replace("EMOD BUILD:", "").Trim() ;
                    IssueNumber = $parserDMC.IssueInfo.issueNumber.ToString();
                    IssueYear    = $parserDMC.Issue_year.'#text';
                    IssueMonth    = $parserDMC.Issue_month.'#text';
                    IssueDay    = $parserDMC.Issue_day.'#text';
                } | Select Book,DMC,Type,EMODBuild,IssueNumber,IssueYear,IssueMonth,IssueDay
    }
}
$prop1 = @{Expression='Book'; Ascending=$true }
$prop2 = @{Expression='DMC'; Ascending=$true }
$prop2 = @{Expression='IssueNumber'; Ascending=$true }

$outputPath = "C:\KC46 Staging\Scripts\Report Generators\Discrepencies\BDS and CAS Discrepencies - $startTime.csv"
$records.GetEnumerator() | Sort-Object -Property $prop1, $prop2 | Export-Csv $outputPath -NoTypeInformation
$outputPath
# *****************************************************************************************************

$ed = Get-Date

$x = $ed.Subtract($sd)
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds

"Process completed"