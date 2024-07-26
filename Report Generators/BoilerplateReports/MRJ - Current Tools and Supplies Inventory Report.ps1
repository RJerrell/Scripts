CLS
$ErrorActionPreference = "Stop"
$error.Clear()
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\KC46 Staging\Scripts56\Modules;"
Import-Module -Name "KC46Common" -Verbose -Force					
# Include all common variables
. 'C:\KC46 Staging\Scripts56\Common\KC46CommonVariables.ps1'

Add-Type -Path "d:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$parserDM = new-object -TypeName S1000D.DataModule_401
$parserPM = new-object -TypeName S1000D.PublicationModule_401
$outputFolder = "d:\KC46 Staging\Scripts\Report Generators\Outputs\BoilerPlateReports-2019"
$reportName = "The Lana Report.csv"
$outputPath = "$outputFolder\$reportName"

# ********************************************************************
[string[]] $PubList   = @("KC46","ABDR","ACS","AMM","ARD","ASIP","FIM","IPB","LOAPS","NDT","SIMR","SPCC","TC","WUC","SSM","SWPM") | Sort-Object
[string[]] $PubList   = @("KC46","AMM","ARD","FIM","NDT","TC","SSM") | Sort-Object

$pathToCsdb = "d:\KC46 Staging\Production\Manuals"
$pmcs = @{}
$objs = @()
<##>

function Get-DMCReportEntry($Book,$DMC, $TechName, $IssueNumber, $IssueYear, $IssueMonth, $IssueDay, $DMC_ID,$EMODBuild,$ICNCollection,$ICNCount) {
 return [pscustomobject][ordered]@{Book=$Book; "DMC"=$DMC;TechName=$TechName;"IssueNumber"=$IssueNumber;IssueYear=$IssueYear;IssueMonth=$IssueMonth;IssueDay=$IssueDay;DMC_ID=$DMC_ID;EMODBuild=$EMODBuild;ICNCollection=$ICNCollection;ICNCount=$ICNCount;}
}

foreach ($Pub in $PubList)
{
    $bookPathToDM = "$pathToCsdb\$pub\S1000D\SDLLIVE"
    $bookPathTogRAPHICS = "$pathToCsdb\$pub\Illustrations\Illustrations"
    $dmFiles = gci -Path $bookPathToDM -Filter DMC*.*
    $pmFiles = gci -Path $bookPathToDM -Filter PMC*.* | Sort-Object -Descending |Select -First 1

    # Get the PMC
    $pmFile = $pmFiles[0]
    $parserPM.ParsePM($pmFile.FullName)
    $dmRefs = $parserPM.DmRefs
    foreach ($dmRef in $dmRefs)
    {
        $dmc = ""
        $techName = ""
        $IssueNumber = ""
        $IssueYear = ""
        $IssueMonth = ""
        $IssueDay = ""
        $DMC_ID = ""
        $EMODBuild = ""
        $ICNCollection = ""
        $ICNCount = 0

        $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
        $dmFile = gci -Path $bookPathToDM -Filter "$dmc`*"
        if($dmFile.Count -eq 0)
        {
            $dmc
            exit
        }

        $parserDM.ParseDM($dmFile.FullName)
        $techName = $parserDM.TechName
        $IssueNumber = $parserDM.IssueInfo.issueNumber

        $issueDateNode = $parserDM.Dmodule.dmodule.identAndStatusSection.dmAddress.dmAddressItems.issueDate

        $IssueYear = $issueDateNode.year
        $IssueMonth = $issueDateNode.month
        $IssueDay = $issueDateNode.day
        $DMC_ID = $parserDM.DmId.'#text'
        $comment = ""
        if($parserDM.Dmodule.dmodule.'#comment'.Contains("EMOD"))
        {
            $comment = [string]($parserDM.Dmodule.dmodule.'#comment'[3]).Trim()
            if(! $comment.StartsWith("EMOD"))
            {
                $comment
                exit
            }        
        }
        $EMODBuild = $comment

        # /dmodule/content/description/levelledPara/levelledPara/figure/graphic
        $ICNArray = @()
        $graphics = $parserDM.Dmodule.dmodule.content.SelectNodes("//graphic")
        if($graphics.Count -gt 0)
        {
            foreach ($graphic in $graphics)
            {
                $ICNArray += $graphic.infoEntityIdent
            }
            $ICNCount = $graphics.Count
            $ICNCollection = ($ICNArray|sort-object|group|Select -ExpandProperty Name) -join ","            
        }
        $obj = Get-DMCReportEntry -Book $Pub  -DMC $dmc -TechName $techName -IssueNumber $IssueNumber -IssueYear $IssueYear -IssueMonth $IssueMonth -IssueDay $IssueDay -DMC_ID $DMC_ID -EMODBuild $comment -ICNCollection $ICNCollection -ICNCount $ICNCount
        $objs += $obj
        $obj = $null
    }
}
$prop1 = @{Expression='Book'; Ascending=$true }
$prop2 = @{Expression='DMC'; Ascending=$true }
$prop3 = @{Expression='TechName'; Ascending=$true }
$prop4 = @{Expression='Preferred'; Descending=$true }
$prop5 = @{Expression='PN'; Ascending=$true }

$objs.GetEnumerator() | Sort-Object -Property $prop1, $prop2,$prop3,$prop4,$prop5 | Export-Csv $outputPath -NoTypeInformation
$outputPath
# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"					
					