CLS
$ErrorActionPreference = "Stop"
$error.Clear()
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\KC46 Staging\Scripts\Modules;"
Import-Module -Name "KC46Common" -Verbose -Force					
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'

Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$parserDM = new-object -TypeName S1000D.DataModule_401
$parserPM = new-object -TypeName S1000D.PublicationModule_401

# ********************************************************************
[string[]] $PubList   = @("KC46","ABDR","ACS","AMM","ARD","ASIP","FIM","IPB","LOAPS","NDT","SIMR","SPCC","TC","WUC","SSM","SWPM") | Sort-Object
[string[]] $PubList   = @("KC46","AMM","ARD","FIM","NDT","TC","SSM","WDM") | Sort-Object
$pathToCsdb = "F:\KC46 Staging\Production\Manuals"
$pmcs = @{}
$objs = @()
<##>

function Get-DMCReportEntry($Book,$DMType, $DMC, $TechName, $IssueNumber, $IssueYear, $IssueMonth, $IssueDay, $DMC_ID,$EMODBuild,$ICNCollection,$ICNCount) {
 return [pscustomobject][ordered]@{Book=$Book; DMType=$DMType;"DMC"=$DMC;TechName=$TechName;"IssueNumber"=$IssueNumber;IssueYear=$IssueYear;IssueMonth=$IssueMonth;IssueDay=$IssueDay;DMC_ID=$DMC_ID;EMODBuild=$EMODBuild;ICNCollection=$ICNCollection;ICNCount=$ICNCount;}
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
    $pmTitle = $parserPM.PMTitle
    $dmRefs = $parserPM.DmRefs
    $comment = ""
    $pmFile.FullName
    $COMMENTS = $parserPM.Pmodule.pm.'#comment'
    foreach($C IN $COMMENTS)
    { 
        if($C.Contains("EMOD BUILD"))
        {
            $comment = $C.Replace("EMOD BUILD: ", "").Trim()
            break
        }
    }
    $obj = Get-DMCReportEntry -Book $Pub -DMType "pm" -DMC $pmFile.Name -TechName $pmTitle -IssueNumber $parserPM.IssueInfo.issueNumber -IssueYear $parserPM.IdentAndStatusSection.pmAddress.pmAddressItems.issueDate.year -IssueMonth $parserPM.IdentAndStatusSection.pmAddress.pmAddressItems.issueDate.month -IssueDay $parserPM.IdentAndStatusSection.pmAddress.pmAddressItems.issueDate.day -DMC_ID "" -EMODBuild $comment -ICNCollection "" -ICNCount 0
    
    $objs += $obj

    $obj = $null

    foreach ($dmRef in $dmRefs)
    {
        $dmc = ""
        $dmType = ""
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
        $dmType = $parserDM.DmType
        $techName = $parserDM.TechName
        
        $IssueNumber = $parserDM.IssueInfo.issueNumber.ToString()

        $issueDateNode = $parserDM.Dmodule.dmodule.identAndStatusSection.dmAddress.dmAddressItems.issueDate

        $IssueYear = $issueDateNode.year
        $IssueMonth = $issueDateNode.month
        $IssueDay = $issueDateNode.day
        $DMC_ID = $parserDM.DmId.'#text'
        $dmComments = $parserDM.Dmodule.dmodule.'#comment'
        $dmComment = ""
        foreach($dC IN $dmComments)
        { 
            if($dC.Contains("EMOD BUILD"))
            {
                $dmComment = $dC.Replace("EMOD BUILD: ", "").Trim()
                break
            }
        }
        
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
        $obj = Get-DMCReportEntry -Book $Pub -DMType $dmType -DMC $dmc -TechName $techName -IssueNumber $IssueNumber -IssueYear $IssueYear -IssueMonth $IssueMonth -IssueDay $IssueDay -DMC_ID $DMC_ID -EMODBuild $dmComment -ICNCollection $ICNCollection -ICNCount $ICNCount
        $objs += $obj
        $obj = $null
    }
}
$prop1 = @{Expression='Book'; Ascending=$true }
$prop2 = @{Expression='DMC'; Ascending=$true }

$outputPath = "C:\KC46 Staging\Scripts\Report Generators\Outputs\BoilerPlateReports\KC46 Tanker - CSDB Master Inventory for all CAS Books - $startTime.csv"
$objs.GetEnumerator() | Sort-Object -Property $prop1, $prop2 | Export-Csv $outputPath -NoTypeInformation
# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"					
					