CLS
$ErrorActionPreference = "Stop"
$error.Clear()
				
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$parserDM = new-object -TypeName S1000D.DataModule_401
# ********************************************************************
[string[]] $PubList   = @("KC46","ABDR","ACS","AMM","ARD","ASIP","FIM","IPB","LOAPS","NDT","SIMR","SPCC","TC","WUC","SSM","SWPM", "WDM") | Sort-Object
$driveFiles = "F:\KC46 Staging\Production\Manuals" 
$ramFiles   = "D:\Shared\IDE cd sets\Releases\Aug 2018 - Release 10.2\CSDB\DVD"
$dmSuffix =  "S1000D\SDLLIVE"

foreach ($Pub in $PubList)
{
$Pub
    $dPath = "$driveFiles\$pub\$dmSuffix"
    $rPath = "$ramFiles\$pub\$dmSuffix"
    $dFiles = gci -path $dPath -Filter DMC*.XML
    foreach ($F in $dFiles)
    {
        $parserDM.ParseDM($F.FullName)
        $issNum1 = $parserDM.IssueInfo.issueNumber
        $name = $F.Name
        $releasedFile = gci -Path $rPath -Filter $name
        $parserDM.ParseDM($releasedFile[0].FullName)
        $issNum2 = $parserDM.IssueInfo.issueNumber
        if($issNum1 -ne $issNum2)
        {
            exit
        }
    }

}



# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"					
					