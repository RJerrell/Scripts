cls

$Global:environment = "Production"  # *************   Override to Production  ************#
$commonRoot = "KC46"
# Where the source S1000D data is located that will eventually become an IETM
$Global:KC46DataRoot = "C:\KC46 Staging"

# BASIC FOLDER STRUCTURES THAT ARE PRETTY HARD CODED AND EXPECTED TO BE THERE.  IF THESE CHANGE, IT ALL FALLS DOWN.
$LiveContentDataFolder =  "C:\LiveContentData\" # leave the dash on the end of this path statement
$semaphoreLocation = "$KC46DataRoot\$environment"
$unpackLocation = "$KC46DataRoot\$environment"
$archiveRootFolder = "$KC46DataRoot\$environment\Archives"
$buildsRootFolder = "$archiveRootFolder\Builds"
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"

[string[]] $PubList   = @("ACS",  "AMM" )
$masterList = @("")

foreach ($pub in $PubList)
{
    gci -Path "$source_BaseLocation\$pub\S1000D\S1000D\PMC*.*" | % {$masterList += "$pub`t" + $_.Name }
}

$masterList | ft | Out-File "C:\KC46 Staging\Scripts\Report Generators\Outputs\ACS-AMM DMClIST.CSV"