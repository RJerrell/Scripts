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

[string[]] $PubList   = @("KC46", "ACS", "ARD", "AMM", "FIM", "IPB", "LOAPS", "MOM", "NDI",  "NDT", "SIMR", "SPCC",  "SRM", "SSM", "SWPM", "WUC", "WDM")
$masterList = @("")

foreach ($pub in $PubList)
   {
       $files = gci -Path "$source_BaseLocation\$pub\S1000D\S1000D\DMC*.*"
       $masterList += $files.Name
   }
   $masterList