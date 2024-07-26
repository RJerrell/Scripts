#region Startup Variables
# Which environment: Production, BetaTR ?
$Global:environment   = "BetaTR"    # **************** Defaulted to BETA *****************#
# $Global:environment = "Production" # *************   Override to Production  ************#

#endregion

#region Variables


# Where the source S1000D data is located that will eventually become an IETM
$Global:KC46DataRoot = "C:\KC46 Staging"

# This variable is used to select a LiveContent configuration definition ONLY!  
# We have 2: KC46 and KC46_All
# More can be defined as needed: See Ed or Roger for more information about Live Content WIETM.XML definitions

$Global:commonRoot = "KC46_All" # ****************Defaulted to BETA publication location*****************
if($environment -eq "Production")
{
    $commonRoot = "KC46" # ****************Defaulted to BETA publication location*****************
}

# BASIC FOLDER STRUCTURES THAT ARE PRETTY HARD CODED AND EXPECTED TO BE THERE.  IF THESE CHANGE, IT ALL FALLS DOWN.
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"

#endregion


$pmEntryList = @{""=""}
$filesDMCodeList = @{""=""}



[string[]] $PubList   = @($commonRoot, "ABDR", "ACS", "ARD", "AMM", "FIM", "IPB", "LOAPS", "MOM", "NDI",  "NDT", "SIMR", "SPCC",  "SRM", "SSM", "SWPM", "TC", "WUC", "WDM")

foreach($pub in $PubList)
{
    $pmcPath = "$source_BaseLocation\$pub\S1000D\S1000D\PMC*.xml"
    
    $files = Get-ChildItem -Path $pmcPath
    foreach($file in $files)
    {
        $pm = [XML] (Get-Content $file.FullName)
    }
}