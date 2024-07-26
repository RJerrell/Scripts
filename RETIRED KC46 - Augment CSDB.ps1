cls
$sd = Get-Date
$ErrorActionPreference = "SilentlyContinue"
$error.Clear()
$a = (Get-Host).PrivateData
$a.WarningBackgroundColor = "Yellow"
$a.WarningForegroundColor = "Black"
$rn = "<br/>"
$a.ErrorBackgroundColor = "red"
$a.ErrorForegroundColor = "White"

$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\KC46 Staging\Scripts\Modules;"
$env:PSModulePath

#Get-Module -ListAvailable
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "KC46S1000DRules" -Verbose -Force
Import-Module -Name "KC46DataManagement" -Verbose -Force

$global:startTime = (Get-Date -Format yyyy-MM-dd-HH-mm-ss)
$global:evtID = 60000
$global:fullLogPath = "$PSScriptRoot\logs\KC46 - $environment - Build And Publish Technical Manuals - $startTime.log"
[bool] $actionReq = $false

#region: Edit these startup variables before running!!!

$Global:environment = "Production"  # *************   Override to Production  ************#

# Set these 2 values for testing to false
[bool] $augmentData = $false
[bool] $augmentICNs = $false
[bool] $incrementBuildNumber = $false
#endregion

#region: Common System Variables
# CEERS
$global:siteURL = "https://collab.web.boeing.com/sites/KC46TankerTechPubs/CEERS"
$global:listName = "KC46 - Tanker Tech Pubs CEERS"

# Where the source S1000D data is located that will eventually become an IETM
$Global:KC46DataRoot = "C:\KC46 Staging"

# This variable is used to select a LiveContent configuration definition ONLY!  
# We have 2: KC46 and KC46_All
# More can be defined as needed: See Ed or Roger for more information about Live Content WIETM.XML definitions

$commonRoot = "KC46" # ****************Defaulted to BETA publication location*****************

# BASIC FOLDER STRUCTURES THAT ARE PRETTY HARD CODED AND EXPECTED TO BE THERE.  IF THESE CHANGE, IT ALL FALLS DOWN.
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"

#Augmentation variables
$augmentorPath = "C:\KC46\Utilities\Augmentor\$environment"
$augmentorFileName = "AugmentS1000DConsole.exe"

#LiveContent build variables
$global:dmDestination = $LiveContentDataFolder + "source\$commonRoot\source_xml"
$propertiesFileLocation ="$LiveContentDataFolder" + "publications\$commonRoot\properties.xml"
$global:configPath  = "$LiveContentDataFolder" + "common\config - $commonRoot Only"

# Graphics Related
#ISODraw Arguments list settings
$figuresWithICNEmbedded = $pubDestination[0] + "\Figures"
$figuresWithICNEmbedded_tmp = "$source_BaseLocation\Figures_tmp"
[string[]] $icnBatchFilePath_args = @("""$figuresWithICNEmbedded_tmp""" , """$figuresWithICNEmbedded""")
$icnBatchFilePath = "$PSScriptRoot\KC46-ICNBranding.bat" # Name and location of the ICN Branding batch file

[string[]] $PubList = @($commonRoot, "ABDR", "ACS", "AMM", "ARD", "FIM", "IPB", "LOAPS", "NDT", "SIMR", "SPCC", "SSM", "SRM", "SWPM", "TC", "WUC", "WDM")
# [string[]] $PubList = @("AMM")
#endregion

#REGION AUGMENT THE ENTIRE CSDB
<#
    1. APPLY ALL THE BDS/USAF AGREED UPON BUSINESS RULES TO THE CSDB THAT WERE NOT DONE BY SOURCE DATA PROVIDERS
    2. RUN THE BDS AUGMENTOR CONSOLE APPLICATION THAT INSERTS DATA MARKINGS, ACRONYMS, REQUIRED PERSONS, AND VERIFICATION 
    TAGS BASED ON THE CURRENT C & V DATA IN OUR BOEING C&V SHAREPOINT.
    3. NOTE: AS FAR AS AUGMENTATION GOES, ICN NUMBERS ARE ADDED DURING THE BUILD.
#>
#attrib -R /s "$source_BaseLocation\*.*"

# 1. APPLY ALL THE BDS/USAF AGREED UPON BUSINESS RULES TO THE CSDB THAT WERE NOT DONE BY SOURCE DATA PROVIDERS
$m4 = "Enforce Business Rules in each file in the CSDB"
# Change safety tag values to "Warning / Safety Tag" in all data modules before doing any other processing                         
Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:455" -strActivity "$m4 : Started"  -stage "Business Rules Enforcement" -actionReq $actionReq
Set-KC46BusinessRules_wf -source_BaseLocation $source_BaseLocation -PubList $PubList
Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:455" -strActivity "$m4 : Completed"  -stage "Business Rules Enforcement" -actionReq $actionReq

#Cleanup duplicate ESDS markers
& 'C:\KC46 Staging\Scripts\KC46 - Cleanup Duplicate ESDS symbols if present.ps1'

# 2. RUN THE BDS AUGMENTOR CONSOLE APPLICATION THAT INSERTS DATA MARKINGS, ACRONYMS, REQUIRED PERSONS, AND VERIFICATION 
# TAGS BASED ON THE CURRENT C & V DATA IN OUR BOEING C&V SHAREPOINT.
$m44 = "Entire CSDB console based augmentation : "
Submit-LogEntry -fullLogPath $fullLogPath -Message $m44  -EventID $evtID -caller  $m44 
Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:455" -strActivity "$m44 started"  -stage "AugmentationProcess" -actionReq $actionReq
Start-AugmentationProcess -augmentorPath $augmentorPath -augmentorFileName $augmentorFileName                   
Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:455" -strActivity "$m44 completed" -stage "AugmentationProcess" -actionReq $actionReq

# 3. Scrub the data for bad characters
Resolve-BadCharacters

$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Total Seconds to complete:`t" + $x.TotalSeconds
"Process completed"