#region Monitoring and Distribution
Set-Variable -Name "PathToDropBox" -Value "\\emodmgt\S1000D_Dropbox_Prod\Inbound" -Description "EMOD Dropbox Location for DDN Pickup"
Set-Variable -Name "ScriptsFolder" -Value "C:\KC46 Staging\Scripts" -Description "Location for all the standard KC46 Base Scripts"
Set-Variable -Name "BasePath" -Value "F:\KC46 Staging\production" -Description "Root location to all the production data within the CSDB"
Set-Variable -Name "ManualsBasePath" -Value "$BasePath\Manuals" -Description "Root location to all the manuals in the CSDB for a given environment"
Set-Variable -Name "SourceFolder" -Value "C:\KC46 Staging\Production\Archives\Source" -Description "Root location to archived DDN zip files that are the source to the CSDB"
Set-Variable -Name "UnpackingFolder" -Value "$SourceFolder\UnpackHere" -Description "Root location to all the manuals in the CSDB for a given environment"
Set-Variable -Name "actionReq" -Value $false -Description "Used in calls to CEERS"
Set-Variable -Name "ABDR" -Value $false -Description "Used to select the correct publication module" 
Set-Variable -Name "LiveContentDataFolder" -Value "F:\LiveContentData\" -Description "Location of artifacts used by the SDL Publisher engine during the creation of the IETM"
Set-Variable -Name "archiveRootFolder" -Value "C:\KC46 Staging\Production\Archives" -Description "Location for all the inbound source data going into the CSDB"
Set-Variable -Name "buildsRootFolder" -Value "C:\KC46 Staging\Production\Archives\Builds" -Description "Destination for the finished IETM builds"
Set-Variable -Name "source_BaseLocation" -Value $ManualsBasePath -Description "Location for all the inbound source data going into the CSDB"
#endregion

#region BASE GLOBALS
$global:sd = Get-Date
$global:startTime = $global:sd.Year.ToString()
$global:startTime += "-" + $global:sd.Month.ToString().PadLeft(2,"0")
$global:startTime += "-" + $global:sd.Day.ToString().PadLeft(2,"0")
$global:startTime += "-" + $global:sd.Hour.ToString().PadLeft(2,"0")
$global:startTime += "-" + $global:sd.Minute.ToString().PadLeft(2,"0")
$global:startTime += "-" + $global:sd.Second.ToString().PadLeft(2,"0")

$global:evtID = 60000
$Global:environment = "Production"  # ************* Production  ************#
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
# CEERS
$global:siteURL = "https://collab.web.boeing.com/sites/KC46TankerTechPubs/CEERS"
$global:listName = "KC46 - Tanker Tech Pubs CEERS"
$global:fullLogPath = "$ScriptsFolder\logs\KC46 - $environment - Build And Publish Technical Manuals - $startTime.log"

# Where the source S1000D data is located that will eventually become an IETM
$Global:KC46DataRoot = "F:\KC46 Staging"

# Packaging and Unpackaging variables
$Global:sourceRoot = $SourceFolder
$Global:unpackLocation = $UnpackingFolder
#endregion

#region Data Management
Set-Variable -Name "autoplay_cdonlyLocation" -Value "$ScriptsFolder\autoplay_cdonly.exe" -Description "Loction of the SDL startup EXE"
Set-Variable -Name "WDMPmcInputFolder" -Value "$BasePath\Manuals\WDM\S1000D\SDLLIVE" -Description "Location for all the inbound source data going into the CSDB"


#LiveContent build variables
$global:dmDestination = $LiveContentDataFolder + "source\KC46\source_xml"
#$global:masterPMCFullPath = "$dmDestination\PMC-1KC46-81205-99999-00.xml"
$propertiesFileLocation ="$LiveContentDataFolder" + "publications\KC46\properties.xml"
$global:configPath  = "$LiveContentDataFolder" + "common\config - KC46 Only"
$pubDestination = @()
$pubDestination += $LiveContentDataFolder + "publications\KC46"
$pubDestination += $LiveContentDataFolder + "publications\WIRING"

# PREDEFINE TARGET LOCATIONS FOR EVERYTHING!

#endregion

#region Common Items
$a = (Get-Host).PrivateData
$a.WarningBackgroundColor = "Yellow"
$a.WarningForegroundColor = "Black"
$rn = "<br/>"
$a.ErrorBackgroundColor = "red"
$a.ErrorForegroundColor = "White"
#endregion

#region A
#endregion

#region B
#endregion

#region Reporting

#endregion

