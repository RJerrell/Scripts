# BASIC FOLDER STRUCTURES THAT ARE PRETTY HARD CODED AND EXPECTED TO BE THERE.  IF THESE CHANGE, IT ALL FALLS DOWN.
$LiveContentDataFolder =  "C:\LiveContentData\" # leave the dash on the end of this path statement
$semaphoreLocation = "$KC46DataRoot\$environment"
$unpackLocation = "$KC46DataRoot\$environment"
$archiveRootFolder = "$KC46DataRoot\$environment\Archives"
$buildsRootFolder = "$archiveRootFolder\Builds"
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"

# Packaging and Unpackaging variables
$Global:sourceRoot = "$archiveRootFolder\Source"
$Global:unpackLocation = "$sourceRoot\UnpackHere"

# LOCATION FOR ALL THE NON-CAS MANUALS THAT BDS IS AUTHORING
$militaryManualLocation = "\\KC46-CSDB-SDL\BDS$"

#Augmentation variables
$augmentorPath = "C:\KC46\Utilities\Augmentor\$environment"
$augmentorFileName = "AugmentS1000DConsole.exe"

#LiveContent build variables
$global:dmDestination = $LiveContentDataFolder + "source\$commonRoot\source_xml"
$global:masterPMCFullPath = "$dmDestination\PMC-1KC46-81205-99999-00.xml"
$propertiesFileLocation ="$LiveContentDataFolder" + "publications\$commonRoot\properties.xml"
$global:configPath  = "$LiveContentDataFolder" + "common\config - $commonRoot Only"
$pubDestination = $LiveContentDataFolder + "publications\$commonRoot"
$autoplay_cdonlyLocation = "$PSScriptRoot\autoplay_cdonly.exe"

# Graphics Related
#ISODraw Arguments list settings
$figuresWithICNEmbedded = "$pubDestination\Figures"
$figuresWithICNEmbedded_tmp = "$source_BaseLocation\Figures_tmp"
[string[]] $icnBatchFilePath_args = @("""$figuresWithICNEmbedded_tmp""" , """$figuresWithICNEmbedded""")
$icnBatchFilePath = "$PSScriptRoot\KC46-ICNBranding.bat" # Name and location of the ICN Branding batch file


# NON -S1000D VARAIBLES : CMMs, Flight documents, So called source data
$topLevelFlightDocumentsFolderFullPath = "\\kc46-csdb-sdl\c$\KC46 Staging\FlightManuals"
$topLevelNonS1000DFolderFullPath = "\\kc46-csdb-sdl\c$\KC46 Staging\NonS1000D"

# PREDEFINE TARGET LOCATIONS FOR EVERYTHING!
$global:publishIETMToThisLocation  = "$archiveRootFolder\builds\$startTime\IETM"
$global:publishNonS1000DOperations  = "$archiveRootFolder\builds\$startTime\ETM\OPERATIONS"
$global:publishNonS1000DMaintenance  = "$archiveRootFolder\builds\$startTime\ETM\MAINTENANCE"
[string[]] $PubList   = @($commonRoot)
#endregion



$a = @() # All the WDM dmCodes
$b = @() # All the dmRefs in all the books

[string[]] $PubList   = @("KC46", "ACS", "ARD", "AMM", "FIM", "IPB", "LOAPS", "NDT", "SIMR", "SRM", "SSM", "SWPM", "TC", "WUC")    

foreach ($pub in $PubList)
{
    $dmListPath = "$basePath\$pub\S1000D\S1000D\dmc*.xml"
    $dms = gci -Path $dmListPath
    foreach ($dm in $dms)
    {
        $dmc = [xml] (Get-Content -Path $dm.FullName)
        # Get a collection of all the dmCodes inside the content/refs tag within each dm
        $dmcs = $dmc.SelectNodes("//refs//dmRef")
        #$dmcs.Count
        # Get each DMC
        foreach ($dmc in $dmcs)
        {
            $modelIdentCode = $dmc.dmRefIdent.dmCode.modelIdentCode
            $systemDiffCode = $dmc.dmRefIdent.dmCode.systemDiffCode
            $systemCode = $dmc.dmRefIdent.dmCode.systemCode
            $subSystemCode = $dmc.dmRefIdent.dmCode.subSystemCode
            $subSubSystemCode = $dmc.dmRefIdent.dmCode.subSubSystemCode
            $assyCode = $dmc.dmRefIdent.dmCode.assyCode
            $disassyCode = $dmc.dmRefIdent.dmCode.disassyCode
            $disassyCodeVariant = $dmc.dmRefIdent.dmCode.disassyCodeVariant
            $infoCode = $dmc.dmRefIdent.dmCode.infoCode
            $infoCodeVariant = $dmc.dmRefIdent.dmCode.infoCodeVariant
            $itemLocationCode = $dmc.dmRefIdent.dmCode.itemLocationCode
            $ch = $systemCode
            $se = $subSystemCode+$subSubSystemCode
            $su = $assyCode

            $fileName =  "$basePath\$pub\S1000D\S1000D\$filePref-$modelIdentCode-$systemDiffCode-$systemCode-$subSystemCode$subSubSystemCode-$assyCode-$disassyCode$disassyCodeVariant-$infoCode$infoCodeVariant-$itemLocationCode.xml"
            $dataModuleCode =  "$filePref-$modelIdentCode-$systemDiffCode-$systemCode-$subSystemCode$subSubSystemCode-$assyCode-$disassyCode$disassyCodeVariant-$infoCode$infoCodeVariant-$itemLocationCode.xml"
            #$dataModuleCode
            $b += $dataModuleCode
        }
    }



    # Convert the dmCode values to a file name and addmAll the references from all books"

}
$wdmFiles = gci -Path "C:\KC46 Staging\Production\Manuals\WDM\S1000D\S1000D\DMC*.XML"
foreach($file in $wdmFiles)
{
#$file.Name
    $a += $file.Name
}

#"Union"
#$a + $b | select -uniq    #union

"Reverse Intersect"
$unreferencedList = $a | ?{$b -contains $_}   # reverse intersection
$unreferencedList | Out-File -FilePath "c:\temp\referencedWDMs.txt"
