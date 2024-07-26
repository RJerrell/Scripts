<#
Title: KC46 - DATAMANAGEMENT - RAM DRIVE VERSION.PS1
Author: Roger Jerrell
Date Created: 03-30-2014
Purpose: This script manages several processes used to ultimately create and publish the IETM
Description of Operation: 
    1. Set the paramters for ICN numbering to $true
    2. Insure the $Publist represents the correct list of books to match the master PMC(s) plural.
    3. Insure the Ram drive is loaded with the data you want to build and publish
    4. Run / execute the script
    5. Go to "C:\KC46 Staging\Production\Archives\builds" to see the published IETM
#>
cls

# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'

$ErrorActionPreference = "Stop"
$error.Clear()
Import-Module -Name "KC46Common" -Verbose -Force
#Import-Module -Name "KC46S1000DRules" -Verbose -Force
Import-Module -Name "KC46DataManagement" -Verbose -Force

 "STARTED AT : " + $startTime
$releaseNumber = "11.1"
$global:publishIETMToThisLocation  = "$buildsRootFolder\$startTime\IETM"

Set-Variable -Name "WDMPmcName" -Value (gci -Path "$WDMPmcInputFolder\PMC*.XML" | Sort-Object -Descending | Select -First 1) -Description "Selects the latest version of the WDM publication module"

function Start-KC46DatamanagementProcessing
{
    # These are the acronyms matching the names of the file folders on the drive and representing each manual in the IETM
    
    [string[]] $PubList   = @("KC46","ABDR","ACS","AMM","ARD","ASIP","FIM","IPB","LOAPS","NDT","NDTS","SIMR","SPCC","TC","WUC","SSM","SWPM", "WDM") | Sort-Object
    # [string[]] $PubList   = @(KC46","ABDR","ACS","AMM","ARD","ASIP","FIM","IPB","LOAPS","NDT","NDTS","SIMR","SPCC","TC","WUC","SSM","SWPM") | Sort-Object
    
    # BDS Military Manual Testing
    # [string[]] $PubList   = @("KC46","ABDR","ACS","AMM","ASIP","LOAPS","NDT","NDTS","SIMR","SPCC","TC","WUC") | Sort-Object


    $msg= "$env:COMPUTERNAME - Initializing $environment build and publishing of the S1000D IETM."

    Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "****************$env:COMPUTERNAME - ****************$rn Initializing $environment Build and Publishing of the S1000D IETM$rn****************$env:COMPUTERNAME - ****************" -stage "Initialization" -actionReq $actionReq
    Set-DefaultFolders -dmDestination $dmDestination -pubDestination $pubDestination   -fullLogPath $fullLogPath -evtID $evtID

    $figuresWithICNEmbedded = $pubDestination[0] + "\figures"
    Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting $rn Start-KC46DataManagement" -stage "Data Preparation" -actionReq $actionReq
    Start-KC46DataManagement -Publist $PubList -source_BaseLocation $source_BaseLocation -LiveContentDataFolder $LiveContentDataFolder -figuresWithICNEmbedded $figuresWithICNEmbedded -dmDestination $dmDestination -pubDestination $pubDestination -fullLogPath $fullLogPath -siteUrl $siteUrl -listName $listName
    Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Completed $rn Start-KC46DataManagement" -stage "Data Preparation" -actionReq $actionReq
    if($PubList.Contains("WDM"))
    {
        # Before we start, we have to insure the WDM PMC will be usable for the SDL viewer
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting - Optimize-WDMPMC" -stage "Data Preparation" -actionReq $actionReq
        $wdmPMCShortName = $WDMPmcName.Name
        Optimize-WDMPMC -pmcName $WDMPmcName -pmcShortName "$dmDestination\$wdmPMCShortName"
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Completed - Optimize-WDMPMC" -stage "Data Preparation" -actionReq $actionReq
    }

    # Copy these files...
    Copy-Item -Path "$source_BaseLocation\KC46\Illustrations\Illustrations\*.*"  -Destination $figuresWithICNEmbedded -Verbose -Force
    Copy-Item -Path "$source_BaseLocation\KC46\S1000D\FrontMatter\Release DML Set\Highlights\DMC*-00UA-D.xml"  -Destination $global:dmDestination -Verbose -Force

    # This file is for the Common folder
    Copy-Item -Path "$source_BaseLocation\KC46\S1000D\FrontMatter\DMC-*.xml"  -Destination $global:dmDestination -Verbose -Force
    Copy-Item -Path "$source_BaseLocation\KC46\S1000D\Master PM\DML-1KC46-Applicability.xml"  -Destination $global:dmDestination -Verbose -Force

    # Build IETM
    Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine)-strActivity "Starting $rn Start-IETM_Build  to $global:publishIETMToThisLocation" -stage "IETM Build" -actionReq $actionReq
   
    # Copy the Master Publication module to the workspace
    Copy-Item -Path "$source_BaseLocation\KC46\S1000D\Master PM\PMC-1KC46-81205-99999-00.xml" -Destination "$global:dmDestination\PMC-1KC46-81205-99999-00.xml" -Verbose -Force -ErrorAction SilentlyContinue
    
    # Special handling for WIRING manuals
    if($PubList.Contains("WDM") -or $PubList.Contains("SSM") -or $PubList.Contains("SWPM") )
    {
        <# 
            ***** ***** WORKAROUND FOR GRAPHICS IN MULTIBOOK BUILDS ***** *****
            Moving WIRING related graphics to the WIRING book ready for publishing.
        #>
        $wiringFiguresDestination = $pubDestination[1] + "\figures"
        if(! (Test-Path -Path $wiringFiguresDestination))
        {
            md $wiringFiguresDestination
        }

        $figSource = $pubDestination[0]

        Move-Item -Path "$figSource\figures\*-KW*.CGM" -Destination $wiringFiguresDestination -Verbose -Force
        Move-Item -Path "$figSource\figures\*-KR*.CGM" -Destination $wiringFiguresDestination -Verbose -Force
        Move-Item -Path "$figSource\figures\*-KU*.CGM" -Destination $wiringFiguresDestination -Verbose -Force

        Sleep -Seconds 30  # let the file move complete its processing
        Copy-Item -Path "$source_BaseLocation\KC46\S1000D\Master PM\PMC-1KC46-81205-99998-00.xml"  -Destination "$global:dmDestination" -Verbose -Force
        Copy-Item -Path "$global:configPath\wietmsdMultiBook.xml"  -Destination "$global:configPath\wietmsd.xml" -Verbose -Force
        
        # Insure pmCode references are accurate
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting - Set-PMCodes" -stage "Data Preparation" -actionReq $actionReq
        
        Set-PMCodes -masterPMCFullPath "$global:dmDestination\PMC-1KC46-81205-99999-00.xml"
        Set-PMCodes -masterPMCFullPath "$global:dmDestination\PMC-1KC46-81205-99998-00.xml"

        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Completed - Set-PMCodes" -stage "Data Preparation" -actionReq $actionReq
        
        # Update the issue date, etc  in the wietm.xml so that the landing page reflects the correct information
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting - Reset-WIETM" -stage "Data Preparation" -actionReq $actionReq
        Reset-WIETM -Path "$global:configPath\wietmsd.xml" -CSVPath "$env:TEMP\TechnicalManualVersions - $startTime.txt" -releaseNumber $releaseNumber
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Completed - Reset-WIETM" -stage "Data Preparation" -actionReq $actionReq

       
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting - Reset-AllDmRefsTitles" -stage "Data Preparation" -actionReq $actionReq
        #Reset-AllDmRefsTitles -sourceFolder $global:dmDestination
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Completed - Reset-AllDmRefsTitles" -stage "Data Preparation" -actionReq $actionReq
       
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting - Start-IETM_Build" -stage "Data Preparation" -actionReq $actionReq
        Start-IETM_Build -propertiesFileLocation $propertiesFileLocation -multiVolume $true
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Completed - Start-IETM_Build" -stage "Data Preparation" -actionReq $actionReq
    }
    else
    {
        Copy-Item -Path "$global:configPath\wietmsdSingleBook.xml"  -Destination "$global:configPath\wietmsd.xml" -Verbose -Force        
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting - Set-PMCodes" -stage "Data Preparation" -actionReq $actionReq
        Set-PMCodes -masterPMCFullPath "$global:dmDestination\PMC-1KC46-81205-99999-00.xml"
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting - Optimize-WDMPMC" -stage "Data Preparation" -actionReq $actionReq
        
        # Update the issue date, etc  in the wietm.xml so that the landing page reflects the correct information
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting - Reset-WIETM" -stage "Data Preparation" -actionReq $actionReq
        Reset-WIETM -Path "$global:configPath\wietmsd.xml" -CSVPath  "$env:TEMP\TechnicalManualVersions - $startTime.txt" -releaseNumber $releaseNumber
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Completed - Reset-WIETM" -stage "Data Preparation" -actionReq $actionReq
        
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting - Reset-AllDmRefsTitles" -stage "Data Preparation" -actionReq $actionReq
        #Reset-AllDmRefsTitles -sourceFolder $global:dmDestination
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Completed - Reset-AllDmRefsTitles" -stage "Data Preparation" -actionReq $actionReq
        
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting - Start-IETM_Build" -stage "Data Preparation" -actionReq $actionReq
        Start-IETM_Build -propertiesFileLocation $propertiesFileLocation -multiVolume $false
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Completed - Start-IETM_Build" -stage "Data Preparation" -actionReq $actionReq
    }

    Sleep -Seconds 5

    # Publish IETM
    Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Starting $rn Publish-IETM to $global:publishIETMToThisLocation" -stage "IETM Publish" -actionReq $actionReq
    if( ! (Test-Path -Path "$archiveRootFolder\builds\$startTime"))
    {
        md "$archiveRootFolder\builds\$startTime"
    }
    
    Publish-IETM -configPath $configPath -publishIETMToThisLocation $global:publishIETMToThisLocation -autoplay_cdonlyLocation $autoplay_cdonlyLocation  

    Copy-Item -Path  "$env:TEMP\TechnicalManualVersions - $startTime.txt" -Destination "$archiveRootFolder\builds\$startTime"
    Copy-Item -Path  "$ScriptsFolder\*Start Here*" -Destination $global:publishIETMToThisLocation
    
    Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :" + (Get-CurrentLine) -strActivity "Completed $rn Publish-IETM to $newLocation" -stage "IETM Publish" -actionReq $actionReq
}

Start-KC46DatamanagementProcessing

# *****************************************************************************************************

$ed = Get-Date
$x = $ed.Subtract($sd)
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"