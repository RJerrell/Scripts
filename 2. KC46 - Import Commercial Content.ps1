Clear-Host
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'

Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "KC46DataManagement" -Verbose -Force
Import-Module -Name "KC46CSDBManager" -Verbose -Force
Import-Module -Name "KC46Augmentor" -Verbose -Force
Import-Module -Name "KC46DataDistribution" -Verbose -Force
Import-Module -Name "KC46S1000DRules" -Verbose -Force

# *** *** *** *** Variable for this script only *** *** *** *** *** ******
Set-Variable -Name "ddnFN" -Value "" -Description "Location for all the inbound source data going into the CSDB"
Set-Variable -Name "ddnShortName" -Value "" -Description "Location for all the inbound source data going into the CSDB"
Set-Variable -Name "baselineDistribute" -Value $false -Description "SETTING THIS FLAG TO TRUE WILL DROP ALL DATA FROM THE CSDB AND REBUILD IT FROM THE BASELINE TO PRESENT"
Set-Variable -Name "zipsInUnpackingFolder" -Value (Get-ChildItem -path "$unpackingFolder\*.zip") -Description "Denotes any zip files in the UNPACKING folder"
Set-Variable -Name "ddnFiles" -Value (Get-ChildItem -Path "$pathToDropBox\*.zip") -Description "Get fileinformation about all the zip files located in the inbound folder"

$msg = "Checking the production drop folder for new data to process."
Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :"  -strActivity $msg -stage " Monitoring" -false $false

# ZIP files in the unpacking folder
$zips = Get-ChildItem -path "$unpackingFolder\*.zip"

# Incoming ZIP files
$ddnFiles = Get-ChildItem -Path "$pathToDropBox\*.zip"

<# 1 . Get the DDN  and archive it #>
if($zips.Count -gt 0)
{
    foreach ($zip in $zips)
    {
        Move-Item -Path  $zip.FullName -Destination $sourceFolder -Force -Verbose
    }
}

if($ddnFiles.Count -eq 1)
{
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :"  -strActivity "Archival: " + $zip.Name + "" -stage "Archival" -actionReq $false
        # Move the data onto our build machine for processing
        $ddnFN = $ddnFiles[0].FullName        
        $ddnShortName = $ddnFiles[0].Name
        if((Test-Path -Path "$sourceFolder\$ddnShortName"))
        {
            Remove-Item -Path "$sourceFolder\$ddnShortName" -Force
        }
        Move-Item -Path $ddnFN -Destination $sourceFolder -Verbose -Force -ErrorAction Inquire
        Copy-Item -Path "$sourceFolder\$ddnShortName" -Destination $unpackingFolder -Verbose -Force -ErrorAction Inquire

        # Insure the DDN filename is embedded in the DDNProcessingOrder.xml for proper didstribution, if needed
        Set-Location $unpackingFolder
        # Test to see if the ddn value is already in the file
        $ddnShortName = $ddnFiles[0].BaseName
        $ddnProcessingXml = New-Object System.Xml.XmlDocument
        $ddnProcessingXml.Load("$unpackingFolder\DDNProcessingOrder.xml")
        $ddnExistsInFile = $false
        $ddnExistsInFile = $ddnProcessingXml.OuterXml.Contains($ddnShortName)
        if( ! $ddnExistsInFile)
        {
            $child = $ddnProcessingXml.CreateElement("ddn")
            $child.InnerText = $ddnShortName.Replace(".zip", "")
            $ddnProcessingXml.root.AppendChild($child)
            
            #ATTRIB -R "$unpackingFolder\DDNProcessingOrder.xml"
            Set-ItemProperty -Path "$unpackingFolder\DDNProcessingOrder.xml" -Name IsReadOnly -Value $false -Force -Verbose
            $ddnProcessingXml.Save("$unpackingFolder\DDNProcessingOrder.xml")
        }
        else
        {            
            "DDN exists in the DDNProcessingOrder.xml file"
        }
    }
else
{
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :"  -strActivity "No new data to process. 1 and only 1 zip file can be processed at a time." -stage "Archival" -false $false
        "No new data to process.  1 and only 1 zip file can be processed at a time."
}

# $ddnShortName = "DDN-1KC46-AAAZZ-81205-2017-00003"

<# 3 . Unzip the new DDN #>
if($ddnShortName.Length -gt 0 -and (Test-Path -Path "$unpackingFolder\$ddnShortName"))
{
    Remove-Item -Path "$unpackingFolder\$ddnShortName"
}

if(!(Test-Path -Path "$unpackingFolder\$ddnShortName"))
{
    md "$unpackingFolder\$ddnShortName"
}

if((Test-Path -Path "$unpackingFolder\$ddnShortName") -and ($ddnShortName.Length -gt 0) -and ($ddnFN.Length -gt 0))
{
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :"  -strActivity "Recursive Unzip executing -- unpacking the new data" -stage "Archival" -false $false
        cd $unpackingFolder        
        & '.\Recursive Unzip.ps1'
}
else
{
    "Folder ($unpackingFolder\$ddnShortName) missing"
    exit
}
<# 

4. Add code here to augment the data before it ever enters the CSDB
4.1 Augment the current DDN folder using the Update-1Folder function in the Augmentor.psm1 module
    - Acronyms are inserted
    - Data markings applied
    - SERD Database preferences set
#>

    Update-1Folder -basepath "$unpackingFolder\$ddnShortName"

#4.2 Augment the ICNs received
    
    $figuresWithICNEmbedded = "$unpackingFolder\$ddnShortName\Figures"
    $figuresWithICNEmbedded_tmp = "$unpackingFolder\$ddnShortName\Figures_tmp"

    if(!(Test-Path -Path $figuresWithICNEmbedded))
    {
        md $figuresWithICNEmbedded
    }
    else
    {
        Remove-Item $figuresWithICNEmbedded -Recurse -Force 
        SLEEP -Seconds 15 
        md $figuresWithICNEmbedded        
    }
    if(!(Test-Path -Path $figuresWithICNEmbedded_tmp))
    {
        md $figuresWithICNEmbedded_tmp
    }
    else
    {
        Remove-Item $figuresWithICNEmbedded_tmp -Recurse -Force 
        SLEEP -Seconds 15 
        md $figuresWithICNEmbedded_tmp        
    }   


    Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:489" -strActivity "$msg99 started" -stage "Illustrations Mgmt" -actionReq $actionReq
   
    if( ! $ddnShortName.ToUpper().Contains( "-P"))
    {
        [string[]] $icnBatchFilePath_args = @("""$figuresWithICNEmbedded_tmp""" , """$figuresWithICNEmbedded""")
        $icnBatchFilePath = "$scriptsFolder\KC46 - ICN Branding.bat" # Name and location of the ICN Branding batch file
            
        # Move all the ICNs in this DNN to a tmp location for processing
        $ff = Get-ChildItem -Path "$unpackingFolder\$ddnShortName" -Filter *.CGM -Recurse 
        foreach( $f in $ff)
        {
            $f.FullName
            $fsname = $f.Name            
            copy-item -Path $f.FullName -Destination "$figuresWithICNEmbedded_tmp\$fsname" -ErrorAction Stop
        }
        #call the function in the KC46DataManagement psm1 module
        Add-ICN_Numbers_ToGraphics -icnBatchFilePath $icnBatchFilePath -figuresWithICNEmbedded $figuresWithICNEmbedded -icnBatchFilePath_args $icnBatchFilePath_args     
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:489" -strActivity "$msg99 completed" -stage "Illustrations Mgmt" -actionReq $actionReq
    }
    else
    {
        $fIPB = Get-ChildItem -Path "$unpackingFolder\$ddnShortName" -Filter *.CGM
        $counterIPB = 0
        $fIPB.Count
        foreach( $f in $fIPB)
        {
            $f.FullName
            $fsname = $f.Name            
            copy-item -Path $f.FullName -Destination "$figuresWithICNEmbedded\$fsname" -ErrorAction Stop
            $counterIPB ++
            $counterIPB
            
        }
        if($counterIPB -ne $fIPB.Count)
        {
           "Graphics processing failed" 
            exit
        }
    }
   
    # Call the function in the DataDistribution psm1 module.  Moved the ICNs into the CSDB within the appropriate manual
    Push-AllICNs -figuresWithICNEmbedded $figuresWithICNEmbedded

#4.3 Apply business rules
    
if( $ddnShortName.ToUpper().Contains( "-P"))
{
    Set-KC46BusinessRules_IPB -path "$unpackingFolder\$ddnShortName"
}
else
{
    Set-KC46BusinessRules -path "$unpackingFolder\$ddnShortName"
}
<# 5. Redistribute the entire CSDB #>
Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :"  -strActivity "Reconstructing the CSDB with the latest content" -stage "Distribution" -false $false
cd $unpackingFolder

if($ddnShortName.Length -gt 0)
{
        if($baselineDistribute)
        {
            # The default is to do this only once in the life of the CSDB or as reset of the CSDB when things go bump in the night.
            # & '.\Baseline Distribute.ps1'
            exit
        }
        else
        {
            & '.\Delta Distribute.ps1' -DDN $ddnShortName
        }
    }

# Insure the DMLs for this release are ready to be processed
$global:exitCode = Set-DMLWarningForm

if($global:exitCode1 -eq 1)
{
        "Cancelled!"
        Exit
    }
if($global:exitCode1 -eq 0)
{
        "Go"
}
 
<# 6 . Load RAM Drive #>
Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :"  -strActivity "Loading the CSDB into a RAM drive to imporve build performance  " -stage "Data Preparartion" -false $false
cd $scriptsFolder
& '.\4. KC46 - Load CSDB into a RAM Disk.ps1'

<#7 . Build and publish the IETM #>
Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :"  -strActivity "IETM Build AND pUBLISH Process: Begin the build process for the current IETM  " -stage "IETM Build Process" -false $false
cd $scriptsFolder
& '.\5. KC46 - Build and  Publish an IETM.ps1'
    
# *** *** *** *** *** *** *** *** *** ******
$ed = Get-Date
$x = $ed.Subtract($sd)

"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "$PSScriptRoot :"  -strActivity "IETM Build AND pUBLISH Process: Complete.\r\n Days:$x.TotalDays, Hours: $x.TotalHours, Minutes: $x.TotalMinutes" -stage "IETM Build Process" -false $false

"Process completed"