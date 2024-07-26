Function Push-AllICNs
{
Param([string] $figuresWithICNEmbedded )
    
    # Attrib -R "$figuresWithICNEmbedded\*.*"
    Get-ChildItem -path $figuresWithICNEmbedded  -Filter *.cgm | %{ Set-ItemProperty -Path $_.FullName -Name IsReadOnly -Value $false -Force}
    
    $icnCollection = Get-ChildItem -Path $figuresWithICNEmbedded -Filter *.cgm

    foreach ($icn in $icnCollection)
    {
        $pub = Get-ICN_ParentBook -ICN_ShortFileName $icn.Name
        $destination = "$ManualsBasePath\$pub\Illustrations\Illustrations"
        $evar = ""
        try
        {
            $destination = $destination.Replace("F:", "C:")
            Move-Item -Path $icn.FullName -Destination $destination -Verbose -Force
        }
        catch
        {
            "Cannot move ICN to destination folder"
            $evar
            exit
        }
        
        
        
    }
}
Function Set-DeltaDistribution
{
Param([string] $ddn, [string] $kc46Highlights_dmlpath)
    $env = "Production"

    $kc46Highlights_dmlpath = "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set\$ddn"
    
    if(Test-Path -Path $kc46Highlights_dmlpath)
    {
        Remove-Item -Path $kc46Highlights_dmlpath -Force -Verbose -Recurse
        md $kc46Highlights_dmlpath
    }
    else
    {
        "$kc46Highlights_dmlpath`r`nFolder does not exist.  `r`nThe DDNProcessingOrder.xml file lists that folder and it does not exist.  `r`nThat folder will be created."
        md  $kc46Highlights_dmlpath
    }

    $dmlFiles = gci -Path "C:\KC46 Staging\Production\Archives\Source\UnpackHere\$ddn\DML*.XML" -Recurse
    foreach ($dmlFile in $dmlFiles)
    {
        if(!(Test-Path -Path $kc46Highlights_dmlpath))
        { md $kc46Highlights_dmlpath}
        Copy-Item -Path $dmlFile.FullName -Destination $kc46Highlights_dmlpath -Verbose
    }

    DistributeData -ddn $ddn -env $env

    #Remove-DataModulesFromCSDBPerDML -DDNFolder "C:\KC46 Staging\Production\Archives\Source\UnpackHere\$ddn" -manualsBasePath "C:\KC46 Staging\$env\Manuals" -actuallyDeleteTheFiles $true		

}
Function Remove-DataModulesFromCSDBPerDML
{
    param([string] $DDNFolder , [string] $manualsBasePath ,[bool] $actuallyDeleteTheFiles = $false)

     
    $dmlFileFNArray = @()
    $dmlFileArray = @()

    $ddnPolderPath = "$DDNFolder\DML*.XML"
    $logFolder = "$DDNFolder\DMLProcessLog"
    $logPath = "$logFolder\DMLProcessLog.txt"
    $dmlFiles = gci -Recurse -Path $ddnPolderPath
 
    if(!(Test-Path -Path $logFolder))
    {
        md $logFolder 
    }
    else
    {       
        Remove-Item -Path $logPath -Recurse -Force -verbose -ErrorAction SilentlyContinue
    }


    # Get the path to the correct DML files
    foreach ($dmlFile in $dmlFiles)
    {
        if($dmlFile.FullName.Contains("SDLLIVE"))
        {
            $dmlFile.FullName
            if(! $dmlFile.Name.Contains("AAA0B")) # Ignore Baggage Cargo Loading Manual data
            {
                $dmlFileFNArray += $dmlFile.FullName
                $dmlFileArray += $dmlFile.Name
            }        
        }
    }

    $dmlFileArray
    $dmlFileFNArray

    # For each DML, process the 'deleted files only from the CSDB.
    <#
        Note: if a dm in the DML is in our ValVer DB as "Verified, it has to be set back to Tab top and not verified in the DB
    #>
    $newDMs =@()
    $changedDMs = @()
    $deletedDMs = @()

    foreach ($dmlFileFName in $dmlFileFNArray)
    {
        $dmlDoc = New-Object System.Xml.XmlDocument
        $dmlDoc.Load($dmlFileFName)
        $dmRefDeleted = $dmlDoc.SelectNodes("/dml/dmlContent/dmEntry[@dmEntryType=`"d`"]")
        if($dmRefDeleted.Count -eq 0)
        {
            $msg = "The DML file, $dmlFileFName, did not contain any entries requiring deletion.  No action taken"
            Submit-LogEntry -fullLogPath $logPath -EventID 60006 -evtType "Information" -caller "KC46DataDistribution.psm1::Remove-DataModulesFromCSDBPerDML" -Message $msg
        }
        else
        {
            foreach ($dmEntry in $dmRefDeleted)
            {
                $ddnEl = $dmEntry.dmRef
                $ddn = Get-FilenameFromDMRef -dmRef $ddnEl -filePref "DMC"        
                $DMDocType = Get-DocTypeFromDMC -dc $ddn
                $logMsg = "$DMDocType | $manualsBasePath\$DMDocType\S1000D\SDLLIVE\$ddn.XML"
                Submit-LogEntry -fullLogPath $logPath -EventID 60007 -evtType "Information" -caller "KC46DataDistribution.psm1::Remove-DataModulesFromCSDBPerDML" -Message $logMsg
            }
        }
        foreach ($dmEntry in $dmRefDeleted)
        {
            $ddnEl = $dmEntry.dmRef
            $ddn = Get-FilenameFromDMRef -dmRef $ddnEl -filePref "DMC"        
            $DMDocType = Get-DocTypeFromDMC -dc $ddn
            
            $pathToTheDMCToDelete = "$manualsBasePath\$DMDocType\S1000D\SDLLIVE\$ddn.XML"
            $pathToTheDMCToDelete
            if((Test-Path -Path $pathToTheDMCToDelete) -and ($actuallyDeleteTheFiles))
            {
                $error.Clear()
                try
                {
                     Remove-Item -Verbose -Force -Path $pathToTheDMCToDelete -ErrorAction Stop
                     Submit-LogEntry -fullLogPath $logPath -EventID 60008 -evtType "Information" -caller "KC46DataDistribution.psm1::Remove-DataModulesFromCSDBPerDML" -Message "File successfully removed from CSDB : $pathToTheDMCToDelete"
                }
                catch 
                {
                    "Logging Failure: "
                    Submit-LogEntry -fullLogPath $logPath -EventID 60009 -evtType "Error" -caller "KC46DataDistribution.psm1::Remove-DataModulesFromCSDBPerDML" -Message "Failed to delete : $pathToTheDMCToDelete" 
                }
 
                finally
                {
                    Write-Host "Failure Logged while removing $pathToTheDMCToDelete"
                }
            }
            elseif((Test-Path -Path $pathToTheDMCToDelete) -and ($actuallyDeleteTheFiles -eq $false))
            { 
                Remove-Item -Verbose -Force -Path $pathToTheDMCToDelete -WhatIf
            }
            else
            {
                "This path was not accessible: " + $pathToTheDMCToDelete
                Submit-LogEntry -fullLogPath $logPath -EventID 60010 -evtType "Error" -caller "KC46DataDistribution.psm1::Remove-DataModulesFromCSDBPerDML" -Message "This path was not accessible: $pathToTheDMCToDelete"
            }
        }    
    }
}
Function Clear-PMC
{
    Param([string] $manual, [string] $env , [string] $source)
    $sourcePMCount = gci -Path "$source\PMC*.XML"
    if($sourcePMCount.Count -gt 0)
    {
        "Source contains a new PMC....`n$source"
        $pmPath = "C:\KC46 Staging\$env\Manuals\$manual\S1000D"
        $pmPath_SDL = "C:\KC46 Staging\$env\Manuals\$manual\S1000D\SDLLIVE"
        $pmCount = gci -Path "$pmPath\PMC*.XML"
        $pmCount_SDL = gci -Path "$pmPath_SDL\PMC*.XML"

        if($pmCount.Count -gt 0)
        {
            Remove-Item -Path "$pmPath\PMC*.XML" -Verbose
        }
        if($pmCount_SDL.Count -gt 0)
        {
            Remove-Item -Path "$pmPath_SDL\PMC*.XML" -Verbose
        }
    }
}
Function Get-DM
{
Param([string] $pathToDM)
    $dm = New-Object System.Xml.XmlDocument
    $dm.Load($pathToDM)
    return $dm
}
Function Update-CSDB 
{
    Param([string] $ddn, [string] $env , [string] $manual , [string] $dmSource)

    if(Test-Path -Path $dmSource)
    {   
        $destination1 = "C:\KC46 Staging\$env\Manuals\$manual\S1000D"
        $destination2 = "C:\KC46 Staging\$env\Manuals\$manual\S1000D\SDLLIVE"
        
        $destination3 = "C:\KC46 Staging\$env\FullCSDB\$manual\S1000D"
        $destination4 = "C:\KC46 Staging\$env\FullCSDB\$manual\S1000D\SDLLIVE"
        if(!(Test-Path -Path $destination1))
            { md $destination1 }
        if(!(Test-Path -Path $destination2))
            { md $destination1 }
        if(!(Test-Path -Path $destination3))
            { md $destination3 }
        if(!(Test-Path -Path $destination4))
            { md $destination4 }


        if($manual -eq "IPB")
        {
            $srcFiles = gci -path "$dmSource\*MC*.XML"
            foreach ($srcFile in $srcFiles)
            {
                $sName = $srcFile.Name
                # Get the basename of each file
                if($srcFile.Name.Contains("_"))
                {
                    $baseArray = $srcFile.Name.Split("_")
                    $baseName = $baseArray[0]
                }
                elseif(! $srcFile.Name.Contains("_"))
                {
                    $baseArray = $srcFile.Name.Split(".")
                    $baseName = $baseArray[0]
                }

                # Insure the new file is in the FullCSDB                
                Copy-Item  -Path $srcFile.FullName -Destination $destination3  -Verbose -ErrorAction Inquire
                if(!(Test-Path -Path "$destination3\$sName"))
                { "Error copying $srcFile.FullName to $destination3"}
                
                Copy-Item  -Path $srcFile.FullName -Destination $destination4  -Verbose -ErrorAction Inquire
                if(!(Test-Path -Path "$destination4\$sName"))
                { "Error copying $srcFile.FullName to $destination4"}


                Remove-Item -Path "$destination1\$baseName`*"  -Force -Verbose  -ErrorAction Inquire
                Remove-Item -Path "$destination2\$baseName`*"  -Force -Verbose -ErrorAction Inquire

                Copy-Item  -Path $srcFile.FullName -Destination $destination1 -Verbose  -ErrorAction Inquire
                Copy-Item  -Path $srcFile.FullName -Destination $destination2 -Verbose -ErrorAction Inquire
            }
        }
        else
        {  
            $srcFiles1 = gci -path "$dmSource\*MC*.XML"
            
            foreach ($srcFile in $srcFiles1)
            {
                $sName = $srcFile.Name
                # Insure the new file is in the FullCSDB
                Copy-Item  -Path $srcFile.FullName -Destination $destination3 -Verbose -ErrorAction Inquire -Force
                
                if(!(Test-Path -Path "$destination3\$sName"))
                { "Error copying $srcFile.FullName to $destination3"}
            } 

            #Load the SDLLIVE folder only
            $srcFiles2 = gci -path "$dmSource\SDLLIVE\*MC*.XML"         
            foreach ($srcFile in $srcFiles2)
            {
                $sName = $srcFile.Name
                # Insure the new file is in the FullCSDB
                Copy-Item  -Path $srcFile.FullName -Destination $destination4 -Verbose -ErrorAction Inquire
                
                if(!(Test-Path -Path "$destination4\$sName"))
                { "Error copying $srcFile.FullName to $destination4"}

                # Get the basename of each file
                $baseArray = $srcFile.Name.Split("_")
                $baseName = $baseArray[0]
                $fileset2 = gci -Path "$destination2\$baseName`*"
                foreach($file in $fileset2)
                {
                    Remove-Item -Path "$destination2\$baseName`*"  -ErrorAction Inquire
                }
                
                Copy-Item  -Path $srcFile.FullName -Destination $destination2 -Verbose -Force  -ErrorAction Inquire
            } 
        }
    }
}
Function Reset-Filename
{
Param([string] $pathToDM,[string] $newName)
    Rename-Item -Verbose -Force -Path $pathToDM -NewName $newName                
}
Function DistributeData
{
Param(
	[string] $ddn, 
    [string] $env)
        "Calling IPB Workflow .... $ddn - $env"
        if($ddn.Contains("1KC46-81205-81205-"))
        {
            Update-CSDB -ddn $ddn -env $env -manual "IPB" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn"
        }
        else
        {
            "Calling ARD Workflow .... $ddn - $env"
		    Update-CSDB -ddn $ddn -env $env -manual "ARD" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn\ARD-KC\ARD-KC_DataModules"
        
            "Calling AMM Workflow .... $ddn - $env"
            Update-CSDB -ddn $ddn -env $env -manual "AMM" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn\AMM-KC\AMM-KC_DataModules" 

            "Calling COMMON Workflow .... $ddn - $env"
            Update-CSDB -ddn $ddn -env $env -manual "KC46" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn\COMMON\COMMON_DataModules" 
        
            "Calling FIM Workflow .... $ddn - $env"
            Update-CSDB -ddn $ddn -env $env -manual "FIM" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn\FIM-KC\FIM-KC_DataModules"

            "Calling TC Workflow .... $ddn - $env"
            Update-CSDB -ddn $ddn -env $env -manual "TC" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn\TC-KC\TC-KC_DataModules" 
        
            "Calling NDT Workflow .... $ddn - $env"
            Update-CSDB -ddn $ddn -env $env -manual "NDT" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn\NDT\NDT_DataModules"
        
            "Calling SRM Workflow .... $ddn - $env"
            Update-CSDB -ddn $ddn -env $env -manual "SRM" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn\SRM\SRM_DataModules"
        
            "Calling SSM Workflow .... $ddn - $env"
            Update-CSDB -ddn $ddn -env $env -manual "SSM" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn\SSM-KC\SSM-KC_DataModules"
        
            "Calling SWPM Workflow .... $ddn - $env"
            Update-CSDB -ddn $ddn -env $env -manual "SWPM" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn\SWPM\SWPM_DataModules"
        
            "Calling WDM Workflow .... $ddn - $env"
            Update-CSDB -ddn $ddn -env $env -manual "WDM" -dmSource "C:\KC46 Staging\$env\Archives\Source\UnpackHere\$ddn\WDM-KC\WDM-KC_DataModules"
        }
    
}
Function Reset-AllCSDBFolders
{
    Param([string[]] $manuals , [string] $env)
    $manuals | Foreach-Object -Parallel {
                  
        "Removing DATA MODULES from C:\KC46 Staging\$env\Manuals\$_\S1000D"
        Remove-Item "C:\KC46 Staging\$env\Manuals\$_\S1000D\*.*" -Recurse -Force
        Remove-Item "C:\KC46 Staging\$env\Manuals\$_\S1000D\SDLLIVE\*.*" -Recurse -Force
        Remove-Item "C:\KC46 Staging\$env\Manuals\$_\ILLUSTRATIONS" -Recurse -Force          
    }
    $manuals | Foreach-Object -Parallel {
        if(!(Test-path -Path "C:\KC46 Staging\$env\Manuals\$_\S1000D\SDLLIVE"))
        {
            MD "C:\KC46 Staging\$env\Manuals\$_\S1000D\SDLLIVE"
        }
        if(!(Test-path -Path "C:\KC46 Staging\$env\Manuals\$_\ILLUSTRATIONS\ILLUSTRATIONS"))
        {
           MD "C:\KC46 Staging\$env\Manuals\$_\ILLUSTRATIONS\ILLUSTRATIONS"
        }
    }
}