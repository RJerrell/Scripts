#region Worker functions
function Set-PowerShellTitle
{
Param([string[]] $PubList , [string] $commonRoot, [string] $environment)
    $powerShellWindowTitle = ""
    foreach($m in  $PubList)
    {
	    $powerShellWindowTitle = $powerShellWindowTitle + "," + $m
    }
    $Host.UI.RawUI.WindowTitle = "$commonRoot`t: $environment - $powerShellWindowTitle"
}
function Add-ICN_Numbers_ToGraphics
{
Param([string] $icnBatchFilePath , [string[]] $icnBatchFilePath_args , [string] $figuresWithICNEmbedded)
    # Process all the figures
    "ICN Batch call syntax:`t$icnBatchFilePath $icnBatchFilePath_args"
    $argList = $icnBatchFilePath_args[0] + " " + $icnBatchFilePath_args[1]
    $batfile = [diagnostics.process]::Start($icnBatchFilePath , $argList)
    $batfile.WaitForExit()
}

function Optimize-WDMPMC
{
    Param([string] $pmcName, [string] $pmcShortName)
        $pmEntry = [xml] @"
    <pmEntry>
        <pmEntryTitle>abc</pmEntryTitle>
    </pmEntry>
"@
    $xmlPMC = New-Object System.Xml.XmlDocument
    $xmlPMC.Load($pmcName)
    if( ! ($xmlPMC.OuterXml.Contains("...")))
                                                                                                                                                                                                                                                                {
    $lists = @("91-00 - EQUIPMENT LIST","91-04 - DISCONNECT BRACKET LIST","91-21 - SPARE WIRE LIST","91-21 - HOOKUP LIST")
    $eqListPrfixes = @("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","Y","Z")
    $dListPrfixes = @("DA","DB","DC","DD","DE","DF","DG","DH","DI","DJ","DK","DL","DM","DN","DO","DP","DQ","DR","DS","DT","DU","DV","DW","DY","DZ","D1","D2","D3","D4","D5","D6","D7","D8","D9","D0")
    foreach ($list in $lists)
    {
        $list
        $parentPM = $xmlPMC.SelectSingleNode("/pm/content/pmEntry/pmEntry[pmEntryTitle = `"$list`"]")
        foreach ($eqListPrfix in $eqListPrfixes)
        {   if($eqListPrfix -eq "D")
            {
                foreach($dListPrfix in $dListPrfixes)
                {
                    
                    $dmRefs = $parentPM.dmRef | ?{ $_.title -match "^$dListPrfix"}
                    if($dmRefs.ChildNodes.Count -gt 0)
                    {
                        $dListPrfix
                        $newPMEntry = $pmEntry.Clone()
                        $newPMEntry.pmEntry.pmEntryTitle = "$dListPrfix ..."
                        $newPMEntry.pmEntry.pmEntryTitle  
            
                        # Now copy all the DmRefs into this new node
                        foreach ($dmRef in $dmRefs)
                        {            
                            $nn = $newPMEntry.ImportNode($dmRef, $true)
                            $null = $newPMEntry.DocumentElement.AppendChild($nn)
                            # Now remove all the master Document DmRefs that we copied to the new element
                            $null = $dmRef.ParentNode.RemoveChild($dmRef)
                        }
                        $newNode = $xmlPMC.ImportNode($newPMEntry.DocumentElement,$true)
                        $null = $parentPM.AppendChild($newNode)
                    }
                }
            }
            else
            {
                $dmRefs = $parentPM.dmRef | ?{ $_.title -match "^$eqListPrfix"}        
                if($dmRefs.Count -gt 0)
                {
                    $newPMEntry = $pmEntry.Clone()
                    $newPMEntry.pmEntry.pmEntryTitle = "$eqListPrfix ..."   
                    $newPMEntry.pmEntry.pmEntryTitle  
            
                    # Now copy all the DmRefs into this new node
                    foreach ($dmRef in $dmRefs)
                    {            
                        $nn = $newPMEntry.ImportNode($dmRef, $true)
                        $null = $newPMEntry.DocumentElement.AppendChild($nn)
                    }

                    # Now remove all the master Document DmRefs that we copied to the new element
                    foreach ($dmRef in $dmRefs)
                    {            
                        $null = $dmRef.ParentNode.RemoveChild($dmRef)
                    }

                    $newNode = $xmlPMC.ImportNode($newPMEntry.DocumentElement,$true)
                    $null = $parentPM.AppendChild($newNode)
                }
            }
        }
    
    }
    Save-PrettyXML -FName $pmcShortName -xmlDoc $xmlPMC
    #$xmlPMC.Save($pmcShortName)
    }
    else
    {
        "This WDM Publication Module has already been Processed"
    } 
}

Function Clear-OlderIllustrations
{
    Param([string[]] $Publist, [string] $source_BaseLocation)
    $Publist | Foreach-Object -ThrottleLimit 5 -Parallel 
    {
    $Illustrations_Folder = "$source_BaseLocation\$_\ILLUSTRATIONS\ILLUSTRATIONS"
    $files = Get-ChildItem -Path $Illustrations_Folder -Exclude *.txt | Sort-Object -Property Name
    foreach ($file in $files)
    {
        # Inwork
        $fileName = $file.Name.Replace(".cgm", "")
        $fileParts = $fileName.Split("-")

        $rootName = $fileParts[0],$fileParts[1],$fileParts[2] -join "-"
        $p = $file.Directory.FullName + "\$rootname*"
        $fileset = Get-ChildItem -Path $p | Sort-Object -Property LastWriteTime -Descending
        if($fileset.Count -gt 1)
        {
           for ($i = 1; $i -lt $fileset.Count; $i++)
           { 
               Remove-Item -Path $fileset[$i].FullName -Verbose
           }
        }
    }
}

}

#endregion

#region Primary functions
function Set-DefaultFolders
{
    param([string] $dmDestination, [string[]] $pubDestination, [string] $fullLogPath , [string] $evtID , [bool] $actionReq = $false)
    $msg = "Initializing the workspace...."          
    Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:496" -strActivity $msg -stage "EnvironmentPreparation" -actionReq $actionReq
    if(Test-Path -Path $dmDestination)
    {
        Remove-Item -Path "$dmDestination\*.*" -Force -Recurse -Verbose
        
    }
    else
    {
        md $dmDestination
    }

    if($pubDestination.Length -gt 0)
    {
        foreach($pdestination in $pubDestination)
        { 
            # A bug in the Remove-Item commandlet requires us to repeat it to insure the folders are clear and ready before we start processing.
            $allFiles = Get-ChildItem -Path $pdestination -Recurse -File
            foreach ($File in $allFiles)
            {
                if($File.Name.ToLower() -ne "properties.xml")
                {
                    Remove-Item -Path $File.FullName -Exclude properties.xml -Force -Recurse -Verbose
                }
            }

            sleep -Seconds 2
            if(!(Test-Path -Path "$pdestination\figures" ))
            {
                md "$pdestination\figures"
            }
        }
    }
            
}

function Reset-WIETM
{
    Param( [string] $Path , [string] $CSVPath, [string]  $releaseNumber )
    $output = ""
    # this is where the XML sample file was saved: $Path = "$env:temp\inventory.xml" # load it into an XML object: 
    $xml = [XML] (Get-Content -Path $Path)
    # note: if your XML is malformed, you will get an exception here # always make sure your node names do not contain spaces  
    # simply traverse the nodes and select the information you want: 
    $bookName = ""
    $bookTitle = ""
    $bookDescription = ""
    $PubIssue = ""
    $PubDate = ""
    $rows = @()
    $rows += "Manual`tManual Title`tManual Description`tPub. Issue`tPub. Number`tPub. Date "

    $books = $Xml.SelectNodes("//book")

    $d = (Get-Date)

    Foreach( $b in $books)
    {
        $bookName = $b.name
        Foreach($c in $b.configitem)
        { 
            switch ($c.name)
            {
            "bookTitle"
                {
                    $bookTitle = $c.value
                }
            "bookDescription"
                {
                    $bookDescription = $c.value
                }
            "PubIssue.value"
                {
                    $c.value = $d.Year.ToString() + "-" + $d.Month.ToString().PadLeft(2,"0")
                    $PubIssue = $c.value
                }
            "PubNumber.value"
                {
                     $c.value= $releaseNumber
                    $PubNumber = $releaseNumber
                }
            "PubDate.value"
                {
                    $c.value = $d.DateTime.ToString()
                    $PubDate = $c.value
                }
            }
        }
        $rows += "$bookName`t$bookTitle`t$bookDescription`t$PubIssue`t$PubNumber`t$PubDate"
    }
    
    # Attrib -R $Path #Remove READ-Only flag
    Set-ItemProperty -Path $Path -Name IsReadOnly -Value $false -Force 
    Save-PrettyXML -FName $Path -xmlDoc $xml
    Set-Content -Path  $CSVPath -Value $rows
}

function Set-PMCodes
{
    Param([string] $masterPMCFullPath)
    #Attrib -r $masterPMCFullPath
    Set-ItemProperty -Path $masterPMCFullPath -Name IsReadOnly -Value $false -Force 
    # This is the Wrapper Publication Module
    $masterPM = [XML] (Get-Content -Path $masterPMCFullPath) 
    
    #For each pmEntry in the master publication module ...
    $pmes = $masterPM.SelectNodes("/pm/content/pmEntry/pmEntry")
    foreach ( $pme in $pmes)
    {
        $pm = $null
        $file = $null
        
        #Get the current Title in the master pm....
        $pme.pmEntryTitle

        #Lookup the entry in this switch and create an in-memory xml of that PM file
        switch ($pme.pmEntryTitle)
          {
                'Aircraft Battle Damage Repair (ABDR)' { 
                if(Test-Path -Path "$dmDestination\PMC-1KC46-81205-G*")
                    {
                        $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-G*"
                        $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Aircraft Cross Servicing (ACS)' { 
                if(Test-Path -Path "$dmDestination\PMC-1KC46-81205-H*")
                {
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-H*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                break;
                }
                'Aircraft Maintenance (AMM)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-A*")
                {                     
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-A*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Aircraft Recovery Document (ARD)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-E*")
                {                     
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-E*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                 break;
                }
                'Fault Isolation (FIM)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-F*")
                {                     
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-F*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Illustrated Parts Catalog (IPC)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-P*")
                {                    
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-P*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Inspection and Maintenance Requirements (-6) Manual' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-J*")
                {                     
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-J*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'List of Applicable Publications (LOAP)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-M*")
                {                      
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-M*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Nondestructive Test (NDT)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-V*")
                {                     
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-V*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }                
                'Nondestructive Test Manaul Supplement(NDTS)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-N*")
                {                     
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-N*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Standard Wiring Practices Manual (SWPM)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-U*")
                {                     
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-U*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Structural Repair Manual (SRM)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-S*")
                {                     
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-S*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'System Schematic Manual (SSM)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-R*")
                {                     
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-R*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                   break;
                }
                'Task Cards (TC)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-T*")
                {                     
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-T*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Wiring Data Manual (WDM)' {
                if(Test-Path -Path  "$dmDestination\PMC-1KC46-81205-W0000*")
                {                    
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-W0000*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Work Unit Code Manual (-06)' {
                if(Test-Path -Path "$dmDestination\PMC-1KC46-81205-D*")
                {
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-D*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Common Lists' {
                if(Test-Path -Path "$dmDestination\PMC-1KC46-81205-Z*")
                {                    
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-Z*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }
                'Aircraft Structural Integrity Program (ASIP)' {
                if(Test-Path -Path "$dmDestination\PMC-1KC46-81205-C*")
                {                    
                    $file = Get-Item -Path "$dmDestination\PMC-1KC46-81205-C*"
                    $pm = [XML] (Get-Content -Path $file.FullName)
                    }
                    break;
                }                               
          }
          if($file -ne $null)
          {
            $pmCode = $pm.pm.identAndStatusSection.pmAddress.pmIdent.pmCode
            $pmCode          
            $pme.pmRef.pmRefIdent.pmCode.modelIdentCode  = $pm.pm.identAndStatusSection.pmAddress.pmIdent.pmCode.modelIdentCode
            $pme.pmRef.pmRefIdent.pmCode.pmIssuer  = $pm.pm.identAndStatusSection.pmAddress.pmIdent.pmCode.pmIssuer
            $pme.pmRef.pmRefIdent.pmCode.pmNumber  = $pm.pm.identAndStatusSection.pmAddress.pmIdent.pmCode.pmNumber
            $pme.pmRef.pmRefIdent.pmCode.pmVolume  = $pm.pm.identAndStatusSection.pmAddress.pmIdent.pmCode.pmVolume
            $pme.pmRef.pmRefAddressItems.issueDate.year = $pm.pm.identAndStatusSection.pmAddress.pmAddressItems.issueDate.year
            $pme.pmRef.pmRefAddressItems.issueDate.month = $pm.pm.identAndStatusSection.pmAddress.pmAddressItems.issueDate.month
            $pme.pmRef.pmRefAddressItems.issueDate.day = $pm.pm.identAndStatusSection.pmAddress.pmAddressItems.issueDate.day
          }
    }
    $masterPM.Save($masterPMCFullPath)
}
function Start-IETM_Build
{
    Param( [string] $propertiesFileLocation , [bool] $multiVolume = $false)
    # Build the IETM
    attrib.exe -R $propertiesFileLocation

    $arg1 =  '-build'
    $arg2 =  $propertiesFileLocation
    $arg3 =  ' -Xms '
    $arg4 =  ' 128m'
    $arg5 =  ' -Xmx '
    $arg6 =  ' 2048m'

    $exe = "C:\Program Files\XyEnterprise\LiveContent\LiveContentPublish.exe"

    if($multiVolume)
    {               
        
        # MAIN BOOK BUILD !
        & $exe $arg1 $arg2 $arg3 $arg4 $arg5 $arg6
        
        # WIRING BOOK BUILD !
        $arg2 = $propertiesFileLocation.Replace("KC46", "WIRING")
        & $exe $arg1 $arg2 $arg3 $arg4 $arg5 $arg6
    }
    else
    {
        "SINGLE BOOK BUILD!!!!"        
        & $exe $arg1 $arg2 $arg3 $arg4 $arg5 $arg6
    }
}
function Publish-IETM
{
Param([string] $configPath , [string] $publishIETMToThisLocation , [string] $autoplay_cdonlyLocation)
    # Publish the contents of the IETM build in preparation for packaging
    "Publishing the IETM to:`t$publishIETMToThisLocation"
    $trace =  CollectionPub.exe $configPath $publishIETMToThisLocation
    $trace
    Copy-Item -Path $autoplay_cdonlyLocation -Destination $publishIETMToThisLocation
}
Function Copy-IllustrationsToTmp
{
    Param([string[]] $Publist , [string] $source_BaseLocation, [string] $figuresWithICNEmbedded_tmp)
    #$figuresWithICNEmbedded_tmp
    
    foreach ( $p in $PubList)
    {    
        $sourceIll = $source_BaseLocation + "\"  + $p + "\ILLUSTRATIONS\ILLUSTRATIONS"
        if(!(Test-Path -Path $figuresWithICNEmbedded_tmp))
        {
            md $figuresWithICNEmbedded_tmp
        }

        if($p -ne "IPB")
        {
            Copy-FolderContents -dest $figuresWithICNEmbedded_tmp -source $sourceIll
        }
    }
}

Workflow Copy-DataModules
{
    Param([string[]] $Publist , [string] $source_BaseLocation, [string] $dmDestination , [string] $figuresWithICNEmbedded_tmp)
    $dmDestination
   
    foreach -parallel( $p in $PubList)
    {          
        $sourceDM = $source_BaseLocation + "\" + $p + "\S1000D\SDLLIVE"
        Copy-FolderContents -source $sourceDM  -dest $dmDestination            
     }
}

Workflow Get-NonS1000DData
{
    Param([string] $topLevelNonS1000DFolderFullPath , [string] $topLevelFlightDocumentsFolderFullPath , [string] $publishNonS1000DMaintenance , [string] $publishNonS1000DOperations)
    $evtID = 61400
    parallel
    {
        Copy-FolderContents -source $topLevelNonS1000DFolderFullPath -dest $publishNonS1000DMaintenance
        Copy-FolderContents -source $topLevelFlightDocumentsFolderFullPath -dest $publishNonS1000DOperations
        $msg = "Copying all the NonS1000D data from:`n$topLevelNonS1000DFolderFullPath `nTo`n$publishNonS1000DMaintenance`n`n`n All the Flight Manuals`nFrom$topLevelFlightDocumentsFolderFullPath`nTo$publishNonS1000DOperations"
        #Submit-LogEntry -fullLogPath $fullLogPath -Message $msg  -EventID $evtID -caller  "WorkFlow Start-KC46DataManagement line 364"
    }
}

WorkFlow Start-KC46DataManagement
{
Param( 
    [string[]] $Publist, 
    [string] $source_BaseLocation, 
    [string] $LiveContentDataFolder,
    [string] $figuresWithICNEmbedded,  
    [string] $dmDestination , 
    [string[]] $pubDestination,
    [string] $fullLogPath,
    [string] $siteUrl,
    [string] $listName
    )
    [bool] $actionReq = $false
         

        $m5 = "Copy the Illustrations to a temporary location :  "        
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:468" -strActivity "$m5 started"  -stage "Copy-IllustrationsToTmp" -actionReq $actionReq                    
        $dest = $pubDestination[0]
        if(! (Test-Path -Path "$dest\figures"))
        {
            md "$dest\figures"
        }
        
        $m1 = "Copy-DataModules data from authoring location `n($source_BaseLocation) `nto our build location ($dmDestination) "
        Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:498" -strActivity "$m1 started" -stage "DataModule Mgmt" -actionReq $actionReq

        parallel
        {         
            Copy-DataModules -Publist $Publist -source_BaseLocation $source_BaseLocation -dmDestination $dmDestination 
            Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:498" -strActivity "$m1 completed" -stage "DataModule Mgmt" -actionReq $actionReq
            foreach -parallel($pub in $Publist)
            {
                $m2 = "Copying all the Illustrations from $source_BaseLocation\$pub\Illustrations\Illustrations to $pubDestination[0]\figures - THEY ARE ALREADY ICN BRANDED " 
                Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:507" -strActivity "$m2 started" -stage "Illustrations Mgmt" -actionReq $actionReq
                
                $ipcDestination = $pubDestination[0] + "\figures"
                $destFolder = $pubDestination[0]
                Copy-FolderContents -source "$source_BaseLocation\$pub\Illustrations\Illustrations" -dest "$destFolder\figures"

                Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 DataManagement.psm1:507" -strActivity "$m2 completed" -stage "Illustrations Mgmt" -actionReq $actionReq
            }
        }        
        #Remove-Item -Path "$figuresWithICNEmbedded_tmp\*.*" -Force -Recurse        
    
}

#endregion

#region Exports
    #export-modulemember -function Start-*
    #export-modulemember -function Publish-IETM
    #export-modulemember -function Set-PMCodes
    #export-modulemember -function Set-DefaultFolders
    #export-modulemember -function Set-PowerShellTitle

    #export-modulemember -function Optimize-WDMPMC
    #export-modulemember -function Reset-*
    #export-modulemember -function Add-*
#endregion