cls
$Pub = ""
$sd = $null
$ed = $null
$sd = Get-Date
$ErrorActionPreference = "SilentlyContinue"
$error.Clear()
$a = (Get-Host).PrivateData
$a.WarningBackgroundColor = "Yellow"
$a.WarningForegroundColor = "Black"
$rn = "<br/>"
$a.ErrorBackgroundColor = "red"
$a.ErrorForegroundColor = "White"
# *****************************************************************************************************
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\KC46 Staging\Scripts\Modules;"
Import-Module -Name "KC46Common" -Verbose -Force

Release 8
$illPath =  "\\a5778954\Releases\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\AMM\Illustrations\Illustrations"
$dataPath = "\\a5778954\Releases\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\AMM\S1000D\SDLLIVE"
$dmlpath = "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set\Release 8 DMLs\DML-1KC46-AAA0A-P-2017-00003_001-00_SX-US.XML" 
$outFolder = "\\a5778954\BCDR\IPR\100 Percent IPR - December 2017\KC46_100_IPR_AMM_12_2017_Rel8\AMM"

<# Release 7

    $illPath = "\\a5778954\Releases\2017-09-18-14-39-31 - Non CDRL Sept 2017 - Release 7\CSDB\DVD\AMM\Illustrations\Illustrations"
    $dataPath = "\\a5778954\Releases\2017-09-18-14-39-31 - Non CDRL Sept 2017 - Release 7\CSDB\DVD\AMM\S1000D\SDLLIVE"
    $dmlpath = "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set\Release 7 DMLs\DML-1KC46-AAA0A-P-2017-00002_001-00_SX-US.xml"
#>
<# Release 6 
$illPath =  "\\a5778954\Releases\2017-06-06-07-18-23 - Non CDRL June 2017 - Release 6\CSDB\DVD\AMM\Illustrations\Illustrations"
$dataPath = "\\a5778954\releases\2017-06-06-07-18-23 - Non CDRL June 2017 - Release 6\CSDB\DVD\AMM\S1000D\SDLLIVE"
$dmlpath = "C:\KC46 Staging\production\manuals\kc46\s1000d\frontmatter\Release DML Set\Release 6 DMLs\DML-1KC46-AAA0A-P-2017-00001_001-00_SX-US.XML" 
$outFolder = "\\a5778954\BCDR\IPR\100 Percent IPR - December 2017\KC46_100_IPR_AMM_12_2017_Rel6\AMM"

#>
$dmlXml = New-Object System.Xml.XmlDocument
$dmlXml.Load($dmlpath )
$dmlEntries = $dmlXml.SelectNodes("/dml/dmlContent//dmEntry[@dmEntryType=`"n`"]")
$dmlEntries.Count
$AMMarray = @()
foreach ($dmlEntry in $dmlEntries)
{
    $AMMarray += (Get-FilenameFromDMRef -dmRef $dmlEntry.dmRef -filePref "DMC").Replace("DMC-","")
}


#     NO AMM IN THIS LIST....IT WAS PROCESSED IN A DIFFERENT MANNER ABOVE
[string[]] $ManualS   = @("ACS","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WDM","WUC")
[string[]] $ManualS   = @("AMM")
foreach ($Manual in $ManualS)
{
    #$outFolder = "\\a5778954\BCDR\IPR\100 Percent IPR - December 2017\KC46_100_IPR_" + $Manual + "_12_2017_Rel8\$Manual"
    
    #$illPath =  "\\a5778954\Releases\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\$Manual\Illustrations\Illustrations"
    #$dataPath = "\\a5778954\Releases\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\$Manual\S1000D\SDLLIVE"

    $pmcName = gci -Path $dataPath -Filter pmc*.xml | Sort-Object -Descending | select -First 1
    
    # Get the PMC from the manual - remember to get the latest version based on the issue # within the file name
    $pmXml = New-Object System.Xml.XmlDocument
    $pmXmlOut = New-Object System.Xml.XmlDocument
    $pmXml.Load($pmcName[0].FullName)

    # Clone the PMC
    $pmXmlOut=$pmXml
    #$outFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs\BCDRDataPrep\$Manual\"
    
    # Establish a new name for the new abbreviated verion of the PMC
    $pmcFileNameOut = "$outFolder\S1000D\" + $pmcName[0].Name
    
	# If the $destinationPath does not exist, it will be created    
	if(!(Test-Path -Path $outFolder))
	{ 
        md "$outFolder\S1000D"
        md "$outFolder\Illustrations"
    }
    else
    {
        Remove-Item -Path $outFolder -Force -ErrorAction Inquire -Recurse -Verbose
        md "$outFolder\S1000D"
        md "$outFolder\Illustrations"
    }
    
    # Get a list of all the dmc references in the PMC
    # /pm/content/pmEntry/pmEntry/pmEntry/dmRef/dmRefIdent/dmCode
    $dmRefs = $pmXmlOut.SelectNodes("/pm/content/*//dmRef")
    $dmRefArray = @()
    
    if($dmRefs.Count -gt 0)
     {
        foreach ($dmRef in $dmRefs)
        {
            $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
            $dmc = $dmc.Replace("DMC-", "")
            $dmRefArray += $dmc
        }        
     }

    # Note: Items in the whiTelist are the only dmcs that should remain in the output PMC
    
    $removalCounter = 0
    $remainCounter = 0
    $remainingdmrefs = @()
    $dmRefCountBefore = $dmRefs.Count
    foreach ($dmRef in $dmRefs)
    {
        $leaveDMCInPMC = $false
        $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
        $dmc = $dmc.Replace("DMC-", "")
        
        if($Manual -eq "AMM")
        {
            if($AMMArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "ACS")
        {                
            if($ACSArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "ARD")
        {                
            if($ARDArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "FIM")
        {                
            if($FIMArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
#[string[]] $ManualS   = @("ACS","AMM","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WDM","WUC")
        elseif($Manual -eq "IPB")
        {                
            if($IPBArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "LOAPS")
        {                
            if($LOAPSArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
#[string[]] $ManualS   = @("ACS","AMM","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WDM","WUC")
        elseif($Manual -eq "NDT")
        {                
            if($NDTArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }

        elseif($Manual -eq "SIMR")
        {                
            if($SIMRArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "SPCC")
        {                
            if($SPCCArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "SSM")
        {                
            if($SSMArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "SWPM")
        {                
            if($SWPMArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "TC")
        {                
            if($TCArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
#[string[]] $ManualS   = @("ACS","AMM","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WDM","WUC")
        elseif($Manual -eq "WDM")
        {                
            if($WDMArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "WUC")
        {                
            if($WUCArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }

        if($leaveDMCInPMC -eq $true)
        {
            $remainCounter ++
            $remainingdmrefs += $dmc
        }
        else
        {
            " Remove the dmref"
            $removalCounter ++
            $pn = $dmRef.ParentNode
            $null = $pn.RemoveChild($dmRef)
        }              
    }

    $remainCounter
    $removalCounter
    "$Manual - DMRef to start`t:" + $dmRefCountBefore
    "$Manual - DMRef counter removed`t:" + $removalCounter
    "$Manual - We should now have`t: "  + ($dmRefCountBefore - $removalCounter)
    "$Manual - Actual remainiing `t: " + ($pmXmlOut.SelectNodes("/pm/content/*//dmRef")).Count
    # After clearing all the dmc that don't pertain, remove whole trees from the pmc have a pmEntry elements with 0 dmrefs      
    $pmCHNodes  = $pmXmlOut.SelectNodes("/pm/content/pmEntry")
    $pmSENodes  = $pmXmlOut.SelectNodes("/pm/content/pmEntry/pmEntry")
    $pmSUNodes  = $pmXmlOut.SelectNodes("/pm/content/pmEntry/pmEntry/pmEntry")
    $pmCHNodes.Count
            
    foreach ($pmSUNode in $pmSUNodes)
    {
        if($pmSUNode.SelectNodes("./*[descendant-or-self::dmRef]").Count -eq 0)    
        {
            # Remove this entire chapter from the new PMC
            $pmSUNode.pmEntryTitle
            $pn = $pmSUNode.ParentNode                
            $null = $pn.RemoveChild($pmSUNode)
        }
    }

    foreach ($pmSENode in $pmSENodes)
    {
        if($pmSENode.SelectNodes("./*[descendant-or-self::dmRef]").Count -eq 0)    
        {
            # Remove this entire chapter from the new PMC
            $pmSENode.pmEntryTitle
            $pn = $pmSENode.ParentNode     
            $null = $pn.RemoveChild($pmSENode)   
        }
    }
       
    foreach ($pmCHNode in $pmCHNodes)
    {
        if($pmCHNode.SelectNodes("./*[descendant-or-self::dmRef]").Count -eq 0)    
        {
            # Remove this entire chapter from the new PMC
            $pmCHNode.pmEntryTitle
            $pn = $pmCHNode.ParentNode
            $null = $pn.RemoveChild($pmCHNode)
        }
    }

    $pmXmlOut.Save($pmcFileNameOut)
    
    $finalSetofDMs = @()
    "Actual remainiing after PMC cleanup `t: " + ($pmXmlOut.SelectNodes("/pm/content/pmEntry/pmEntry/pmEntry/dmRef")).Count
    foreach ($dmRef in $pmXmlOut.SelectNodes("/pm/content/*//dmRef"))
    {
        $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
        $dmc = $dmc.Replace("DMC-", "")
        $finalSetofDMs += $dmc
    }
    
    Compare-Object -ReferenceObject $remainingdmrefs -DifferenceObject $finalSetofDMs
[string[]] $ManualS   = @("ACS","AMM","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WDM","WUC")    
    if($Manual -eq "ACS")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($ACSArray |Sort-Object) }
    elseif($Manual -eq "AMM")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($AMMArray |Sort-Object) }
    elseif($Manual -eq "ARD")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($ARDArray |Sort-Object) }
    elseif($Manual -eq "FIM")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($FIMArray |Sort-Object) }
    elseif($Manual -eq "IPB")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($IPBArray |Sort-Object) }
    elseif($Manual -eq "LOAPS")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($LOAPSArray |Sort-Object) }
    elseif($Manual -eq "NDT")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($NDTArray |Sort-Object) }
    elseif($Manual -eq "SIMR")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($SIMRArray |Sort-Object) }
     elseif($Manual -eq "SPCC")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($SPCCArray |Sort-Object) }
    elseif($Manual -eq "SSM")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($SSMArray |Sort-Object) }
    elseif($Manual -eq "SWPM")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($SWPMArray |Sort-Object) }
    elseif($Manual -eq "TC")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($TCArray |Sort-Object) }
    elseif($Manual -eq "WDM")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($WDMArray |Sort-Object) }
    elseif($Manual -eq "WUC")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($WUCArray |Sort-Object) }

                   
    $pmXmlOut.Save($pmcFileNameOut)

    foreach ($finalSetofDM in $finalSetofDMs)
    {
        # copy the correct files to the folder
        $fn = $finalSetofDM
        $fn = "DMC-$finalSetofDM" + "*.XML"
                
        $FFN = "$dataPath\$fn"

        $dm = New-Object System.Xml.XmlDocument
        $dm.Load((gci -Path $FFN)[0].FullName)
        $FFN
        $graphics = $dm.SelectNodes("//graphic")
        $s1000dDestination = "$outFolder\S1000D"
        $illustrationsdDestination = "$outFolder\Illustrations"

        foreach($graphic in $graphics)
        {
            $icnName = $graphic.infoEntityIdent + ".cgm"
            $icnSourcePath = "$illPath\$icnName"
            Copy-Item -Path $icnSourcePath -Destination $illustrationsdDestination -Verbose -ErrorAction Inquire
        }
        Copy-Item -Path $FFN -Destination $s1000dDestination -Verbose -ErrorAction Inquire
    }

    $x | Out-File "$outFolder\$Manual - Discrepency List.txt"
    $finalSetofDMs | Out-File "$outFolder\$Manual - List of DMs for BCDR Review.txt"
}

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$ManualList"
"Total Days to complete:`t" + $x.Days
"Total Hours to complete:`t" + $x.Hours
"Total Minutes to complete:`t" + $x.Minutes
"Total Seconds to complete:`t" + $x.Seconds
"Process completed"