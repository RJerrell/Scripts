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
$localhash = @()

$AMMArray = @()
$FIMArray = @()
$IPBArray = @()
$NDTArray = @()
$ARDArray = @()
$SSMArray = @()
#$path2WhiteList = "\\nw\newgen-tanker\Spt_TrngSys\SptSys\Tech Pubs\IPR - InProcessReview\100% IPR DMCs\100 Percent IPR DMCs Summary by Manual - BCDR.csv"
$path2WhiteList = "\\nw\newgen-tanker\Spt_TrngSys\SptSys\Tech Pubs\IPR - InProcessReview\100% IPR DMCs\100 Percent IPR DMCs Summary by Manual - BCDR.csv"
$csdbPath = "F:\KC46 Staging\Production\Manuals\PUB\S1000D\SDLLIVE\PMC*.*"
$illistrationsPath = "F:\KC46 Staging\Production\Manuals\PUB\Illustrations\Illustrations"
# Load the white list of all manuals

$localhash = Import-Csv -Path $path2WhiteList -Delimiter "," -Header 'AMM','FIM','IPB','NDT','ARD', 'SSM' -Verbose

foreach($entry in $localhash)
{
    if($entry.AMM.Length -gt 0)
    {
        $AMMArray += $entry.AMM
    }

    if($entry.FIM.Length -gt 0)
    {
        $FIMArray += $entry.FIM
    }

    if($entry.IPB.Length -gt 0)
    {
        $IPBArray += $entry.IPB
    }

    if($entry.NDT.Length -gt 0)
    {
        $NDTArray += $entry.NDT
    }

    if($entry.ARD.Length -gt 0)
    {
        $ARDArray += $entry.ARD
    } 
    if($entry.SSM.Length -gt 0)
    {
        $SSMArray += $entry.SSM
    } 
}
$ManualList = ( @('AMM','ARD','FIM','IPB','NDT') | Sort-Object)
# $ManualList = ( @('AMM') | Sort-Object)
foreach ($Manual in $ManualList)
{
    $csdbPath2 = $csdbPath.Replace("PUB", $Manual)

    $pmcName = gci -Path $csdbPath2 | Sort-Object -Descending | select -First 1
    
    # Get the PMC fror the manual - remember to get the latest version based on the issue # within the file name
    $pmXml = New-Object System.Xml.XmlDocument
    $pmXmlOut = New-Object System.Xml.XmlDocument
    $pmXml.Load($pmcName[0].FullName)

    # Clone the PMC
    $pmXmlOut=$pmXml
    #$outFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs\BCDRDataPrep\$Manual\"
    $outFolder = "D:\Shared\BCDR\IPR\100 Percent IPR - June 2017\KC46_100_IPR_$Manual" + "_06_2017\$Manual"
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
        Remove-Item -Path "$outFolder\*.*" -Force -ErrorAction Inquire -Recurse -Verbose
        #md $outFolder
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
        elseif($Manual -eq "FIM")
        {                
            if($FIMArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "IPB")
        {                
            if($IPBArray.Contains($dmc) )
            {
                " Exists"
                $leaveDMCInPMC = $true
            }
        }
        elseif($Manual -eq "NDT")
        {                
            if($NDTArray.Contains($dmc) )
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
        elseif($Manual -eq "SSM")
        {                
            if($SSMArray.Contains($dmc) )
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
    
    if($Manual -eq "AMM")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($AMMArray |Sort-Object) }
    elseif($Manual -eq "ARD")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($ARDArray |Sort-Object) }
    elseif($Manual -eq "FIM")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($FIMArray |Sort-Object) }
    elseif($Manual -eq "IPB")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($IPBArray |Sort-Object) }
    elseif($Manual -eq "NDT")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($NDTArray |Sort-Object) }
    elseif($Manual -eq "SSM")
    {    $x = Compare-Object -ReferenceObject ($finalSetofDMs |Sort-Object) -DifferenceObject ($SSMArray |Sort-Object) }
        
    $pmXmlOut.Save($pmcFileNameOut)

    foreach ($finalSetofDM in $finalSetofDMs)
    {
        # copy the correct files to the folder
        $fn = $finalSetofDM
        $fn = "DMC-$finalSetofDM" + "*.XML"
        # $csdbPath = "F:\KC46 Staging\Production\Manuals\PUB\S1000D\SDLLIVE\PMC*.*"
        $FFN = $csdbPath.Replace("PMC*.*", "").Replace("PUB", $Manual) + $fn
        $dm = New-Object System.Xml.XmlDocument
        $dm.Load((gci -Path $FFN)[0].FullName)
        $FFN
        $graphics = $dm.SelectNodes("//graphic")
        $s1000dDestination = "$outFolder\S1000D"
        $illustrationsdDestination = "$outFolder\Illustrations"

        foreach($graphic in $graphics)
        {
            $icnName = $graphic.infoEntityIdent + ".cgm"
            $icnSourcePath = $illistrationsPath.Replace("PUB", $Manual)
            $icnSourcePath = "$icnSourcePath\$icnName"
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