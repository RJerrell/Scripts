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

#     NO AMM IN THIS LIST....IT WAS PROCESSED IN A DIFFERENT MANNER ABOVE
[string[]] $ManualS   = @("ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WDM","WUC")
[string[]] $ManualS   = @("ACS")
$pathSuffix = "_11_2017_Rel8"
$cred = Get-Credential -Credential -Verbose
foreach ($Manual in $ManualS)
{
    $outFolder = "\\a5778954\BCDR\IPR\100 Percent IPR - December 2017\KC46_100_IPR_$Manual$pathSuffix\$Manual"
    
    $illPath =  "\\a5778954\Releases\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\$Manual\Illustrations\Illustrations"

    $s1000dDestination = "$outFolder\S1000D"
    $illustrationsdDestination = "$outFolder\Illustrations"
    if( !(Test-Path -Path $s1000dDestination))
    {
        MD $s1000dDestination
    }
    else
    {
        Remove-Item -Path $s1000dDestination -Recurse -Verbose -Force
        MD $s1000dDestination
    }
    if( !(Test-Path -Path $illustrationsdDestination))
    {
        MD $illustrationsdDestination
    }
    else
    {
        Remove-Item -Path $illustrationsdDestination -Recurse -Verbose -Force
        MD $illustrationsdDestination
    }
    $dataPath = "\\a5778954\Releases\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\$Manual\S1000D\SDLLIVE"

    $pmcName = gci -Path $dataPath -Filter pmc*.xml | Sort-Object -Descending | select -First 1
    
    # Get the PMC from the manual - remember to get the latest version based on the issue # within the file name
    $pmXml = New-Object System.Xml.XmlDocument
    $pmXmlOut = New-Object System.Xml.XmlDocument
    $pmXml.Load($pmcName[0].FullName)

    # Clone the PMC
    $pmXmlOut = $pmXml
    #$outFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs\BCDRDataPrep\$Manual\"
    
    # Establish a new name for the new abbreviated verion of the PMC
    $pmcFileNameOut = "$outFolder\S1000D\" + $pmcName[0].Name
     
   
    # Get a list of all the dmc references in the PMC
    # /pm/content/pmEntry/pmEntry/pmEntry/dmRef/dmRefIdent/dmCode
    $dmRefs = $pmXmlOut.SelectNodes("/pm/content/*//dmRef")
    $dmNames = @()
    
    if($dmRefs.Count -gt 0)
     {
        foreach ($dmRef in $dmRefs)
        {
            $dmc = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
            $dmc = $dmc.Replace("DMC-", "")
            $dmNames += $dmc
        }        
     }

    foreach ($dmName in $dmNames)
    {
        # copy the correct files to the folder
        $fn = $dmName
        $fn = "DMC-$dmName" + "*.XML"
                
        $FFN = "$dataPath\$fn"

        $dm = New-Object System.Xml.XmlDocument
        $dm.Load((gci -Path $FFN)[0].FullName)
        $FFN
        $graphics = $dm.SelectNodes("//graphic")

        foreach($graphic in $graphics)
        {            
            $icnName = $graphic.infoEntityIdent + ".cgm"
            if($icnName.Contains("-KA"))
            {
             $illPath =  "\\a5778954\Releases\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\AMM\Illustrations\Illustrations"
            }
            elseif($icnName.Contains("-KE"))
            {
             $illPath =  "\\a5778954\Releases\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\ARD\Illustrations\Illustrations"
            }
            elseif($icnName.Contains("-KH"))
            {
             $illPath =  "\\a5778954\Releases\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\ACS\Illustrations\Illustrations"
            }
            $icnSourcePath = "$illPath\$icnName"
            Copy-Item -Path $icnSourcePath -Destination $illustrationsdDestination -Verbose -ErrorAction Inquire  -Credential $cred
        }
        Copy-Item -Path $FFN -Destination $s1000dDestination -Verbose -ErrorAction Inquire -Credential $cred
    }
    $pmXmlOut.Save($pmcFileNameOut)
    $x | Out-File "$outFolder\$Manual - Discrepency List.txt"
    $dmNames | Out-File "$outFolder\$Manual - List of DMs for BCDR Review.txt"
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