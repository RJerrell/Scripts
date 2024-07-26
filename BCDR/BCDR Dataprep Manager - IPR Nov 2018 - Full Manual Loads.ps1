
cls
$sd = Get-Date
$ErrorActionPreference = "Stop"
$error.Clear()

# Include all common variables
. 'C:\KC46 Staging\Scripts\BCDR\Common\CommonVariablesJacks.ps1'
Import-Module -Name "KC46Common" -Verbose -Force

Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose

$parserPM = new-object S1000D.PublicationModule_401
$parserDM = new-object S1000D.DataModule_401

#     NO AMM IN THIS LIST....IT WAS PROCESSED IN A DIFFERENT MANNER ABOVE
[string[]] $ManualS   = @("ABDR","AMM","ARD","ASIP", "FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WDM","WUC")
[string[]] $ManualS   = @("AMM","IPB","SIMR","SSM","TC","WDM","WUC")
foreach ($Manual in $ManualS)
{
    $s1000dDestination = "$outputFolderBaseName$Manual$pathSuffix\S1000D"
    
    $illustrationsdDestination =  "$outputFolderBaseName$Manual$pathSuffix\Illustrations"

    if( !(Test-Path -Path $s1000dDestination) )
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

    $dataPath = "$inputPathBaseFolder\$Manual\S1000D\SDLLIVE"

    $pmcFile= gci -Path $dataPath -Filter pmc*.xml | Sort-Object -Descending | select -First 1
    
    # Get the PMC from the manual - remember to get the latest version based on the issue # within the file name
    #$pmXml = New-Object System.Xml.XmlDocument

    #$pmXmlOut = New-Object System.Xml.XmlDocument

    #$pmXml.Load($pmcName[0].FullName)
    $pmcFN = $pmcFile[0].FullName
    $pmcSN = $pmcFile[0].Name
    $parserPM.ParsePM($pmcFN)

    #$pmXmlOut = $parserPM.Pmodule
    # Clone the PMC
    #$pmXmlOut = $pmXml

    $outFolder = "$outputFolderBaseName$Manual$pathSuffix"
    
    # Establish a new name for the new abbreviated verion of the PMC
    #$pmcFileNameOut = "$s1000dDestination\" + $pmcName
     
   
    # Get a list of all the dmc references in the PMC
    # /pm/content/pmEntry/pmEntry/pmEntry/dmRef/dmRefIdent/dmCode

    $dmRefs = $parserPM.DmRefs

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
        $fileToLoad = gci -Path $FFN | Sort-Object -Descending |Select-Object -First 1
        try
        {
            $parserDM.ParseDM($fileToLoad.Fullname)
            #$dm.Load((gci -Path $FFN)[0].FullName)
        
            $graphics = $parserDM.Graphics

            foreach($graphic in $graphics)
            {            
                $icnName = $graphic.infoEntityIdent
                if($icnName.Contains("-KA"))
                {
                 $illPath =  "$inputPathBaseFolder\AMM\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KE"))
                {
                 $illPath =  "$inputPathBaseFolder\ARD\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KH"))
                {
                 $illPath =  "$inputPathBaseFolder\ACS\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KF"))
                {
                 $illPath =  "$inputPathBaseFolder\FIM\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KP"))
                {
                 $illPath =  "$inputPathBaseFolder\IPB\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KC"))
                {
                 $illPath =  "$inputPathBaseFolder\ASIP\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KM"))
                {
                 $illPath =  "$inputPathBaseFolder\LOAPS\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KV"))
                {
                 $illPath =  "$inputPathBaseFolder\NDT\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KJ"))
                {
                 $illPath =  "$inputPathBaseFolder\SIMR\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KQ"))
                {
                 $illPath =  "$inputPathBaseFolder\SPCC\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KR"))
                {
                 $illPath =  "$inputPathBaseFolder\SSM\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KW"))
                {
                 $illPath =  "$inputPathBaseFolder\WDM\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KT"))
                {
                 $illPath =  "$inputPathBaseFolder\TC\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KD"))
                {
                 $illPath =  "$inputPathBaseFolder\WUC\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KZ"))
                {
                 $illPath =  "$inputPathBaseFolder\KC46\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KG"))
                {
                 $illPath =  "$inputPathBaseFolder\ABDR\Illustrations\Illustrations"
                }
                elseif($icnName.Contains("-KU"))
                {
                 $illPath =  "$inputPathBaseFolder\SWPM\Illustrations\Illustrations"
                }
                else
                {
                    "Unknown graphic type :`t" + $graphic.OuterXml
                    SLEEP -Seconds 2
                    #exit
                }
                if($icnName.Length -gt 0)
                {
                    $icnSourcePath = "$illPath\$icnName`*"
                    $icnSourcePath 
                    Copy-Item -Path $icnSourcePath -Destination $illustrationsdDestination -ErrorAction stop
                }
            }
            $ffn
            Copy-Item -Path $FFN -Destination $s1000dDestination -ErrorAction Stop
        }
        catch 
        {
            Write-Host "Referenced DM is outside this book and will not be processed ..."
        }

    }
    #$pmXmlOut.Save($pmcFileNameOut)
    Copy-Item -Path $pmcFN -Destination $s1000dDestination -Force

    #$x | Out-File "$outFolder\$Manual - Discrepency List.txt"

    $dmNames | Out-File "$outFolder - List of DMs for BCDR Review.txt"
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