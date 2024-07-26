cls

[string[]] $masterList = @()

$global:modelIdentCode = ""
$global:systemDiffCode = ""
$global:systemCode = ""
$global:subSystemCode = ""
$global:subSubSystemCode = ""
$global:assyCode = ""
$global:disassyCode = ""
$global:disassyCodeVariant = ""
$global:infoCode = ""
$global:infoCodeVariant = ""
$global:itemLocationCode = ""
$global:dmType = ""
$global:dmc = $null
$global:dmcFileName = ""

$global:a = ""
$global:b = ""
$global:c = ""
$global:d = ""
$global:e = ""
$global:f = ""
$global:g = ""
$global:h = ""
$global:i = ""
$global:j = ""
$global:k = ""
$global:l = ""
$global:refDMType = ""
$global:refPub = ""
function Get-DMCMetaDataFromDMCode
{
    param([string] $file)        
        $global:dmc = [xml](Get-Content -Path $file)
        $global:modelIdentCode = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.modelIdentCode
        $global:systemDiffCode = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.systemDiffCode
        $global:systemCode = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.systemCode
        $global:subSystemCode = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.subSystemCode
        $global:subSubSystemCode = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.subSubSystemCode
        $global:assyCode = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.assyCode
        $global:disassyCode = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.disassyCode
        $global:disassyCodeVariant = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.disassyCodeVariant
        $global:infoCode = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.infoCode
        $global:infoCodeVariant = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.infoCodeVariant
        $global:itemLocationCode = $dmc.DocumentElement.identAndStatusSection.dmAddress.dmIdent.dmCode.itemLocationCode
        if($dmc.DocumentElement.content.ChildNodes.Count -gt 1)
        {
            $global:dmType = $dmc.DocumentElement.content.ChildNodes[1].Name
        }
        else
        {
            $global:dmType = $dmc.DocumentElement.content.ChildNodes[0].Name
        }
}

function Create-DMCFileNameFromDMCode
{
    param([System.Xml.XmlNode] $dmCode)
        $global:a = $dmCode.modelIdentCode
        $global:b = $dmCode.systemDiffCode
        $global:c = $dmCode.systemCode
        $global:d = $dmCode.subSystemCode
        $global:e = $dmCode.subSubSystemCode
        $global:f = $dmCode.assyCode
        $global:g = $dmCode.disassyCode
        $global:h = $dmCode.disassyCodeVariant
        $global:i = $dmCode.infoCode
        $global:j = $dmCode.infoCodeVariant
        $global:k = $dmCode.itemLocationCode

        $refP = $dmCode.disassyCodeVariant.Substring(2,1)
        switch ($refP)
        {
            'A' {
            $global:refPub = "AMM"
            break
            }
            'F' {
            $global:refPub = "FIM"
            break
            }
            'P' {
            $global:refPub = "IPB"
            break
            }
            'W' {
            $global:refPub = "WDM"
            break
            }
            'E' {
            $global:refPub = "ARD"
            break
            }
            'V' {
            $global:refPub = "NDT"
            break
            }
            'H' {
            $global:refPub = "ACS"
            break
            }
            'U' {
            $global:refPub = "SWPM"
            break
            }
            'S' {
            $global:refPub = "SRM"
            break
            }
            'Z' {
            $global:refPub = "KC46"
            break
            }
            'R' {
            $global:refPub = "SSM"
            break
            }
            'D' {
            $global:refPub = "WUC"
            break
            }
            'T' {
            $global:refPub = "TC"
            break
            }            
            'M' {
            $global:refPub = "LOAPS"
            break
            }
            'J' {
            $global:refPub = "NDI"
            break
            }
            'K' {
            $global:refPub = "SPCC"
            break
            }
        }        

        $global:l = "DMC-$a-$b-$c-$d$e-$f-$g$h-$i$j-$k"
}

$global:environment = "Production"  # *************   Override to Production  ************#

# Where the source S1000D data is located that will eventually become an IETM
$global:KC46DataRoot = "C:\KC46 Staging"
$commonRoot = "KC46"
# BASIC FOLDER STRUCTURES THAT ARE PRETTY HARD CODED AND EXPECTED TO BE THERE.  IF THESE CHANGE, IT ALL FALLS DOWN.
$LiveContentDataFolder =  "C:\LiveContentData\" # leave the dash on the end of this path statement
$semaphoreLocation = "$KC46DataRoot\$environment"
$unpackLocation = "$KC46DataRoot\$environment"
$archiveRootFolder = "$KC46DataRoot\$environment\Archives"
$buildsRootFolder = "$archiveRootFolder\Builds"
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"

# [string[]] $PubList   = @("KC46", "ACS", "ARD", "AMM", "FIM", "IPB", "LOAPS", "MOM", "NDI",  "NDT", "SIMR", "SPCC",  "SRM", "SSM", "SWPM", "WUC", "WDM")
 [string[]] $PubList   = @("KC46", "ACS", "ARD")
$masterList = @("")

foreach ($pub in $PubList)
   {
       
       $pub
       $files = gci -Path "$source_BaseLocation\$pub\S1000D\SDLLIVE\DMC*.*"
       $fCounter = 1
       foreach ($item in $files)
       {
            Write-Progress -Activity “Processing the $pub folder ...” -status “Finding file $item” -percentComplete ($fCounter / $files.count*100)
            $fCounter ++
            Get-DMCMetaDataFromDMCode -file $item.FullName
            $fileName = $item.Name.Replace(".xml","")
            $baseRec = "$pub|$fileName|$global:dmType|$global:modelIdentCode|$global:systemDiffCode|$global:systemCode|$global:subSystemCode$global:subSubSystemCode|$global:assyCode|$global:disassyCode|$global:disassyCodeVariant|$global:infoCode|$global:infoCodeVariant|$global:itemLocationCode"
            $dmc =[xml](Get-Content -Path $item.FullName)
            $dmcodeColl = $dmc.SelectNodes("/dmodule/content/refs//dmCode")
            $masterList += $baseRec
            if($dmcodeColl.Count -gt 0)
            {
                [int32] $ctr =  0
                foreach($dmCode in $dmcodeColl)
                {
                    Create-DMCFileNameFromDMCode -dmCode $dmCode
                    $refDM = [xml] ( Get-Content -Path "$source_BaseLocation\$global:refPub\S1000D\SDLLIVE\$l.xml" )
                    
                    if($refDM.DocumentElement.content.ChildNodes.Count -gt 1)
                    {
                        $global:refDMType = $refDM.DocumentElement.content.ChildNodes[1].Name
                    }
                    else
                    {
                        $global:refDMType = $refDM.DocumentElement.content.ChildNodes[0].Name
                    }
                    if($ctr -eq 0)
                    {
                        $arrL = $masterList.Length -1
                        
                        $masterList[$arrL] = $masterList[$arrL] + "|$global:refPub|$global:l|$global:refDMType|$global:a|$global:b|$global:c|$global:d$global:e|$global:f|$global:g|$global:h|$global:i|$global:j|$global:k"
                        $ctr = $ctr + 1
                    }
                    else
                    {
                        $masterList += "$baseRec|$global:refPub|$global:l|$global:refDMType|$global:a|$global:b|$global:c|$global:d$global:e|$global:f|$global:g|$global:h|$global:i|$global:j|$global:k"
                    }
                     
                }
        
            }
       }       
   }
    
   if(!(Test-Path -Path "C:\KC46 Staging\Scripts\Report Generators"))
   {
    md "C:\KC46 Staging\Scripts\Report Generators"
   }

   $masterList | Out-File "C:\KC46 Staging\Scripts\Report Generators\AllMetadataReport.txt" -Force