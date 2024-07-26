cls

$Error.Clear()
$ErrorActionPreference = "Stop"
"Error action preference now set to`t$ErrorActionPreference on hard errors"
#Get-ChildItem NoSuchFile.txt -ErrorAction SilentlyContinue;
#"2 - $ErrorActionPreference;"
#Get-ChildItem NoSuchFile.txt -ErrorAction Stop;
#"3 - $ErrorActionPreference;"

$myArray = @()
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
            'Z' {
            $global:refPub = "Common"
            break
            }

        }        

        $global:l = "DMC-$a-$b-$c-$d$e-$f-$g$h-$i$j-$k"
}

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

[string[]] $PubList   = @($commonRoot, "ACS", "ARD", "AMM", "FIM", "IPB", "NDT", "SIMR", "SRM", "SSM", "SWPM", "WUC", "WDM")
#[string[]] $PubList   = @("IPB", "WUC")
$masterList = @()
foreach ($pub in $PubList)
   {       
       $pub
       $files = gci -Path "$source_BaseLocation\$pub\S1000D\S1000D\DMC*.XML"
       $fCounter = 0
       foreach ($file in $files)
       {
            $fCounter ++            
            #Write-Progress -Activity “Processing the $pub folder ...” -status “Finding file $file” -percentComplete ($fCounter / $files.count*100)            
            $file.FullName
            $dmc =[xml](Get-Content -Path $file.FullName)

            Get-DMCMetaDataFromDMCode -file $file.FullName            
            $dmcodeColl = $dmc.SelectNodes("/dmodule/content/refs//dmCode")
            if($dmcodeColl.Count -gt 0)
            {
                $fileName = $file.Name.Replace(".xml","")
                $techName = $DMC.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
                $infoName = $DMC.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName
                #$obj = [pscustomobject][ordered]@{Manual=$pub;DMC=$fileName;TechName=$techName;InfoName=$infoName;DMType=$global:dmType;
                # RefManual="";RefDMC="";RefTechName="";RefInfoName="";RefDMType=""}

                [int32] $ctr =  0

                foreach($dmCode in $dmcodeColl)
                {
                    $obj = New-Object System.Object
                    $obj | Add-Member -name 'Manual' -MemberType NoteProperty -Value $pub
                    $obj | Add-Member -name 'DMC' -MemberType NoteProperty -Value $fileName
                    $obj | Add-Member -name 'TechName' -MemberType NoteProperty -Value $techName
                    $obj | Add-Member -name 'InfoName' -MemberType NoteProperty -Value $infoName
                    $obj | Add-Member -name 'DMType' -MemberType NoteProperty -Value $global:dmType
                    $obj | Add-Member -name 'RefDMType' -MemberType NoteProperty -Value ""
                    $obj | Add-Member -name 'RefManual' -MemberType NoteProperty -Value ""             
                    $obj | Add-Member -name 'RefDMC' -MemberType NoteProperty -Value ""
                    $obj | Add-Member -name 'RefTechName' -MemberType NoteProperty -Value ""
                    $obj | Add-Member -name 'RefInfoName' -MemberType NoteProperty -Value ""

                    Create-DMCFileNameFromDMCode -dmCode $dmCode
                    if($l -eq "DMC-1KC46-A-05-00-0000-02A0A-912A-A")
                    {
                        "stop here"
                    }
                    $path = "$source_BaseLocation\$global:refPub\s1000d\s1000d\$l.xml"

                    $fi = gci -Path $path

                    $refDM = [xml] ( Get-Content -Path $path )

                    $obj.RefManual =  $global:refPub              
                    $obj.RefDMC = $fi[0].Name.Replace(".xml", "")
                    $obj.RefTechName = $refDM.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
                    $obj.RefInfoName = $refDM.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName

                    if(($refDM.DocumentElement.content.ChildNodes.Count -eq 1) -or ($refDM.DocumentElement.content.ChildNodes.Count -eq 2))
                    {
                        $obj.RefDMType = $refDM.DocumentElement.content.ChildNodes[1].Name
                    }
                    elseif($refDM.DocumentElement.content.ChildNodes.Count -eq 0)
                    {
                        $obj.RefDMType = $refDM.DocumentElement.content.ChildNodes[0].Name
                    }
                    else
                    {
                        $obj.RefDMType = "$pub--DocType Unknown"
                    }
                                        
                    $myArray += $obj
                    $obj=$null
                }      
            }
       }
   }    
if(!(Test-Path -Path "C:\KC46 Staging\Scripts\Report Generators"))
{
    md "C:\KC46 Staging\Scripts\Report Generators"
}

<#

$stream = [System.IO.StreamWriter] "C:\KC46 Staging\Scripts\Report Generators\Outputs\t.txt"
$s = "abc"
1..10000 | % {

      $stream.WriteLine($s)

}

$stream.close()
#>

$myArray | Export-Csv "C:\KC46 Staging\Scripts\Report Generators\Outputs\AllMetadataReport.csv"
$Error.Sort()
$Error