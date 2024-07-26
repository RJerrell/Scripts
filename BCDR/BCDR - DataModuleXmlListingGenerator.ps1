CLS
$ErrorActionPreference = "Stop"
$error.Clear()
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\KC46 Staging\Scripts\Modules;"
Import-Module -Name "KC46Common" -Verbose -Force					
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
# ********************************************************************
#region Vars

    $outputPath = "$resourcePath\OUTPUTS"
    [string[]] $ManualS   = @("ABDR","AMM","ARD","ASIP", "FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WUC","WDM") | Sort-Object
    [string[]] $ManualS   = @("AMM","IPB","SIMR","SSM","TC","WDM","WUC")
    $dmData_Rootpath = "D:\Shared\BCDR\IPR\100 Percent IPR - Nov 2018\S1000D"
    $BCDRTemplate = "C:\KC46 Staging\Scripts\Templates\BCDR - DataModuleXmlListingGenerator.xml"

#endregion

Function ProcessDm
{
    Param([string] $path)
    $files = gci -Path "$path\DMC*.xml" -Verbose -Recurse
    $pmList = gci -Path "$path\PMC*.xml" -Verbose -Recurse | Sort-Object -Descending | Select-Object -First 1
    foreach($pm in $pmList)
    {
        $parserPM.ParsePM($pm)
        $dmCodes = $parserPM.DmRef_DmCodes
        $allFiles = gci -Path $path -Filter "DMC`*" -Recurse
        $removed = 0  
        foreach($f in $allFiles)
        {
            $dmcParts = $f.Name.Split("_")
            $dmcPart = $dmcParts[0]
            if(! ($dmCodes.Contains($dmcPart )))
            {
               Remove-Item -Path $f.Fullname -Verbose   
                $removed ++  
            }
        }
        $removed
        foreach($dmCode in $dmCodes)
        {
            $files = gci -Path $path -Filter "$dmCode`*" -Recurse
            foreach($file in $files)
            {   $dm = New-Object System.Xml.XmlDocument
                $fullName = $file.FullName
                $dm.Load($fullName)
                $dirty = $false
                if($dm.SelectNodes("/dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle").ChildNodes.Count -gt 2)
                {
                    $dm.removeChild("/dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle/infoName")
                    $dirty = $true
                }

                if($dm.SelectNodes("/dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle").ChildNodes.Count -lt 2)
                { 
                    $infoNameEle = [System.Xml.XmlNode] $dm.createElement("infoname")
                    $infoNameEle.InnerText = "UNK"
                    $dm.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.AppendChild($infoNameEle)
                    $dirty = $true
                }
                if($dirty)
                {
                    $dm.Save($fullName)  
                    $dm.Load($fullName)
                    $file.Name
                }
                Set-Element -dm $dm -Name $file.Name
            }
        }
    }
}

Function Set-Element
{
Param([xml] $dm, [string] $Name )

    $element = $global:IPRXml.dmodule.FirstChild.clone()
    $element.dmkey = $file.Name.Substring(4,$Name.Length - 8)
    $element.dmtechname = ([string] $dm.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName).Replace("`r`n","")
    $element.dminfoname = ([string] $dm.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName).Replace("`r`n","")
    $element.dmtitle = ([string] $element.dmtechname + "-" + $element.dminfoname).Replace("`r`n","")
    $null = $global:IPRXml.DocumentElement.AppendChild($element) 
}

$global:templates = [XML] (Get-Content -Path $BCDRTemplate)

$iprNames = (gci -Path $dmData_Rootpath -Directory) | Sort-Object -Descending

foreach ($iprName in $iprNames)
{
    "Prpocessing: `t" + $iprName
    $global:IPRXml = New-Object System.Xml.XmlDocument
    $global:IPRXml.LoadXml($global:templates.OuterXml)
    $dirName = $iprName.Name
    
    ProcessDm -path "$dmData_Rootpath\$dirName"

    $null = $global:IPRXml.dmodule.RemoveChild($IPRXml.dmodule.FirstChild)
    $global:IPRXml.Save("$dmData_Rootpath\$dirName\$iprName.xml")
}

# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"
