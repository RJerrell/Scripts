
Function Start-BCDRDataModuleXmlListingGenerator
{
    PARAM([string[]] $ManualS,[STRING] $dmData_Rootpath, [string] $outputFolder, [string] $BCDRTemplatePath) 
    if(($ManualS.Count -ne 0) -and (Test-Path - Path $dmData_Rootpath))
    {
        "exiting...."
        exit
    }    
    
    #$dmData_Rootpath = "D:\Shared\BCDR\IPR\100 Percent IPR - March 2018\KC46_100_IPR__03_2018_Rel9"
    #$BCDRTemplate = "C:\KC46 Staging\Scripts\Templates\BCDR - DataModuleXmlListingGenerator.xml"
    $global:templates = [XML] (Get-Content -Path $BCDRTemplatePath)

    $iprNames = gci -Path $dmData_Rootpath -Directory -R

    foreach ($iprName in $iprNames)
    {
        $global:IPRXml = New-Object System.Xml.XmlDocument
        $global:IPRXml.LoadXml($global:templates.OuterXml)
        $dirName = $iprName.Name
        ProcessDm -path "$dmData_Rootpath\$dirName"
        $null = $global:IPRXml.dmodule.RemoveChild($IPRXml.dmodule.FirstChild)
        $global:IPRXml.Save("$dmData_Rootpath\$dirName\$iprName.xml")
    }
}

Function ProcessDm
{
    Param([string] $path)
    $files = gci -Path "$path\DMC*.xml" -Verbose -Recurse
    foreach($file in $files)
    {   $dm = New-Object System.Xml.XmlDocument
        $fullName = $file.FullName
        $dm.Load($fullName)
        $dirty = $false
        if($dm.SelectNodes("/dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle").ChildNodes.Count -ge 2)
        {
            $dm.removeChild("/dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle/infoname")
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
Function Set-Element
{
Param([xml] $dm, [string] $Name )

    $element = $global:IPRXml.dmodule.FirstChild.clone()
    $element.dmkey = $file.Name.Substring(4,$Name.Length - 8)
    $element.dmtechname = $dm.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
    $element.dminfoname = $dm.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName
    $element.dmtitle = $element.dmtechname + "-" + $element.dminfoname
    $null = $global:IPRXml.DocumentElement.AppendChild($element) 
}

