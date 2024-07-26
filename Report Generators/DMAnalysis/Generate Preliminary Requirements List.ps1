cls
"Start"
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\KC46 Staging\Scripts\Modules;"
$env:PSModulePath
Import-Module -Name "KC46Common" -Verbose -Force

function Truncate-PreliminaryReqNode
{
    Param([System.Xml.XmlElement] $prnode)
    $seColl = $prnode.SelectNodes("//supportEquipDescr")
    $seArray = @()
    foreach ($se in $seColl)
    {
        $sehash = [string] ($se.Name + $se.toolRef.toolNumber)
        if($se -notcontains $sehash)
        {
            $seArray += $sehash
            $prnode.RemoveChild($se)
        }

    }
    return $revisedPRNode
}

function Create-DM
{
    Param(
        [System.Xml.XmlDocument] $refDM_XML, 
        [string] $manual , 
        [System.Xml.XmlElement] $parentElement = $null, 
        [string] $id)

    "Processing . . . . DMC-1KC46-A-05-51-0100-04A0A-311A-A"
    
    $pr = $false
    $reqCondGroup = $refDM_XML.SelectNodes("/dmodule/content/procedure/preliminaryRqmts/reqCondGroup").ChildNodes.IsEmpty
    $reqSupportEquips = $refDM_XML.SelectNodes("/dmodule/content/procedure/preliminaryRqmts/reqSupportEquips").ChildNodes.IsEmpty
    $reqSupplies = $refDM_XML.SelectNodes("/dmodule/content/procedure/preliminaryRqmts/reqSupplies").ChildNodes.IsEmpty
    $reqSpares = $refDM_XML.SelectNodes("/dmodule/content/procedure/preliminaryRqmts/reqSpares").ChildNodes.IsEmpty
    $reqSafety = $refDM_XML.SelectNodes("/dmodule/content/procedure/preliminaryRqmts/reqSafety").ChildNodes.IsEmpty

    $refDMS = $refDM_XML.SelectNodes("/dmodule/content/refs//dmCode")
    $reqSupportEquips
    if($reqCondGroup -eq $false -or $reqSupportEquips -eq $false -or $reqSupplies -eq $false-or $reqSpares -eq $false -or $reqSafety -eq $false)
    {
        $pr = $true
    }
    if($pr -eq $false -and $refDMS.Count -eq 0)
    {
        "Do nothing"
        return
    }
    else
    {   
        # Creation of a node and its text
        $xmlElt = $global:shoppingCartXML.CreateElement("datamodule")
        $idATT = $global:shoppingCartXML.CreateAttribute("id")
        $manualATT = $global:shoppingCartXML.CreateAttribute("manual")
        $manualATT.Value = $manual
        $idATT.Value = $id
        $xmlElt.Attributes.Append($idATT)
        $xmlElt.Attributes.Append($manualATT)
        $dmCode  = $global:shoppingCartXML.ImportNode($refDM_XML.DocumentElement.identAndStatusSection.dmAddress, $true)
        $xmlElt.AppendChild($dmCode)
        if($pr -eq $true)
        {
            #"/dmodule/content/procedure/preliminaryRqmts"
            $prNode  = $global:shoppingCartXML.ImportNode($refDM_XML.DocumentElement.content.procedure.preliminaryRqmts, $true)
            $revPRNode = Truncate-PreliminaryReqNode -prnode $prNode
            $xmlElt.AppendChild($revPRNode)
        }
        $refDMS = $refDM_XML.SelectNodes("/dmodule/content/refs//dmCode")        
        foreach ($refDM in $refDMS)
        {
            $dmRefdmCode = Create-DMCFileNameFromDMCode -dmCode $refDM
            if($id -ne $dmRefdmCode)
            {
                $dmRefDoctype = Get-DocTypeFromDMC -dc $dmRefdmCode
                $dmRefpath = "$global:source_BaseLocation\$dmRefDoctype\S1000D\S1000D\$dmRefdmCode.xml"
                $RefDataModuleXml = [xml] (Get-Content -Path $dmRefpath -Raw)
                $isProcedure =  $RefDataModuleXml.SelectNodes("/dmodule/content/procedure").Count
                
                # THIS TEST IS LIKELY NOT CORRECT - RESEARCH IT AND CHANGE IT!!!
                if(($isProcedure -gt 0) -and ( $global:RefList -notcontains  $dmRefdmCode))
                {
                    if($parentElement -ne $null)
                    {
                        $parentElement.AppendChild($xmlElt);       
                    }
                    else
                    {
                        $shoppingCartXML.LastChild.AppendChild($xmlElt);                        
                    }
                    
                    $global:RefList += $dmRefdmCode                            
                    Create-DM -refDM_XML $RefDataModuleXml -shoppingCartXML $global:shoppingCartXML -parentElement $xmlElt -manual $dmRefDoctype -id $dmRefdmCode                    
                }
            }
        }
    }    
}

$global:environment = "Production"  # *************   Override to Production  ************#

# Where the source S1000D data is located that will eventually become an IETM
$Global:KC46DataRoot = "C:\KC46 Staging"

$commonRoot = "KC46"
$global:source_BaseLocation = "$KC46DataRoot\$environment\Manuals"

[string[]] $PubList   = @($commonRoot, "ACS", "ARD", "AMM", "FIM", "IPB", "LOAPS", "NDT", "SIMR", "SRM", "SSM", "SWPM", "WUC", "WDM")
[string[]] $PubList   = @("AMM")

$global:shoppingCartXML = New-Object system.Xml.XmlDocument
$global:shoppingCartXML.LoadXml("<?xml version=`"1.0`" encoding=`"utf-8`"?><datamodules></datamodules>")
$global:ParentID = $null
$global:RefList = @()
$global:shoppingCartReportFolder = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$global:shoppingCartReportName = "SHOPPINGCART.xml"

foreach ($pub in $PubList)
{ 

    $path = "$global:source_BaseLocation\$pub\S1000D\S1000D\DMC*.XML"
    $files = (gci -Path $path) | Sort-Object -Property Name
    foreach ($file in $files)
    {  
        if($file.Name -eq "DMC-1KC46-A-05-00-0100-16A0A-310A-A.xml")
        {
        # Parent Data Module
        $dmXml = [xml](Get-Content -Path $file.FullName)
        $isProcedure =  ($dmXml.SelectNodes("/dmodule/content/procedure").Count -gt 0)
        if($isProcedure)
        {
            $dmCode = Create-DMCFileNameFromDMCode -dmCode $dmXml.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode
            $global:parentID = $null
            Create-DM -refDM_XML $dmXml -shoppingCartXML $global:shoppingCartXML -manual $pub -id $dmCode
            $shoppingCartXML.Save("$global:shoppingCartReportFolder\$global:shoppingCartReportName")
        }        
        $dmXml = $null
        $RefList = @()
        }
    }
}

# Store to a file
$shoppingCartXML.Save("$shoppingCartReportFolder\$shoppingCartReportName")