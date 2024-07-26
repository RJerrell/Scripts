Function Get-FileExistance
{
    Param([string] $fullName)
    $exists = $false
    if(Test-Path -Path $fullName)
    {
        $exists = $true
    }
    $exists
}
Function Remove-DuplicateSymbols
{
    Param([string] $path)
    $findThis = '<symbol infoEntityIdent="ICN-81205-KZ90000001-001-01" /><symbol infoEntityIdent="ICN-81205-KZ90000001-001-01" />'
    $findThis2 = '<symbol infoEntityIdent="ICN-81205-KZ90000001-001-01" /> <symbol infoEntityIdent="ICN-81205-KZ90000001-001-01" />'
    $ChangeToThis = '<symbol infoEntityIdent="ICN-81205-KZ90000001-001-01" />'
    
    $global:siteURL = "https://collab.web.boeing.com/sites/KC46TankerTechPubs/CEERS"
    $global:listName = "KC46 - Tanker Tech Pubs CEERS"
    $files = Get-ChildItem -Path $path -File -Recurse -Filter DMC*.xml

    $settings = New-Object System.Xml.XmlWriterSettings;

    # Insure that no Unicode Byte Order Mark precedes the preamble in the file when saved to disk            
    # ************************ Critical that this encoding is used ****************************** 
    $encoding = New-Object System.Text.UTF8Encoding($False)
    # ************************ Critical that this encoding is used ****************************** 

    $xdoc = New-Object System.Xml.XmlDocument
    $ctr = 1
    foreach ($file in $files)
    {
        $sreader = New-Object System.IO.StreamReader($file.FullName)
        $ctr
        $fileContent = $sreader.ReadToEnd()
        $sreader.Close()
        $sreader.Dispose()
        if($fileContent.Contains($findThis) -or $fileContent.Contains($findThis2) )
        {
            try
            {
                $revisedContents =  $fileContent.Replace($findThis, $ChangeToThis)
                $revisedContents =  $fileContent.Replace($findThis2, $ChangeToThis)
                $xdoc.LoadXml($revisedContents)
        
                "Saving : `t" + $file.FullName
                # attrib -R $file.FullName;
                Set-ItemProperty -Path $file.FullName -Name IsReadOnly -Value $false -Force -Verbose
                #"Saving : `t" + $file.FullName  
                # *****************************************************************************************
                $settings.Indent = $true;
                $settings.OmitXmlDeclaration = $false;
                $settings.NewLineOnAttributes = $false;
                $settings.Encoding = $encoding
                $settings.WriteEndDocumentOnClose = $false
                $settings.CheckCharacters = $true
                $settings.DoNotEscapeUriAttributes = $false

                #Create an XmlWriter to insure the output XML conforms to the settings above
                $writer = [System.Xml.XmlWriter]::Create($file.FullName,$settings)
                "Cleaning duplicate ESDS tags from DM: `t" + $file.FullName 
                $xdoc.Save($writer);
                $writer.Flush();
                $writer.Close();
                $writer.Dispose();
                Add-CEERSMessage -siteUrl $siteUrl -listName $listName -strSource "KC46 - Cleanup Duplicate ESDS sysmbols if present.ps1" -strActivity "$m4 : Completed"  -stage "Post Processing Cleanup" -actionReq $actionReq
            }
            catch [System.Exception]
            {
                "Loading and Saving of " + $file.FullName  + " failed."
            }
        
        }
        $ctr ++
    }
}
Function Get-ICN_ParentBook
{
    Param([string] $ICN_ShortFileName)
    # 2nd position of the 3rd segment
    $a = $ICN_ShortFileName.Split("-")[2].ToUpper().Substring(1,1)
    switch ($a) 
        { 
            G {$pub = "ABDR"}
            H {$pub = "ACS"}
            A {$pub = "AMM"}            
            E {$pub = "ARD"}
            F {$pub = "FIM"}
            P {$pub = "IPB"} 
            Z {$pub = "KC46"}
            M {$pub = "LOAPS"} 
            V {$pub = "NDT"}
            J {$pub = "SIMR"}
            Q {$pub = "SPCC"}
            S {$pub = "SRM"}
            R {$pub = "SSM"} 
            U {$pub = "SWPM"}
            T {$pub = "TC"} 
            W {$pub = "WDM"}
            D {$pub = "WUC"}            
            default {$pub = "unknown"}
        }
    $pub
}
Function Get-Bookletters
{
    return $bookLetters = @{"ABDR"="G";"ACS"="H";"AMM"="A";"ARD"="E";"ASIP"="X";"FIM"="F";"IPB"="P";"KC46"="Z";"LOAPS"="M";"NDT"="V";"SIMR"="U";"SPCC"="Q";"SSM"="R";"SWPM"="U";"TC"="T";"WDM"="W";"WUC"="D";}
}

Function Set-DMLEntry
{
Param([string] $FFN)

$dmlEntryTemplate = [XML] @"
<dmEntry dmEntryType="c">
    <dmRef >
	    <dmRefIdent>
		    <dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="00" subSystemCode="0" subSubSystemCode="0" assyCode="0000" disassyCode="01" disassyCodeVariant="A0K" infoCode="018" infoCodeVariant="A" itemLocationCode="D"/>
		    <language countryIsoCode="US" languageIsoCode="sx"/>
	    </dmRefIdent>
	    <dmRefAddressItems>
		    <dmTitle>
			    <techName></techName>
			    <infoName></infoName>
		    </dmTitle>
	    </dmRefAddressItems>
    </dmRef>
    <security securityClassification="" commercialClassification="cc52"/>
    <responsiblePartnerCompany>
	    <enterpriseName>Headquarters, Department of The Air Force</enterpriseName>
    </responsiblePartnerCompany>
</dmEntry>
"@
    $dmc = New-Object System.Xml.XmlDocument
    $dmc.Load($FFN)
    $dmlEntryNode = [xml] $dmlEntryTemplate.Clone()
    $dmEntryTypeValue = ""
    $issueType = $dmc.dmodule.identAndStatusSection.dmStatus.issueType
    switch ($issueType.ToLower())
    {
        'new' {
            $dmEntryTypeValue = "n"
            break
        }
        'deleted' {
            $dmEntryTypeValue = "d"
            break
        }
        default {
            $dmEntryTypeValue = "c"
            break
        }
    }
    $dmlEntryNode.dmEntry.dmEntryType = $dmEntryTypeValue
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.modelIdentCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.modelIdentCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.systemDiffCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.systemDiffCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.systemCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.systemCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.subSystemCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.subSystemCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.subSubSystemCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.subSubSystemCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.assyCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.assyCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.disassyCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.disassyCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.disassyCodeVariant = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.disassyCodeVariant.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.infoCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.infoCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.infoCodeVariant = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.infoCodeVariant.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.dmCode.itemLocationCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.itemLocationCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.language.countryIsoCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.language.countryIsoCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefIdent.language.languageIsoCode = $dmc.dmodule.identAndStatusSection.dmAddress.dmIdent.language.languageIsoCode.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefAddressItems.dmTitle.techName  = $dmc.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName.Trim()
    $dmlEntryNode.dmEntry.dmRef.dmRefAddressItems.dmTitle.infoName = $dmc.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName.Trim()
    $dmlEntryNode.dmEntry.security.securityClassification = $dmc.dmodule.identAndStatusSection.dmStatus.security.securityClassification.Trim()
    $dmlEntryNode.dmEntry.security.commercialClassification = "cc52"
    $dmlEntryNode.dmEntry.responsiblePartnerCompany.enterpriseName = "Headquarters, Department of The Air Force"
    return $dmlEntryNode.OuterXml.ToString()
}

Function Save-PrettyXML
{
Param([string] $FName, [System.Xml.XmlDocument] $xmlDoc)
    # Insure that no Unicode Byte Order Mark precedes the preamble in the file when saved to disk            
    # ************************ Critical that this encoding is used ****************************** 
    $encoding = New-Object System.Text.UTF8Encoding($False)
    # *******************************************************************************************
    $settings = New-Object System.Xml.XmlWriterSettings;
    $settings.Indent = $true;
    $settings.OmitXmlDeclaration = $false;
    $settings.NewLineOnAttributes = $false;
    $settings.Encoding = $encoding
    $settings.WriteEndDocumentOnClose = $false
    $settings.CheckCharacters = $false
    $settings.DoNotEscapeUriAttributes = $false

    #Create an XmlWriter to insure the output XML conforms to the settings above
    $writer = [System.Xml.XmlWriter]::Create($FName,$settings)
    $xmlDoc.Save($writer);
    $writer.Flush();
    $writer.Close();
    $writer.Dispose();
}
Function Resolve-BadCharactersInDDN
{
    Param([string] $ddnShortName)
    $path = "C:\KC46 Staging\Production\Archives\Source\UnpackHere\$ddnShortName"
    $files = Get-ChildItem -Path $path -Filter "DMC*.XML"
    foreach ($file in $files)
    {
        $sr = [System.IO.StreamReader] $file.FullName
        $c = $sr.ReadToEnd();
        $sr.Dispose()

        if($c.Contains('�') -or $c.Contains('&#x2013;') -or $c.Contains('“'))
        {
           
            $file.FullName
            #attrib -r $file.FullName
            Set-ItemProperty -Path $file.FullName -Name IsReadOnly -Value $false -Force -Verbose
            $sw = new-Object System.IO.StreamWriter($file.FullName)
                 
            $c = $c.Replace('“', '&#8220;')
            $c = $c.Replace('�', '&#8221;')
            $sw.Write($c)
            $sw.Flush()
            $sw.Dispose()
            $sw = $null
            #attrib +r $file.FullName
        }        
    }       
}
Function Resolve-BadCharacters
{
    $commonRoot = "KC46"
    [string[]] $PubList   = @("KC46", "ABDR", "ACS", "AMM", "ARD", "ASIP", "FIM", "IPB", "LOAPS", "NDT", "SIMR",  "SSM", "SRM", "SWPM", "TC", "WUC", "WDM")
    foreach($pub in $PubList)
    {   
        $path1 = "C:\KC46 Staging\Production\Manuals\$pub\S1000D\SDLLIVE\dmc*"
        $path2 = "C:\KC46 Staging\Production\Manuals\$pub\S1000D\S1000D\dmc*"
        $files1 = Get-ChildItem -Path $path1
        #$files2 = Get-ChildItem -Path $path2
        foreach ($file in $files1)
        {
            $sr = [System.IO.StreamReader] $file.FullName
            $c = $sr.ReadToEnd();
            $sr.Dispose()

            if($c.Contains('�') -or $c.Contains('&#x2013;') -or $c.Contains('“'))
            {

                $file.FullName
                #attrib -r $file.FullName
                Set-ItemProperty -Path $file.FullName -Name IsReadOnly -Value $false -Force -Verbose
                $sw = new-Object System.IO.StreamWriter($file.FullName)
             
                $sw.Write($c)
                $sw.Flush()
                $sw.Dispose()
                $sw = $null
                #attrib +r $file.FullName
            }        
        }
        foreach ($file in $files2)
        {
            $sr = [System.IO.StreamReader] $file.FullName
            $c = $sr.ReadToEnd();
            $sr.Dispose()
            # “Normal Loading”

            if($c.Contains('�') -or $c.Contains('&#x2013;') -or $c.Contains('“'))
            {
 
                $file.FullName
                #attrib -r $file.FullName
                Set-ItemProperty -Path $file.FullName -Name IsReadOnly -Value $false -Force -Verbose
                $sw = new-Object System.IO.StreamWriter($file.FullName)
                      
                #$c = $c.Replace('“', "&#8220;")
                #$c = $c.Replace('�', "&#8221;")
                $sw.Write($c)
                $sw.Flush()
                $sw.Dispose()
                $sw = $null
                #attrib +r $file.FullName
            }        
        }
    }
}

Function Reset-DoctypesPriorToBuild
{
    Param([string] $pathtoXml)
    $files = Get-ChildItem -Path $pathtoXml -Filter DMC*.XML
    
    foreach ($file in $files)
    {         
        Set-ItemProperty $file.FullName -name IsReadOnly -value $false -Force       
        $sr = new-object System.IO.StreamReader $file.FullName
        $c = $sr.ReadToEnd();
        $sr.Close()
        $sr.Dispose()
        $sr = $null

        $posStart = 0
        $posEnd = 0
        $docTypeString = "?>"
        $docTypeEndString = "<dmodule"
        $posStart = $c.indexOf($docTypeString) + 2
        $posEnd = $c.indexOf($docTypeEndString)
        if($c.Contains($docTypeString) -and  ($posEnd -gt 0) -and ($posStart -gt 0))
        {

            # Remove the doctype nodes from the top of each data module
            $c = $c.Remove($posStart, ($posEnd - $posStart))
            $sw = new-Object System.IO.StreamWriter($file.FullName)
            $sw.Write($c)
            $sw.Flush()
            $sw.Dispose()
            $sw = $null
        }        
    }       
}

Function Get-CurrentLine
{
    $Myinvocation.ScriptlineNumber
}

function Add-CEERSMessage
{
    Param([string] $siteUrl, [string] $listName, [string] $strSource, [string] $strActivity, [string] $stage, [bool] $actionReq = $false , [string] $eventType = "Information")

    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    $context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext($siteUrl)
    $context.Credentials = New-Object System.Net.NetworkCredential("SVCKC46_SP_LOGGER", 'hu3"Gx7Wmf%!r$5%5GA64Lg%$&4"3jir', "NW")

    [Microsoft.SharePoint.Client.ListCollection] $listCollection = $context.Web.Lists
    [Microsoft.SharePoint.Client.List] $targetList = $context.Web.Lists.GetByTitle($listName)
    $user = [Microsoft.SharePoint.Client.User]::Get
    $context.Load($targetList)
    [Microsoft.SharePoint.Client.ListItemCreationInformation] $itemCreateInfo = New-Object  -TypeName Microsoft.SharePoint.Client.ListItemCreationInformation
    [Microsoft.SharePoint.Client.ListItem] $newItem = $targetList.AddItem($itemCreateInfo)

    $newItem["Title"] = $strSource
    $newItem["Activity"] = $strActivity
    $newItem["Stage"] = $stage
    $newItem["ActionRequired"] = $actionReq
    $newItem["User"] = "$env:USERDOMAIN\$env:USERNAME"
    $newItem["ProcessingServer"] = $env:COMPUTERNAME
    $newItem.Update()
    try
    {
        $context.ExecuteQuery()
        $context.Web.CurrentUser.LoginName
    }
    catch
    {
        
    }
}

Function Get-FileNameArray
{
 Param([string] $fileName)
 $parts = $fileName.Split("-")
 return $parts
}

Function Get-TechNameInfoNameFromDMC
{
    Param([string] $fileName)
    $dm = [xml](Get-Content -Path $fileName)
    $DocArray = @()
    $TName = $dm.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
    $IName = $dm.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName
    $docTypeContentNode = $dm.dmodule.content
    # /dmodule/content
    $nodeCount = $docTypeContentNode.ChildNodes.Count
    if($nodeCount -eq 1)
    {
        $docType = $dm.dmodule.content.ChildNodes[0].Name        
    }
    else
    {
        $docType = $dm.dmodule.content.ChildNodes[1].Name
    }
    #$docType = $docTypeContentNode.childNodes.LastChild
    $DocArray += $TName
    $DocArray += $IName
    $DocArray += $docType
    return $DocArray
}

Function Get-CVRecord
{
    Param([string] $siteUrl, [string] $listName)

    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    $context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext($siteUrl)
    $context.Credentials = New-Object System.Net.NetworkCredential("SVCKC46_SP_LOGGER", 'Boeing$1', "NW")
    $web = $context.Web
    $context.Load($web, $w -eq $w.Title, $w -eq $w.Description)
    $context.ExecuteQuery()
    $web.Title
}

Function Get-S1000DIssueNumber
{
    Param([string] $fullPath)

    $dm = New-Object System.Xml.XmlDocument
    $dm.Load($fullPath)
    # /dmodule/identAndStatusSection/dmAddress/dmIdent/issueInfo
    $issueNumber = $dm.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.issueNumber

    return $issueNumber
}

Function Get-PubNameFromDCV
{
param([string] $dcv)
    $book = [string] $dcv.Substring(2,1)
    $pub = "UNKNOWN"
    switch ($book)
    {
        'A' {$pub = "AMM"}
        'D' {$pub = "WUC"}
        'E' {$pub = "ARD"}
        'F' {$pub = "FIM"}
        'G' {$pub = "ABDR"}
        'H' {$pub = "ACS"}
        'J' {$pub = "SIMR"}
        'M' {$pub = "LOAPS"}
        'P' {$pub = "IPB"}
        'R' {$pub = "SSM"}
        'S' {$pub = "SRM"}
        'T' {$pub = "TC"}
        'U' {$pub = "SWPM"}
        'V' {$pub = "NDT"}
        'W' {$pub = "WDM"}
        'Z' {$pub = "KC46"}

    }
    return $pub
}
Function  Reset-AllDmRefsTitles
{
    Param( [string[]] $PubList ) 
    
    foreach ($pub in $PubList)
    {        
        #$path1 = "$source_BaseLocation\$pub\S1000D"
        #$path2 =  "$source_BaseLocation\$pub\S1000D\SDLLIVE"
        $path1 = "C:\KC46 Staging\Dev\Manuals\ABDR\S1000D\SDLLIVE"
        $files1 = Get-ChildItem -Path "$path1\DMC*.XML"
        #$files2 = Get-ChildItem -Path "$path2\DMC*.XML"

        foreach ($file in $files1)
        {
            Reset-SingleDmRefsTitle -filename $file.FullName -sourceFolder $path1
        }       
        break
        foreach ($file in $files2)
        {
            Reset-SingleDmRefsTitle -filename $file.FullName -sourceFolder $path2
        }
    }    
}

Function Reset-SingleDmRefsTitle
{
    Param([string] $filename , [string] $sourceFolder)
    
    [System.XML.XMLDocument] $xdoc = New-Object System.XML.XMLDocument

    $dmAddressTemplate = [xml]@"
<dmRefAddressItems><dmTitle><techName></techName><infoName></infoName></dmTitle></dmRefAddressItems>
"@


    $xdoc.Load($filename)
    $dmRefs = $xdoc.SelectNodes("//content//dmRef")
    $dirty = $false

    # Cycle through each dmRef and go get the address portion of the dmcode
    foreach ($dmRef in $dmRefs)
    {       
        
        # Get the techName from the referred to DM          
        $dmRefCode = [System.Xml.XmlNode] $dmRef.dmRefIdent.dmCode
            
        $pub = Get-PubNameFromDCV -dcv $dmRefCode.disassyCodeVariant

        $refFileName = Create-DMCFileNameFromDMCode -dmCode $dmRefCode

        $dms = Get-ChildItem -Path "c:\kc46 staging\production\manuals\$pub\s1000d\sdllive" -Filter "$refFileName`*" |Sort-Object -Descending | Select-Object -First 1
        [System.XML.XMLDocument] $xrefdoc = New-Object System.XML.XMLDocument
        if($dms.Count -eq 0)
        {
            $filename
            $refFileName            
        }
        else
        {
            #$dms[0].FullName
            if($dmRef.dmRefAddressItems.ChildNodes.Count -eq 0)
            {
                $xrefdoc.Load($dms[0].FullName)

                $tName = [string] $xrefdoc.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
                $infoName = [string] $xrefdoc.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName

                $dmAddressNode = $dmAddressTemplate.Clone()
                $dmAddressNode.dmRefAddressItems.dmTitle.techName = $tName
                $dmAddressNode.dmRefAddressItems.dmTitle.infoName = $infoName

                $newNode = $xdoc.ImportNode($dmAddressNode.ChildNodes[0], $true)
                $null = $dmRef.AppendChild($newNode)

                $dirty = $true
            }
        }        
    }
    if($dirty)
    {
        #"Saving`t" + $filename
        # attrib -R $filename
        Set-ItemProperty -Path $filename -Name IsReadOnly -Value $false -Force -Verbose
        $xdoc.Save($filename)
    }
}

Function Create-DMCFileNameFromDMCode
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

        return "DMC-$a-$b-$c-$d$e-$f-$g$h-$i$j-$k"
}

Function Get-DocTypeFromDMC
{
    param([string] $dc)

    $dmcArray = $dc.Split("-")
    $refP = $dmcArray[6].Substring(($dmcArray[6].Length - 1),1)
    switch ($refP)
        {
            'A' {
            $doctype =  "AMM"
            break
            }
            'G' {
            $doctype =  "ABDR"
            break
            }
            'F' {
            $doctype =  "FIM"
            break
            }
            'J' {
            $doctype =  "SIMR"
            break
            }
            'P' {
            $doctype =  "IPB"
            break
            }
            'W' {
            $doctype =  "WDM"
            break
            }
            'E' {
            $doctype =  "ARD"
            break
            }
            'V' {
            $doctype =  "NDT"
            break
            }
            'H' {
            $doctype =  "ACS"
            break
            }
            'U' {
            $doctype =  "SWPM"
            break
            }
            'S' {
            $doctype =  "SRM"
            break
            }
            'Z' {
            $doctype =  "KC46"
            break
            }
            'R' {
            $doctype =  "SSM"
            break
            }
            'D' {
            $doctype =  "WUC"
            break
            }
            'T' {
            $doctype =  "TC"
            break
            }            
            'M' {
            $doctype =  "LOAPS"
            break
            }
            'Z' {
            $doctype =  "Common"
            break
            }
        }
    return $doctype

}

Function Get-FilenameFromDMRef
{
    Param([System.Xml.XmlElement] $dmRef , [string] $filePref )
    <# 
        Inputs
            dmRef    : The S1000D dmRef element and its children
            filePref : the "DMC" prefix on file names
        Outputs
            This Function returns the concatenated assembly of all the attributes of an S1000D dmcode element    
    #>
        $modelIdentCode = $dmRef.dmRefIdent.dmCode.modelIdentCode
        $systemDiffCode = $dmRef.dmRefIdent.dmCode.systemDiffCode
        $systemCode = $dmRef.dmRefIdent.dmCode.systemCode
        $subSystemCode = $dmRef.dmRefIdent.dmCode.subSystemCode
        $subSubSystemCode = $dmRef.dmRefIdent.dmCode.subSubSystemCode
        $assyCode = $dmRef.dmRefIdent.dmCode.assyCode
        $disassyCode = $dmRef.dmRefIdent.dmCode.disassyCode
        $disassyCodeVariant = $dmRef.dmRefIdent.dmCode.disassyCodeVariant
        $infoCode = $dmRef.dmRefIdent.dmCode.infoCode
        $infoCodeVariant = $dmRef.dmRefIdent.dmCode.infoCodeVariant
        $itemLocationCode = $dmRef.dmRefIdent.dmCode.itemLocationCode

        return "$filePref-$modelIdentCode-$systemDiffCode-$systemCode-$subSystemCode$subSubSystemCode-$assyCode-$disassyCode$disassyCodeVariant-$infoCode$infoCodeVariant-$itemLocationCode"
}

Function Create-DMCCodeElementFromDMCFileName
{
    param([string] $dmCode)
    $dmcNodeTemplate = [xml] @"
    <dmCode assyCode="" disassyCode="" disassyCodeVariant="" infoCode="" infoCodeVariant="" itemLocationCode="" modelIdentCode="" subSubSystemCode="" subSystemCode="" systemCode="" systemDiffCode="" />
"@
        $dmArray = $dmCode.Split("-")
        $dmcNode = $dmcNodeTemplate.Clone()
        $dmcNode.dmCode.modelIdentCode = [string] $dmArray[1]
        $dmcNode.dmCode.systemDiffCode = [string] $dmArray[2]
        $dmcNode.dmCode.systemCode = [string] $dmArray[3]
        $dmcNode.dmCode.subSystemCode = [string] $dmArray[4].Substring(0,1)
        $dmcNode.dmCode.subSubSystemCode = [string] $dmArray[4].Substring(1,1)
        $dmcNode.dmCode.assyCode = [string] $dmArray[5]
        $dmcNode.dmCode.disassyCode = [string] $dmArray[6].Substring(0,2)
        $dmcNode.dmCode.disassyCodeVariant = [string] $dmArray[6].Substring(2,3)
        $dmcNode.dmCode.infoCode = [string] $dmArray[7].Substring(0,3)
        $dmcNode.dmCode.infoCodeVariant = [string] $dmArray[7].Substring(3,1)
        $dmcNode.dmCode.itemLocationCode  = [string] $dmArray[8].Substring(0,1)

        return $dmcNode
}

Function Submit-LogEntry
{
    #Logging        : 63000
    Param([string] $fullLogPath , [string] $Message , [int] $EventID , [string] $evtType = "Information"  , [string] $caller)
    if(!([System.Diagnostics.EventLog]::SourceExists('KC46DataManagement')))
    {
        New-EventLog �LogName Application �Source "KC46DataManagement" -ErrorAction Stop
    }
    $Message = "Time:" + (Get-Date) +"`n`t EntryType:`t $evtType`n `tEventID:``t $EventID `n Caller:`t $caller`n$Message"   
    
    Out-File -FilePath $fullLogPath -Append -InputObject $Message    
}

Function Copy-FolderContents
{
    Param(  [string] $dest , [string] $source )
    try
    {
        Robocopy.exe /e /nc /ns /np /njs /njh /ndl /nfl  $source  $dest     
    }
    catch
    {
        
    }     
}

#Export-ModuleMember -Function Submit-LogEntry,Add-CEERSMessage