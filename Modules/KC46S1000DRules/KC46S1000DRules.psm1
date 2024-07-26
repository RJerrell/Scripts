FUNCTION Set-KC46BusinessRules_wf
{
Param([string] $source_BaseLocation, [string[]] $PubList)
    foreach($pub in $PubList | Sort-Object)
    {
        $pub
        #$path = "$source_BaseLocation\DMC*.xml"    
        $path1 = "$source_BaseLocation\$pub\S1000D"
        $path2 = "$source_BaseLocation\$pub\S1000D\SDLLIVE"

        if($pub -eq "IPB")
        {
            #parallel
            #{
                Set-KC46BusinessRules_IPB -path $path1
                Set-KC46BusinessRules_IPB -path $path2
            #}
        }
        if($pub -ne "IPB")
        {
            #parallel
            #{
                Set-KC46BusinessRules -path $path1
                Set-KC46BusinessRules -path $path2
            #}
        }    
    }    
}

function Set-KC46BusinessRules_IPB
{
Param([string] $path)
    
    $files = gci -Path $path -Filter dmc*.xml
    
    foreach($file in $files)
    {
        $dirty = $false
        $fullName = ""
        $oName = $file.FullName

        if(! $file.Name.Contains("_"))
        {
            $dmDoc = New-Object System.Xml.XmlDocument
            $dmDoc.Load($file.FullName)
            $issueNum = $dmDoc.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.issueNumber
            $inWork= $dmDoc.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.inWork
            $cCode = $dmDoc.dmodule.identAndStatusSection.dmAddress.dmIdent.language.countryIsoCode
            $languageIsoCode = $dmDoc.dmodule.identAndStatusSection.dmAddress.dmIdent.language.languageIsoCode

            $fullName = $file.FullName.ToUpper().Replace(".XML","")
            $fullName = $fullName +"_" + $issueNum + "-" + $inWork + "_" + $languageIsoCode.ToUpper() + "-" + $cCode + ".XML"
            $dirty = $true
        }
        $sreader = $null

        # Load the document as a string only -- THIS IS NOT AN XML DOCUMENT
        $sreader = New-Object System.IO.StreamReader($file.FullName) 
        $xdoc = $sreader.ReadToEnd() # Comes in as just a string
        $sreader.Close()
        $sreader.Dispose()
 
        if($xdoc.Contains("applicRefIds"))
        {           
            $xdoc = $xdoc.Replace("applicRefIds", "applicRefId")
            $dirty = $true
        }
        if($fullName -eq "")
        {         
            $fullName = $file.FullName
        }
        
        "Saving : `t" + $fullName

        #attrib -R $file.FullName;
        Set-ItemProperty -Path $file.FullName -Name IsReadOnly -Value $false -Force            
        if($dirty)
        {
            try
            {
                $newXdoc = New-Object System.Xml.XmlDocument
                $newXdoc.LoadXml($xdoc)
                Save-PrettyXML -FName $fullName -xmlDoc $newXdoc
                if($oName -ne $fullName)
                {
                    Remove-Item -Path $file.Fullname -Force -Verbose
                }
            }
            catch
            {
                "Failed to save $fullName "
                Exit
            }                  
        }
    }
    Set-KC46BusinessRules -path $path
}

function Set-KC46BusinessRules
{
Param([string] $path )
    
    # Roll through each file and replace "Safety Tag" with "Warning / Safety Tag"

    $files = gci -Path $path -Filter dmc*.xml -Recurse
    $xmlDoc = $null
    
    foreach ($file in $files)
    {
        <#
            Boeing KC46 Business Rules Compliance functionality.

            - Set the //originator//enterpriseName[.="The Boeing Company"] attribute value to "The Boeing Company"
            - Put an id attribute on all tables within each data module and sequentially number the id values within the data module
            - Safety tag changes from safety tag to 'Warning / Safety Tags in all para's, notepara, and WarningandCustionpara tags
            - ESDS symbols added to all paras
        #>
        $fullName = ""
        if(! $file.Name.Contains("_"))
        {
            $dmDoc = New-Object System.Xml.XmlDocument
            $dmDoc.Load($file.FullName)
            $issueNum = $dmDoc.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.issueNumber
            $inWork= $dmDoc.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.inWork
            $cCode = $dmDoc.dmodule.identAndStatusSection.dmAddress.dmIdent.language.countryIsoCode
            $languageIsoCode = $dmDoc.dmodule.identAndStatusSection.dmAddress.dmIdent.language.languageIsoCode

            $fullName = $file.FullName.ToUpper().Replace(".XML","")
            $fullName = $fullName +"_" + $issueNum + "-" + $inWork + "_" + $languageIsoCode.ToUpper() + "-" + $cCode + ".XML"
        }
        if($fullName.Length -eq 0)
        {
            $fullName = $file.FullName
        }

        $file.Attributes = "Archive"

		$dirty = $false
        # Regulatory tags and "this to that" content management
        $sreader = New-Object System.IO.StreamReader($fullName)
        $fileContent = $sreader.ReadToEnd()
        $sreader.Close()
        $sreader.Dispose()
        $x = $fileContent.Contains("DO-NOT-OPERATE")
        $y = $fileContent -contains "DO-NOT-OPERATE"

        $hasRegualtoryContent = $fileContent.Contains( "<emphasis emphasisType=`"em52`">")
        $hasRegualtorySymbols = $fileContent.Contains("ICN-81205-KZ90000002-001-01")
        $swdirty = $false
        
        # Do-Not-Operate tag fixes
        if($x -or $y)
        {
            $fileContent = $fileContent.Replace('DO-NOT-OPERATE' , 'Warning / Safety')
            $fileContent = $fileContent.Replace('do-not-operate' , 'Warning / Safety')
            $swdirty = $true
        }
        
        # Add symbols before each regualtory tag to be in sync with the PDF -- not a true business rule but required nonetheless
        if($hasRegualtoryContent -and (!$hasRegualtorySymbols))
        {
            $fileContent = $fileContent.Replace("<emphasis emphasisType=`"em52`">","<symbol infoEntityIdent=`"ICN-81205-KZ90000002-001-01`" /><emphasis emphasisType=`"em52`">")
            $swdirty = $true
        }
        if($swdirty -or $x -or $y)
        {            
            "Saving -->: " +  $fullName   
            $swriter = New-Object System.IO.StreamWriter($fullName)
            $swriter.AutoFlush = $true
            $swriter.Write($fileContent)
            $swriter.Close()
            $swriter.Dispose()
        }




        [System.XML.XMLDocument] $xdoc = New-Object System.XML.XMLDocument
        $xdoc.Load($fullName)

   
        [System.Xml.XmlLinkedNode[]]$paras =  [System.Xml.XmlLinkedNode[]] $xdoc.SelectNodes("//para")
        [System.Xml.XmlNodeList]$paras_N =  [System.Xml.XmlNodeList] $xdoc.SelectNodes("//notePara")
        [System.Xml.XmlNodeList]$paras_WC =  [System.Xml.XmlNodeList] $xdoc.SelectNodes("//warningAndCautionPara")
        
        $tableWithoutAnIDAttributes = $xdoc.SelectNodes("//table[not(@id)]")
        $leveledParaWithoutAnIDAttributes = $xdoc.SelectNodes("//levelledPara[not(@id)]")
        $proceduralStepWithoutAnIDAttributes = $xdoc.SelectNodes("//proceduralStep[not(@id)]")
        $crewDrillWithoutAnIDAttributes = $xdoc.SelectNodes("//crewDrill[not(@id)]")
        $multimediaObjectWithoutAnIDAttributes = $xdoc.SelectNodes("//multimediaObject[not(@id)]")
        $languageElements = $xdoc.SelectNodes("//dmRefIdent//language")

        #region Manage ID attrributes
        if($xdoc.dmodule.identAndStatusSection.dmStatus.originator.enterpriseName -eq $null)
        {
            # $fullName
            $newNode = $xdoc.CreateElement("enterpriseName")
            $newNode.InnerText = "The Boeing Company"
            $origNode = $xdoc.SelectSingleNode("/dmodule/identAndStatusSection/dmStatus/originator")                                                
            $origNode.AppendChild($newNode)
            $xdoc.dmodule.identAndStatusSection.dmStatus.originator.enterpriseName
            $dirty = $true
        }
        elseif($xdoc.dmodule.identAndStatusSection.dmStatus.originator.enterpriseName -ne "The Boeing Company")
        {
            $xdoc.dmodule.identAndStatusSection.dmStatus.originator.enterpriseName = "The Boeing Company"
            $xdoc.dmodule.identAndStatusSection.dmStatus.originator.enterpriseName
            $dirty = $true
        }
        foreach ($le in $languageElements)
        {            
            $le.ParentNode.RemoveChild($le)
            $dirty = $true
        }
        foreach ($node in $tableWithoutAnIDAttributes)
        {
            $Id = "KC46-" + [GUID]::NewGuid()
            $node.SetAttribute("id",$Id)
            #$node.id
            $dirty = $true
        }
        foreach ($leveledPara in $leveledParaWithoutAnIDAttributes)
        {
            $Id = "KC46-" + [GUID]::NewGuid()
            $leveledPara.SetAttribute("id",$Id)
            #$leveledPara.id
            $dirty = $true
        }
        foreach ($proceduralStep in $proceduralStepWithoutAnIDAttributes)
        {
            $Id = "KC46-" + [GUID]::NewGuid()
            $proceduralStep.SetAttribute("id",$Id)
            #$proceduralStep.id
            $dirty = $true
        }
        foreach ($node in $crewDrillWithoutAnIDAttributes)
        {
            $Id = "KC46-" + [GUID]::NewGuid()
            $crewDrill.SetAttribute("id",$Id)
            #$crewDrill.id
            $dirty = $true
        }                
        foreach ($node in $multimediaObjectWithoutAnIDAttributes)
        {
            $Id = "KC46-" + [GUID]::NewGuid()
            $multimediaObject.SetAttribute("id",$Id)
            #$multimediaObject.id
            $dirty = $true
        }        

        #endregion
        
        #region Safety tag changes from safety tag to 'Warning / Safety Tags in all para's, notepara, and WarningandCustionpara tags
        if( ! $xdoc.dmodule.InnerText.Contains('Warning / Safety Tag'))
        {            
            foreach($para in $paras)
            {
                if($para.InnerText.Contains('safety tag'))
                {
                    $para.InnerText = $para.InnerText.Replace('safety tag' , 'Warning / Safety Tag')
                    $dirty = $true
                }
                if($para.InnerText.Contains('safety tags'))
                {
                    $para.InnerText = $para.InnerText.Replace('safety tags' , 'Warning / Safety Tags')
                    $dirty = $true
                }
            }
        }

        #endregion
        #region - ESDS symbols added to all paras 
        <# 
            Linh Tang requested items
            •	Electrostatic Sensitive Devises
            •	Electrostatic Sensitive Components
            •	Electrical Static Discharge
            •	Static Electricity
        #>
        [bool] $esdsProcessingRequired = $false
        [bool] $esdsProcessingRequired = $xdoc.OuterXml.Contains('infoEntityIdent="ICN-81205-KZ90000001-001-01"')
        if($esdsProcessingRequired -ne $true)
        {               
            <# Electrostatic Sensitive Devises , Electrostatic Sensitive Components, Electrical Static Discharge, Static Electricity #>
            
            # Just a normal para tag
            foreach( $p in $paras)
            {  
                if(!($p.FirstChild.Name -and $p.LastChild.Name) -eq "dmRef")
                {
                    # This test is case-insensitive
                    if( $p.InnerXml.ToString().Contains("ESDS")   `
                        -or $p.InnerXml.ToString().Contains("ELECTROSTATIC DISCHARGE") `
                        -or $p.InnerXml.ToString().Contains("electrostatic discharge") `
                        -or $p.InnerXml.ToString().Contains("ELECTROSTATIC SENSITIVE")  `
                        -or $p.InnerXml.ToString().Contains("Electrical Static Discharge")  `
                        -or $p.InnerXml.ToString().Contains("static electricity")  `
                        )
                    {
                        #"Para tag updated"
                        $p.InnerXml = $p.InnerXml.ToString().Replace("ESDS", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> ESDS")
                        $p.InnerXml = $p.InnerXml.ToString().Replace("ELECTROSTATIC DISCHARGE", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> ELECTROSTATIC DISCHARGE")
                        $p.InnerXml = $p.InnerXml.ToString().Replace("electrostatic discharge", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> electrostatic discharge")
                        $p.InnerXml = $p.InnerXml.ToString().Replace("ELECTROSTATIC SENSITIVE", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> ELECTROSTATIC SENSITIVE")     
                        $p.InnerXml = $p.InnerXml.ToString().Replace("Electrostatic Sensitive Device", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> Electrostatic Sensitive Device")
                        $p.InnerXml = $p.InnerXml.ToString().Replace("Electrical Static Discharge", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> Electrical Static Discharge")                  
                        $p.InnerXml = $p.InnerXml.ToString().Replace("static electricity", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> static electricity")     
                        $p.InnerXml = $p.InnerXml.ToString().Replace("STATIC ELECTRICITY", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> STATIC ELECTRICITY") 
                        $p.InnerXml = $p.InnerXml.ToString().Replace("<symbol infoEntityIdent=`"ICN-81205-KZ90000001-001-01`" /><symbol infoEntityIdent=`"ICN-81205-KZ90000001-001-01`" />", "<symbol infoEntityIdent=`"ICN-81205-KZ90000001-001-01`" />")
                        $dirty = $true 
                    }

                    # Change all the instances of Do-Not-Operate to WARNING - CASE SENSITIVE so check for all lower, all upper, and Camel case
                    # This code satisfies IPR comment # 2630 from Harold Doubet and Brian Johnson in Arpil 2017 from the 90% IPR
                    if( $p.InnerXml.ToString().Contains("Do-Not-Operate") -or $p.InnerXml.ToString().Contains("DO-NOT-OPERATE") -or $p.InnerXml.ToString().Contains("do-not-operate"))
                    {
                        $p.InnerXml = $p.InnerXml.ToString().Replace("Do-Not-Operate", "Warning")
                        $p.InnerXml = $p.InnerXml.ToString().Replace("DO-NOT-OPERATE", "WARNING")
                        $p.InnerXml = $p.InnerXml.ToString().Replace("do-not-operate", "warning")
                    }
                }
            }
            
            #Note
            foreach( $pn in $paras_N)
            {  # This test is case-insensitive
                if( $pn.InnerXml.ToString().Contains("ESDS")  `
                    -or $pn.InnerXml.ToString().Contains("ELECTROSTATIC DISCHARGE") `
                    -or $pn.InnerXml.ToString().Contains("electrostatic discharge") `
                    -or $pn.InnerXml.ToString().Contains("ELECTROSTATIC SENSITIVE")  `
                    -or $pn.InnerXml.ToString().Contains("Electrical Static Discharge")  `
                    -or $pn.InnerXml.ToString().Contains("static electricity")  `
                    )
                {
                    #"notePara tag updated"
                    $pn.InnerXml = $pn.InnerXml.ToString().Replace("ESDS", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> ESDS")                    
                    $pn.InnerXml = $pn.InnerXml.ToString().Replace("ELECTROSTATIC DISCHARGE", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> ELECTROSTATIC DISCHARGE")
                    $pn.InnerXml = $pn.InnerXml.ToString().Replace("electrostatic discharge ", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> electrostatic discharge ")
                    $pn.InnerXml = $pn.InnerXml.ToString().Replace("ELECTROSTATIC SENSITIVE ", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> ELECTROSTATIC SENSITIVE ")     
                    $pn.InnerXml = $pn.InnerXml.ToString().Replace("Electrostatic Sensitive Device ", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> Electrostatic Sensitive Device ")
                    $pn.InnerXml = $pn.InnerXml.ToString().Replace("Electrical Static Discharge ", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> Electrical Static Discharge ")                         
                    $pn.InnerXml = $pn.InnerXml.ToString().Replace("electrical static discharge", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> electrical static discharge")
                    $pn.InnerXml = $pn.InnerXml.ToString().Replace("static electricity", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> static electricity")     
                    $pn.InnerXml = $pn.InnerXml.ToString().Replace("STATIC ELECTRICITY", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> STATIC ELECTRICITY")
                    $pn.InnerXml = $pn.InnerXml.ToString().Replace("<symbol infoEntityIdent=`"ICN-81205-KZ90000001-001-01`" /><symbol infoEntityIdent=`"ICN-81205-KZ90000001-001-01`" />", "<symbol infoEntityIdent=`"ICN-81205-KZ90000001-001-01`" />")
                    $dirty = $true 
                }
            }

            # Warnings and Cautions (warningAndCautionPara)
            foreach( $pwc in $paras_WC)
            {                
               # This test is case-insensitive
                if( $pwc.InnerXml.ToString().Contains("ESDS") `
                    -or $pwc.InnerXml.ToString().Contains("ELECTROSTATIC DISCHARGE") `
                    -or $pwc.InnerXml.ToString().Contains("electrostatic discharge") `
                    -or $pwc.InnerXml.ToString().Contains("ELECTROSTATIC SENSITIVE")  `
                    -or $pwc.InnerXml.ToString().Contains("Electrical Static Discharge")  `
                    -or $pwc.InnerXml.ToString().Contains("static electricity")  `
                    )
                {
                    # warningAndCautionPara content updated
                    $pwc.InnerXml = $pwc.InnerXml.ToString().Replace("ESDS", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> ESDS")
                    $pwc.InnerXml = $pwc.InnerXml.ToString().Replace("ELECTROSTATIC DISCHARGE", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> ELECTROSTATIC DISCHARGE")
                    $pwc.InnerXml = $pwc.InnerXml.ToString().Replace("electrostatic discharge", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> ELECTROSTATIC SENSITIVE")     
                    $pwc.InnerXml = $pwc.InnerXml.ToString().Replace("electrical static discharge", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> electrical static discharge")
                    $pwc.InnerXml = $pwc.InnerXml.ToString().Replace("ELECTROSTATIC SENSITIVE", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> electrostatic discharge")     
                    $pwc.InnerXml = $pwc.InnerXml.ToString().Replace("Electrostatic Sensitive Device ", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> Electrostatic Sensitive Device ")                     
                    $pwc.InnerXml = $pwc.InnerXml.ToString().Replace("static electricity", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> static electricity")
                    $pwc.InnerXml = $pwc.InnerXml.ToString().Replace("STATIC ELECTRICITY", "<symbol infoEntityIdent=""ICN-81205-KZ90000001-001-01"" /> STATIC ELECTRICITY")
                    $pwc.InnerXml = $pwc.InnerXml.ToString().Replace("<symbol infoEntityIdent=`"ICN-81205-KZ90000001-001-01`" /><symbol infoEntityIdent=`"ICN-81205-KZ90000001-001-01`" />", "<symbol infoEntityIdent=`"ICN-81205-KZ90000001-001-01`" />")
                    $dirty = $true 
                }
            }
        }
        #endregion

        if($dirty -eq $true)
        {   
            "Saving : `t" + $file.FullName
            $file.Attributes = "Archive"
            #"Saving : `t" + $file.FullName  
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
            $settings.CheckCharacters = $true
            $settings.DoNotEscapeUriAttributes = $false

            #Create an XmlWriter to insure the output XML conforms to the settings above
            $writer = [System.Xml.XmlWriter]::Create($file.FullName,$settings)
            # "Saving : `t" + $file.FullName 
            $xdoc.Save($writer);
            $writer.Flush();
            $writer.Close();
            $writer.Dispose();
        }
        $dirty = $false
    }
}

function Start-AugmentationProcess
{
Param( [string] $augmentorPath , [string] $augmentorFileName)
    try
    {
        Start-Process -FilePath "$augmentorPath\$augmentorFileName" -Wait -WindowStyle Minimized -ErrorAction Stop| Out-Null
    }
   
    catch [System.Net.WebException],[System.Exception]
    {
        Write-Host "Other exception"
    }
    finally
    {
        Write-Host "cleaning up ..."
    }    
}