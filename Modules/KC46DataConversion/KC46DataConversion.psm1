#region WorkFlows
#endregion

#region Functions
Function Get-SWPM_CRI
{   
    $cri = @"
    <pmEntry>
	<pmEntryTitle>Cross Reference Index</pmEntryTitle>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-0000-00A0U-013A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="0000" disassyCode="00" disassyCodeVariant="A0U" infoCode="013" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-0001-00A0U-013A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="0001" disassyCode="00" disassyCodeVariant="A0U" infoCode="013" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-0002-00A0U-013A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="0002" disassyCode="00" disassyCodeVariant="A0U" infoCode="013" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-0003-00A0U-013A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="0003" disassyCode="00" disassyCodeVariant="A0U" infoCode="013" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-0004-00A0U-013A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="0004" disassyCode="00" disassyCodeVariant="A0U" infoCode="013" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-0005-00A0U-013A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="0005" disassyCode="00" disassyCodeVariant="A0U" infoCode="013" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-0006-00A0U-013A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="0006" disassyCode="00" disassyCodeVariant="A0U" infoCode="013" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-0007-00A0U-013A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="0007" disassyCode="00" disassyCodeVariant="A0U" infoCode="013" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-0008-00A0U-013A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="0008" disassyCode="00" disassyCodeVariant="A0U" infoCode="013" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-0009-00A0U-013A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="0009" disassyCode="00" disassyCodeVariant="A0U" infoCode="013" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000A-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000A" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000B-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000B" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000C-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000C" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000D-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000D" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000E-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000E" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000F-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000F" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000G-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000G" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000H-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000H" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000I-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000I" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000J-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000J" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000K-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000K" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000L-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000L" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000M-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000M" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000N-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000N" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000O-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000O" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000P-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000P" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000Q-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000Q" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000R-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000R" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000S-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000S" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000T-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000T" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000U-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000U" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000V-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000V" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000W-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000W" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000X-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000X" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000Y-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000Y" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
	<dmRef>
		<dmRefIdent>
			<!--  DMC:	DMC-1KC46-A-20-00-000Z-00A0U-014A-A -->
			<dmCode modelIdentCode="1KC46" systemDiffCode="A" systemCode="20" subSystemCode="0" subSubSystemCode="0" assyCode="000Z" disassyCode="00" disassyCodeVariant="A0U" infoCode="014" infoCodeVariant="A" itemLocationCode="A" />
			<language countryIsoCode="US" languageIsoCode="sx" />
		</dmRefIdent>
	</dmRef>
</pmEntry>
"@
return $cri
}
Function Assert-RandomList
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType , [string] $listBulletType = "pf02" ) 
    $openTagTxt = "<randomList listItemPrefix=`"$listBulletType`">"
    $closeTagTxt =     "</randomList>"
    $childNodes = $EL2Convert.ChildNodes        
    if($ELType -eq "HWINFO")
    {
        $txt = $EL2Convert.PNR
        $childnode_text_out = $childnode_text_out + "<listItem><para>$txt</para></listItem>"
    }
    elseif($ELType -eq "TERMINFO")
    {
       $mfrName =  $EL2Convert.MFRNAME
       $termpnr = $EL2Convert.TERMPNR
       $termname = $EL2Convert.TERMNAME
       $childnode_text_out = $childnode_text_out + "<listItem><para>$mfrName</para></listItem>"
       $childnode_text_out = $childnode_text_out + "<listItem><para>$termpnr</para></listItem>"
       $childnode_text_out = $childnode_text_out + "<listItem><para>$termname</para></listItem>"
    }
    else
    {    
        foreach ($childNode in $childNodes)
        {
            $txt = $childNode."#text"
            $childnode_text_out = $childnode_text_out + "<listItem><para>$txt</para></listItem>"
        }  
    }
    return $openTagTxt + $childnode_text_out + $closeTagTxt
}
Function Assert-WCN
{
	Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType)
	switch ($EL2Convert.Name.ToUpper())
	{
		'CAUTION' {
			$openTagTxt = "<caution>"
			$closeTagTxt = "</caution>"
			break
		}        
		'WARNING' {
			$openTagTxt = "<warning>"
			$closeTagTxt = "</warning>"
			break
		}    
	}
	$childnode_text_out = $EL2Convert."text"
    $childNodes = $EL2Convert.ChildNodes     
	foreach ($childNode in $childNodes)
	{        
	
        if ($childNode.Name -eq "PARA")
        {
            $childnode_text_out +=   ConvertTo-S1000D4_0_1 -EL2Convert $childNode -ELType $childNode.Name
        
            $childnode_text_out = $childnode_text_out.replace("<para>", "<warningAndCautionPara>")
            $childnode_text_out = $childnode_text_out.replace("</para>", "</warningAndCautionPara>")
        }    
        
        if ($childNode.Name -eq "UNLIST")
        {
            $childnode_text_out +=   ConvertTo-S1000D4_0_1 -EL2Convert $childNode -ELType $childNode.Name

            $childnode_text_out = $childnode_text_out.replace("<para><randomList", "<warningAndCautionPara><randomList")
            $childnode_text_out = $childnode_text_out.replace("</randomList></para>", "</randomList></warningAndCautionPara>")

            $childnode_text_out = $childnode_text_out.replace("<randomList", "<attentionRandomList")
            $childnode_text_out = $childnode_text_out.replace("</randomList>", "</attentionRandomList>")

            $childnode_text_out = $childnode_text_out.replace("<listItem>", "<attentionRandomListItem>")
            $childnode_text_out = $childnode_text_out.replace("</listItem>", "</attentionRandomListItem>")

            $childnode_text_out = $childnode_text_out.replace("<para>", "<attentionListItemPara>")
            $childnode_text_out = $childnode_text_out.replace("</para>", "</attentionListItemPara>")

        }  
        
	}
	return $openTagTxt + $childnode_text_out + $closeTagTxt
}
Function ConvertFrom-GRAPHIC
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType , [boolean] $createIRef = $false ) 
    $figure_id = $EL2Convert.KEY
    $childNodes = $EL2Convert.ChildNodes
    if($createIRef)
    {
        $childnode_text_out = "<internalRef internalRefTargetType=`"figure`" internalRefId=`"$figure_id`" /><figure id=`"$figure_id`" >"
    }
    else
    {
        $childnode_text_out = "<figure id=`"$figure_id`" >"
    }
    foreach ($childNode in $childNodes)
    {
        #echo "*** Now we are in GRAPHIC ChildNode!"
        if ($childNode.Name -eq "TITLE")
        {
            #$str = ($childNode."#text")
            $str = $childNode.InnerText        # Clean up bad data
            if ($str -match "^Figure \d+ - (.*?)$")
            {
                $childnode_text_out += "<title>" + $matches[1] + "</title>"
            }
            else
            {
                $childnode_text_out += "<title>" + $childNode."#text" + "</title>"
            }
        }
        if ($childNode.Name -eq "SHEET")
        {
            $GNBR =  [string] $childNode.GNBR
            $graphic_icn_name = [string] ($global:keyICNHash.GetEnumerator() | ? {$_.Name -eq $GNBR }).Value
            
            if($graphic_icn_name.Length -lt 1 )
            {
               for ($i = 0; $i -lt $global:keyICNHash.Count; $i++)
               { 
                    $x = $global:keyICNHash[$i]
                    if($x.Keys[0].ToString() -eq $GNBR)
                    {
                        $graphic_icn_name = $x.Values[0]
                    }
                }
            }
            $childnode_text_out += "<graphic id=`"$GNBR`" infoEntityIdent=`"$graphic_icn_name`" />"             
        }
    }
    return $childnode_text_out + "</figure>"
}
Function ConvertFrom-GRPHCREF
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType)
    $openTagTxt = "<internalRef"   # THE END > CHARACTER PURPOSELY LEFT OFF THE END OF THE TAG
    $refID  = $EL2Convert.REFID
    $childnode_text_out = "$openTagTxt internalRefTargetType=`"figure`" internalRefId=`"" + $refID + "`"/>" 
    return $childnode_text_out
}
Function ConvertFrom-LIST
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType)

    $nodeName = $EL2Convert.Name
    $nodeType = $ELType
    $childnode_text_out = ""
    #$childnode_text_out += ConvertFrom-TITLE -EL2Convert $EL2Convert.TITLE -ELType $EL2Convert.Name
    $childNodes = $EL2Convert.ChildNodes
        
    foreach ($childNode in $childNodes)
    {   
        if($childNode.Name.ToUpper() -in "L1ITEM" ,"L2ITEM" , "L3ITEM","L4ITEM","L5ITEM","L6ITEM","L7ITEM" )
        {
            $openTagTxt = "<levelledPara>"
	        $closeTagTxt =     "</levelledPara>"
            $childnode_text_out +=   ConvertFrom-LISTITEM $childNode -ELType $childNode.Name
        }
        else
        {
            $childnode_text_out +=   ConvertTo-S1000D4_0_1 -EL2Convert $childNode -ELType $childNode.Name
        }
    }
    
    return $childnode_text_out

}
Function ConvertFrom-LISTITEM
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType)    
    $nodeName = $EL2Convert.Name
    $nodeType = $ELType
    if($EL2Convert.KEY.Length -gt 0)
    {
        $openTagTxt = "<levelledPara id=`"" + $EL2Convert.KEY + "`">" 
    }
    else
    {
        $openTagTxt = "<levelledPara>"
    }
    $closeTagTxt =     "</levelledPara>"

    $childnode_text_out = ""
    $childNodes = $EL2Convert.ChildNodes
    foreach ($childNode in $childNodes)
    {
        if( $childNode.Name.ToUpper() -in "LIST1" ,"LIST2" , "LIST3","LIST4","LIST5","LIST6","LIST7" )    
        {
            $childnode_text_out +=   ConvertFrom-LIST $childNode -ELType $childNode.Name
        }
        else
        {
            $childnode_text_out +=   ConvertTo-S1000D4_0_1 -EL2Convert $childNode -ELType $childNode.Name
        }
    }
    return $openTagTxt + $childnode_text_out + $closeTagTxt
}
Function ConvertFrom-NOTE
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType)
    $childnode_text_out = "" 
    $childNodes = $EL2Convert.ChildNodes
    if ($ELType.ToUpper() -eq "NOTE")
    {
        #$openTagTxt = "<notePara><attentionRandomList>"
        #$closeTagTxt = "</attentionRandomList></notePara>"

        $openTagTxt = "<note>"
        $closeTagTxt = "</note>"
    }
    else
    {
        #$openTagTxt = "<warningAndCautionPara><attentionRandomList>"
        #$closeTagTxt = "</attentionRandomList></warningAndCautionPara>"

        $openTagTxt = "<warningAndCautionPara>"
        $closeTagTxt = "</warningAndCautionPara>"

    }

    foreach($childNode in $childNodes)
    {
        if ($childNode.Name -eq "PARA")
        {
            $childnode_text_out +=   ConvertTo-S1000D4_0_1 -EL2Convert $childNode -ELType $childNode.Name
        
            $childnode_text_out = $childnode_text_out.replace("<para>", "<notePara>")
            $childnode_text_out = $childnode_text_out.replace("</para>", "</notePara>")
        }    
        
        if ($childNode.Name -eq "UNLIST")
        {

            $childnode_text_out +=   ConvertTo-S1000D4_0_1 -EL2Convert $childNode -ELType $childNode.Name

            $childnode_text_out = $childnode_text_out.replace("<para><randomList", "<notePara><randomList")
            $childnode_text_out = $childnode_text_out.replace("</randomList></para>", "</randomList></notePara>")

            $childnode_text_out = $childnode_text_out.replace("<randomList", "<attentionRandomList")
            $childnode_text_out = $childnode_text_out.replace("</randomList>", "</attentionRandomList>")

            $childnode_text_out = $childnode_text_out.replace("<listItem>", "<attentionRandomListItem>")
            $childnode_text_out = $childnode_text_out.replace("</listItem>", "</attentionRandomListItem>")

            $childnode_text_out = $childnode_text_out.replace("<para>", "<attentionListItemPara>")
            $childnode_text_out = $childnode_text_out.replace("</para>", "</attentionListItemPara>")


        }

    
    }
    return $openTagTxt + $childnode_text_out + $closeTagTxt
}
Function ConvertFrom-PARA
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType)
    if($EL2Convert.KEY.Length -gt 0)
    {
        $openTagTxt = "<para id=`"" + $EL2Convert.KEY + "`">" 
    }
    else
    {
        $openTagTxt = "<para>"
    }
    $closeTagTxt =     "</para>"
    $childnode_text_out = ""  
  
    foreach ($childNode in $EL2Convert.ChildNodes)
    {
       # if ($childNode.NodeType -eq "TEXT")
        $para_text = ""

        if ($childNode.Name -eq "#text")
        {            
            $para_text = $childNode.value
            $childnode_text_out = $childnode_text_out + $para_text
        }
        elseif($childNode.TERMINFO.InnerText.Length -gt 0)
        {
            $childnode_text_out +=  ConvertTo-S1000D4_0_1 -EL2Convert $childNode.TERMINFO -ELType "TERMINFO"
        }
        elseif($childNode.Name -eq "PARA")
        {
            $para_text = $childNode.'#text'
            $childnode_text_out = $childnode_text_out + $para_text
        }

        else
        {           
            $childnode_text_out +=  ConvertTo-S1000D4_0_1 -EL2Convert $childNode -ELType $childNode.Name
        }
    }
     return $openTagTxt + $childnode_text_out + $closeTagTxt
}
Function ConvertFrom-REFINT
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType)    
    
    $openTagTxt = "<internalRef"   # THE END > CHARACTER PURPOSELY LEFT OFF THE END OF THE TAG
    $refID  = $EL2Convert.REFID
    if($EL2Convert.REFTYPE.ToUpper() -eq "SPSUBJ")
    {
        $openTagTxt = "<dmRef>"
        $closeTagTxt = "</dmRef>"
        $childnode_text_out = ""
        $dmcFileName =  ($KeyToFile.GetEnumerator() | ? { $_.KEY -eq $refID }).Value
        $refint_systemCode = $dmcFileName.SubString(12,2)
        $refint_subSystemCode = $dmcFileName.SubString(15,1)
        $refint_subSubSystemCode = $dmcFileName.SubString(16,1) 
        $refint_assyCode = $dmcFileName.SubString(18,4)
        $refint_infoCode = $dmcFileName.SubString(29,3)
        $childnode_text_out = "<dmRef><dmRefIdent><dmCode id=`"$refID`" modelIdentCode=`"1KC46`" systemDiffCode=`"A`" systemCode=`"$refint_systemCode`" subSystemCode=`"$refint_subSystemCode`" subSubSystemCode=`"$refint_subSubSystemCode`" assyCode=`"$refint_assyCode`" disassyCode=`"00`" disassyCodeVariant=`"A0U`" infoCode=`"$refint_infoCode`" infoCodeVariant=`"A`" itemLocationCode=`"A`" /></dmRefIdent></dmRef>"
        return $childnode_text_out
    }
    elseif($EL2Convert.REFTYPE.ToUpper() -eq "TABLE")
    {
        $childnode_text_out = "$openTagTxt internalRefTargetType=`"table`" internalRefId=`"" + $refID + "`"/>" 
    }
    elseif($EL2Convert.REFTYPE.ToUpper() -eq "L1ITEM" -OR $EL2Convert.REFTYPE.ToUpper() -eq "L2ITEM" -OR $EL2Convert.REFTYPE.ToUpper() -eq "L3ITEM" -OR $EL2Convert.REFTYPE.ToUpper() -eq "L4ITEM" -OR $EL2Convert.REFTYPE.ToUpper() -eq "L5ITEM" -OR $EL2Convert.REFTYPE.ToUpper() -eq "L6ITEM" -OR $EL2Convert.REFTYPE.ToUpper() -eq "L7ITEM" )
    {
         $childnode_text_out = "$openTagTxt internalRefTargetType=`"para`" internalRefId=`"" + $refID + "`"/>" 
    }

    elseif($EL2Convert.REFTYPE.ToUpper() -eq "GRAPHIC")
    {
       $childnode_text_out = "$openTagTxt internalRefTargetType=`"figure`" internalRefId=`"" + $refID + "`"/>" 
    }    
    elseif($EL2Convert.REFTYPE.ToUpper() -in {"L1ITEM","L2ITEM", "L3ITEM", "L4ITEM", "L5ITEM", "L6ITEM", "L7ITEM"})
    {
        $childnode_text_out = "$openTagTxt internalRefTargetType=`"para`" internalRefId=`"" + $refID + "`"/>" 
    }
    return $childnode_text_out
}
Function ConvertFrom-TABLE
{
Param([parameter(Mandatory=$true)] [Xml.XmlElement[]] $EL2Convert, [int] $startRowNum, [int] $endRowNum ,[ref] $morerows)
   
<#
        The 'maximumSupportedRowsPerTable' variable holds a value that determines how a table might be
        broken up into more table.  Some viewers cannot render large table and breaking them up into smaller
        tables seems to work.   Experiment with your viewer to determine the best value based on your content
    #>
    #$maximumSupportedRowsPerTable = 250 # rows

    $arrRow = @()    
    $i = [int] 0
	$tableFrame = ""
	$tableColsep = ""
	$tableRowsep = ""
	$tableOrient = ""
	$tablePgwide = ""
    $tableOrient = ""
    $tableFrame = ""
    try
	{
		$tableFrame = [string] $EL2Convert.FRAME.ToLower()
		$tableColsep = $EL2Convert.COLSEP
		$tableRowsep = $EL2Convert.ROWSEP
		$tableOrient = [string] $EL2Convert.ORIENT.ToLower()
		$tablePgwide = $EL2Convert.PGWIDE
        $tableOrient = $tableOrient.ToLower()
        $tableFrame = $tableFrame.ToLower()    
        }
	catch {}

    <# For KC46, WE'LL SET ALL THE TABLE WIDTHS TO FILL THE TASK VIEWPORT BY SETTING THE WIDTH TO 100 #>
    $tablePgwide = "0"

    $tableId = $EL2Convert.ID
    $tableTitle = $EL2Convert.Title
		cls
    $tableTitle	
    if($tableTitle -eq "CONNECTOR BACKSHELL TOOLS")
    {
        "stop"
    }
    #take out the Table ## text
    if ($tableTitle -match "^Table \d+ - (.*?)$")
    {
        $tableTitleText = $matches[1]
        $tableTitle = $tableTitleText
    }
    $tableTagTxt = "<table tabstyle=`"CALS`" frame=`"$tableFrame`" colsep=`"$tableColsep`" rowsep=`"$tableRowsep`" orient=`"$tableOrient`" pgwide=`"$tablePgwide`" id=`"$tableId`">"
    if($tableTitle.length -lt 1)
    {
        $titleTagTxt = ""
        $titleTagTxt2 = ""
    }
    else
    {           
        $titleTagTxt = "<title>$tableTitle</title>"
        $titleTagTxt2 = "<title>$tableTitle (-continued)</title>"
    }
    

    if($startRowNum -gt 0)
    {
        $titleTagTxt = $titleTagTxt2
    }

    $tableTgroupCols = $EL2Convert.TGROUP.COLS
    $tableTgroupColsep = $EL2Convert.TGROUP.COLSEP
    $tableTgroupRowsep = $EL2Convert.TGROUP.ROWSEP
    $tableTgroupAlign = [string] $EL2Convert.TGROUP.ALIGN
    $tableTgroupCharoff = $EL2Convert.TGROUP.CHAROFF
    $tableTgroupChar = $EL2Convert.TGROUP.CHAR
    $tableTgroupAlign = $tableTgroupAlign.ToLower()

    if ($tableTgroupColsep.length -gt 0)
    {
        $tableTgroupColsepText = " colsep=`"$tableTgroupColsep`""
    }
    else
    {
        $tableTgroupColsepText = ""
    }


    if ($tableTgroupRowsep.length -gt 0)
    {
        $tableTgroupRowsepText = " rowsep=`"$tableTgroupRowsep`""
    }
    else
    {
        $tableTgroupRowsepText = ""
    }
    $tableTgroupColsepRowsepText = "$tableTgroupColsepText$tableTgroupRowsepText"
    $tableTgroupTagTxt = "<tgroup cols=`"$tableTgroupCols`"$tableTgroupColsepRowsepText align=`"$tableTgroupAlign`" charoff=`"$tableTgroupCharoff`" char=`"$tableTgroupChar`">"
    $colSpecTxt = "";
    $theadTxt = "";
    $tbodyTxt = "";
    $childNodes = $EL2Convert.TGROUP.ChildNodes
    foreach($childNode in $childNodes)
    {
        if ($childNode.Name -eq "COLSPEC")
        {              
            $colSpec_colnum = [string] $childNode.COLNUM
            $colSpec_colname = [string] $childNode.COLNAME
            $colSpec_align = [string] $childNode.ALIGN
            $colSpec_charoff = $childNode.CHAROFF
            $colSpec_char = $childNode.CHAR
            $colSpec_colwidth = $childNode.COLWIDTH                
            $colSpec_colnum = $colSpec_colnum.ToLower()
            $colSpec_colname = $colSpec_colname.ToLower()
            $colSpec_align = $colSpec_align.ToLower()
            $colSpec_colsep = $childNode.COLSEP;
            $colSpec_rowsep = $childNode.ROWSEP;
            $colSpecTxt = $colSpecTxt + "<colspec colnum=`"$colSpec_colnum`" colname=`"$colSpec_colname`" colwidth=`"$colSpec_colwidth`" align=`"$colSpec_align`" charoff=`"$colSpec_charoff`" char=`"$colSpec_char`" colsep=`"$colSpec_colsep`" rowsep=`"$colSpec_rowsep`" />"
        }
        elseif ($childNode.Name -eq "THEAD")
        {
            $theadTxt = "<thead>"
            
            foreach($childNode in $childNode.ChildNodes)
            {                       
                if ($childNode.Name -eq "ROW")
                {
                    $theadTxt = $theadTxt + "<row>"

                    foreach ($childNode in $childNode.ChildNodes)
                    {
                        if ($childNode.Name -eq "ENTRY")
                        {
                            try
                            {
                                $TRowEntry_colname = ""
                                $TRowEntry_valign = ""
                                $TRowEntry_align = ""
                                $TRowEntry_morerows = ""
                                $colSpec_namest = ""
                                $colSpec_nameend = ""

                                $TRowEntry_colname = [string] $childNode.COLNAME
                                $TRowEntry_valign = [string] $childNode.VALIGN
                                $TRowEntry_align = [string] $childNode.ALIGN
                                $TRowEntry_morerows = $childNode.MOREROWS
                                $colSpec_namest = [string] $childNode.NAMEST
                                $colSpec_nameend = [string]$childNode.NAMEEND

                                $TRowEntry_colname = $TRowEntry_colname.ToLower()
                                $TRowEntry_valign = $TRowEntry_valign.ToLower()
                                $TRowEntry_align = $TRowEntry_align.ToLower()
                                $TRowEntry_morerows = $TRowEntry_morerows.ToLower()
                                $colSpec_namest = $colSpec_namest.ToLower()
                                $colSpec_nameend = $colSpec_nameend.ToLower()
                            }
                            catch{}
                            $TRowEntry_para = $childNode.PARA
                            if ($colSpec_namest.length -gt 0)
                            {                                
                                $colSpec_namest = $colSpec_namest.ToLower()
                                $colSpec_nameend = $colSpec_nameend.ToLower()
                                $colSpanText = "namest=`"$colSpec_namest`" nameend=`"$colSpec_nameend`""
                            }
                            else
                            {
                                $colSpanText = ""
                            }

                            if ($TRowEntry_morerows.Length -gt 0)
                            {
                                $theadTxt = $theadTxt + "<entry colname=`"$TRowEntry_colname`" $colSpanText morerows=`"$TRowEntry_morerows`" valign=`"$TRowEntry_valign`" align=`"$TRowEntry_align`"><para>$TRowEntry_para</para></entry>"
                            }
                            else
                            {
                                $theadTxt = $theadTxt + "<entry colname=`"$TRowEntry_colname`" $colSpanText valign=`"$TRowEntry_valign`" align=`"$TRowEntry_align`"><para>$TRowEntry_para</para></entry>"
                            }                                       
                        } 
                    }
                    $theadTxt = $theadTxt + "</row>"
                }                        
            }
            $theadTxt = $theadTxt + "</thead>"
        }
        elseif ($childNode.Name -eq "TBODY")
        {
            if($childNode.Attributes.count -gt 0)
            {
                $tbodyValign = [string] ($childNode.Attributes[0].'#text').ToLower()
            }
            else
            {
                $tbodyValign = "left"
            }
            $tbodyTxt = "<tbody valign=`"$tbodyValign`">" 
            $tbodyRowTxt = ""
            $rowID = 0
            $rows = $EL2Convert.TGROUP.TBODY.ChildNodes
            $rowCount = $rows.Count
            for ($i = $startRowNum; $i -lt $endRowNum; $i++)
            {
                $tbodyRowTxt = "<row id=`"row-$rowID`">"
                $tbody_para_text = ""
                # process each entry in the row
                foreach ($childNode  in $rows[$i].ChildNodes)
                {
                    try
                    {
                        $tbodyRowEntry_colname = ""
                        $tbodyRowEntry_valign = ""
                        $tbodyRowEntry_align = ""
                        $tbodyRowEntry_morerows = ""

                        $tbodyRowEntry_colname = $childNode.COLNAME
                        $tbodyRowEntry_valign = [string]$childNode.VALIGN
                        $tbodyRowEntry_align = [string] $childNode.ALIGN
                        $tbodyRowEntry_morerows =$childNode.MOREROWS

                        $tbodyRowEntry_colname = $tbodyRowEntry_colname.ToLower()
                        $tbodyRowEntry_valign = $tbodyRowEntry_valign.ToLower()
                        $tbodyRowEntry_align = $tbodyRowEntry_align.ToLower()
                        $tbodyRowEntry_morerows = $tbodyRowEntry_morerows.ToLower()
                    }
                    catch{}

                    #if (($childNode.PARA.Length -gt 0 ) -or ($childNode.PARA.'#text'.Length -gt 0) -or ($childNode.PARA.ChildNodes.Count -gt 0))
                    if($childNode.FirstChild.Name -eq "REVST")
                    {
                        <#
                            $revText = $childNode.FirstChild.InnerText
                            $tbody_para_text = "<para>$revText</para>"
                        #>
                    }
                     elseif($childNode.FirstChild.TOOLINFO.TOOLPNR.Length -gt 0)
                    {
                        $tbody_para_text = "<para>" + [string] ($childNode.FirstChild.TOOLINFO.TOOLPNR).ToUpper() + "</para>"
                    }                   
                    
                    elseif($childNode.FirstChild.HWINFO.PNR.Length -gt 0)
                    {
                        $tbody_para_text = "<para>" + [string] ($childNode.FirstChild.HWINFO.PNR).ToUpper() + "</para>"
                    }
                    elseif(($childNode.FirstChild -eq "PARA") -and ($childNode.LastChild -eq "PARA"))
                    {                                                
                        $tbody_para_text = ConvertFrom-PARA -EL2Convert $childNode -ELType $childNode.FirstChild.Name                                         
                    }
                    else
                    {
                        $tbody_para_text = ConvertTo-S1000D4_0_1 -EL2Convert $childNode.FirstChild -ELType $childNode.FirstChild.Name
                    }
                
                    if ($tbodyRowEntry_morerows.Length -gt 0)
                    {
                        $tbodyRowTxt = $tbodyRowTxt + "<entry colname=`"$tbodyRowEntry_colname`" morerows=`"$tbodyRowEntry_morerows`" valign=`"$tbodyRowEntry_valign`" align=`"$tbodyRowEntry_align`">$tbody_para_text</entry>"
                    }else
                    {
                        $tbodyRowTxt = $tbodyRowTxt + "<entry colname=`"$tbodyRowEntry_colname`" valign=`"$tbodyRowEntry_valign`" align=`"$tbodyRowEntry_align`">$tbody_para_text</entry>"
                    }
                    if(($i + $tbodyRowEntry_morerows) -gt $endRowNum)
                    {
                        $endRowNum = $endRowNum + $tbodyRowEntry_morerows
                        $morerows.Value =  [int] $tbodyRowEntry_morerows
                    }
                }
                $tbodyRowTxt = $tbodyRowTxt + "</row>"
                $arrRow += $tbodyRowTxt
                $rowID ++
            }
            

            # iterate through Array.
            foreach( $r in $arrRow)
            {
                $addRowText += $r
            }
            


            $tblTxt = $tblTxt + "$tableTagTxt$titleTagTxt$tableTgroupTagTxt$colSpecTxt$theadTxt$tbodyTxt$addRowText</tbody></tgroup></table>"                                    
        }        
    }
    return $tblTxt
}
Function ConvertFrom-TITLE
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType )

    if($EL2Convert.ParentNode.Name -eq "SPSUBJ")
    {
        return ""
    }else
    {
        $openTagTxt = "<title>"
        $closeTagTxt = "</title>"    
        return $openTagTxt + $EL2Convert."#text" + $closeTagTxt
    }    
}
Function ConvertFrom-UNLIST
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType)
    $openTagTxt = "<para><randomList listItemPrefix=`"`pf03`">"
    $closeTagTxt = "</randomList></para>"
    $childnode_text_out = ""
    foreach ($childNode in $EL2Convert.ChildNodes)
    {
        $childnode_text_out +=   ConvertFrom-UNLITEM -EL2Convert $childNode -ELType $EL2Convert.Name
    }
    return $openTagTxt + $childnode_text_out + $closeTagTxt
}
Function ConvertFrom-UNLITEM
{
    Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType)
    $openTagTxt = "<listItem>"
    $closeTagTxt = "</listItem>"
    $childnode_text_out = ""
    foreach ($childNode in $EL2Convert.ChildNodes)
    { 
        $childnode_text_out += $openTagTxt
        $childnode_text_out += ConvertTo-S1000D4_0_1 -EL2Convert $childNode -ELType $childNode.Name
        $childnode_text_out += $closeTagTxt
    }    
    return $childnode_text_out
}
Function Save-String2ToFolder_BAK
{
    Param([string] $string2Store, [string] $fullPath)
    #Set-Content -Path $fullPath -Value $string2Store 
    $stream = [System.IO.StreamWriter] $fullPath;
    $stream.WriteLine($string2Store);
    $stream.close();
}
#endregion