cls
$sd = Get-Date
$ErrorActionPreference = "Stop"
$error.Clear()

# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
# *****************************************************************************************************
$parserDMC = new-object -TypeName S1000D.DataModule_401
$parserDML = New-Object -TypeName S1000D.DataModuleList
$parserCOMM = New-Object -TypeName S1000D.CommonFunctions

$StartDate=(GET-DATE)
$currYear = (GET-DATE).Year
$currMonth= (GET-DATE).Month
$currDay= (GET-DATE).Day

#region Table Templates
$tableTemp = [xml] @"
            <table colsep="0" frame="none" rowsep="0" tabstyle="CALS">
                <title></title>
				<tgroup cols="10" colsep="0" rowsep="0">
					<colspec colname="Manual" colnum="1" colsep="0" colwidth="5*" rowsep="0" />
					<colspec colname="DMC" colnum="2" colsep="0" colwidth="10*" rowsep="0" />
					<colspec colname="CH" colnum="2" colsep="0" colwidth="5*" rowsep="0" />
					<colspec colname="SE" colnum="3" colsep="0" colwidth="5*" rowsep="0" />
					<colspec colname="SU" colnum="4" colsep="0" colwidth="5*" rowsep="0" />
					<colspec colname="INumber" colnum="1" colsep="0" colwidth="5*" rowsep="0" />
					<colspec colname="IYear" colnum="2" colsep="0" colwidth="10*" rowsep="0" />
					<colspec colname="IMonth" colnum="3" colsep="0" colwidth="5*" rowsep="0" />
					<colspec colname="IDay" colnum="4" colsep="0" colwidth="5*" rowsep="0" />
					<colspec colname="RFU" colnum="4" colsep="0" colwidth="45*" rowsep="0" />
					<thead>
						<row>
							<entry align="left" rowsep="1" valign="bottom">
								<para>Manual</para>
							</entry>
							<entry align="left" valign="bottom">
								<para>DMC</para>
							</entry>
							<entry align="center" valign="bottom">
								<para>Chapter</para>
							</entry>
							<entry align="center" rowsep="1" valign="bottom">
								<para>Section</para>
							</entry>
							<entry align="center" valign="bottom">
								<para>Subject</para>
							</entry>
							<entry align="center">
								<para>Issue Number</para>
							</entry>
							<entry align="center">
								<para>Issue Year</para>
							</entry>
							<entry align="center">
								<para>Issue Month</para>
							</entry>
							<entry align="center">
								<para>Issue Day</para>
							</entry>
							<entry align="left">
								<para>RFU (blank if new)</para>
							</entry>
						</row>
					</thead>
					<tbody><row><entry></entry></row></tbody>
				</tgroup>
			</table>
"@
$tbody = [xml] @"
    <tbody></tbody>                         
"@
$newRow    = [xml] @"
    <row>
        <entry><para/></entry>
        <entry><para/></entry>
        <entry><para/></entry>
        <entry><para/></entry>
        <entry><para/></entry>
        <entry><para/></entry>
        <entry><para/></entry>
        <entry><para/></entry>
        <entry><para/></entry>
        <entry><para/></entry>
    </row>                         
"@

#endregion
#region Variables

$dataModuleCode , $fileName , $pub = ""
$filePref = "DMC"
$dmcs_withhighlight_Tags = @{};
$dmc_to_put_in_hilites = @{};
$oldHash =  @{}
$PathArray = @{}

$Text = "updateHighlight=`"1`""
$basePath = "C:\KC46 Staging\Production\Manuals"
$dmlFolder = "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set"
"$PSScriptRoot\Templates"
$highlightsTemplateFolder = "$PSScriptRoot\Templates"
$highlightsTemplateName = "DMC-1KC46-A-00-00-0000-00A0K-00UA-D.xml"
$highlightsTemplateNamePrefix = "DMC-1KC46-A-00-00-0000-00A"
$highlightsTemplateNameSuffix = "K-00UA-D.xml"
$highlightTemplatePath = "$highlightsTemplateFolder\$highlightsTemplateName"
$highlightOutputFolder = "$dmlFolder\Highlights"
$highlightsDMNames = [hashtable] @{
    "ACS" = "H|Aircraft Cross Servicing Guide (ACS)";
    "AMM" = "A|Aircraft Maintenance Manual (AMM)";
    "ARD" = "E|Aircraft Recovery Dcoument (ARD)";
    "ASIP" = "C|Aircraft Structural Integrity Program (ASIP)";
    "FIM" = "F|Fault Isolation Manual (FIM)";
    "IPB" = "P|Illustrated Parts Book (IPB)";
    "KC46" = "Z|KC46 Common Lists";
    "LOAPS" = "M|List of Applicable Programs (LOAP)";
    "NDT" = "V|Nondestructive Test Manual (NDT)";
	"NDTS" = "N|Nondestructive Test Manual Supplement (NDTS)";
    "SIMR" = "J|Inspection and Maintenance Requirements (-6)";
    "SSM" = "R|System Schematic Manual (SSM";
    "SWPM" = "U|Standard Wiring Practices Manual (SWPM)";
    "TC" = "T|Task Cards (TC)";
    "WDM" = "W|Wiring Diagram Manual (WDM)";
    "WUC" = "D|Work Unit Code Manual (WUC)";
    "ABDR" = "G|Aircraft Battle Damage Repair Manual (ABDR)";
}

#endregion
#region Setup / cleanup of any existing outputs

if( (Test-Path -Path $highlightsTemplateFolder ) -ne $true )
{
    md $highlightsTemplateFolder
}
if( (Test-Path -Path $highlightOutputFolder ) -ne $true )
{
    md $highlightOutputFolder
}
Remove-Item -Path "$basePath\KC46\S1000D\SDLLIVE\*K-00UA-D.xml" -Force -Verbose
Remove-Item -Path "$highlightOutputFolder\*K-00UA-D.xml" -Force

#endregion
#region Process each DML file

$dmlFiles = gci -Path "$dmlFolder\DML*.XML" | Sort-Object -Property Length 
foreach($file in $dmlFiles)
{
    $pub = ""    
    #region Choose the DML name
    if($file.Name -like "DML-1KC46-AAA0A*")
    { 
            $pub = "AMM"
            }
    if($file.Name -like "DML-1KC46-AAA0E*")
    { 
            $pub = "ARD"
            }
    if($file.Name -like "DML-1KC46-AAA0G*")
    { 
            $pub = "ABDR"
            }
    if($file.Name -like "DML-1KC46-AAA0H*")
    { 
            $pub = "ACS"
            }
    elseif($file.Name -like "DML-1KC46-AAA0B*")
    { 
            $pub = "BCLM"
            }
    elseif($file.Name -like "DML-1KC46-AAA0F*")
    { 
            $pub = "FIM"
            }
    elseif($file.Name -like "DML-1KC46-AAA0R*")
    { 
            $pub = "SSM" 
            }
    elseif($file.Name -like "DML-1KC46-AAA0S*")
    { 
            $pub = "SRM"
            }
    elseif($file.Name -like "DML-1KC46-AAA0T*")
    { 
            $pub = "TC"
            }
    elseif($file.Name -like "DML-1KC46-AAA0V*")
    { 
            $pub = "NDT"
            }        
    elseif($file.Name -like "DML-1KC46-AAA0N*")
    { 
            $pub = "NDTS"
            } 
    elseif($file.Name -like "DML-1KC46-AAA0W*")
    { 
            $pub = "WDM"
            }          
    elseif($file.Name -like "DML-1KC46-AAA0Z*")
    { 
            $pub = "KC46"
            }
    elseif($file.Name -like "DML-1KC46-81205*")
    { 
            $pub = "IPB"
            }
    elseif($file.Name -like "DML-1KC46-AAA0D*")
    { 
            $pub = "WUC"
            }
    elseif($file.Name -like "DML-1KC46-AAA0U*")
    { 
            $pub = "SWPM"
            }
    elseif($file.Name -like "DML-1KC46-AAA0T*")
    { 
            $pub = "TC"
            }
    elseif($file.Name -like "DML-1KC46-AAA0M*")
    { 
            $pub = "LOAPS"
            }
    elseif($file.Name -like "DML-1KC46-AAA0J*")
    { 
            $pub = "SIMR"
            }
    elseif($file.Name -like "DML-1KC46-AAA0C*")
    { 
            $pub = "ASIP"
            }
    #endregion 
    $file.FullName
    $parserDML.ParsePM($file.FullName)
    
    $dmEntrys = $parserDML.DMEntries
    foreach( $dmEntry in $dmEntrys )
    {            
        # .SelectNodes("//[@dmEntryType != 'd']")
        if($dmEntry.dmEntryType -ne "d")
        {
            $dataModuleCode = Get-FilenameFromDMRef -dmRef $dmEntry.dmRef -filePref "DMC"

            $systemCode = $dmEntry.dmRef.dmRefIdent.dmCode.systemCode
            $subSystemCode = $dmEntry.dmRef.dmRefIdent.dmCode.subSystemCode
            $subSubSystemCode = $dmEntry.dmRef.dmRefIdent.dmCode.subSubSystemCode
            $assyCode = $dmEntry.dmRef.dmRefIdent.dmCode.assyCode
            $ch = $systemCode
            $se = $subSystemCode+$subSubSystemCode
            $su = $assyCode
            $fileName =  "$basePath\$pub\S1000D\SDLLIVE\$dataModuleCode*.xml"

            $fs = gci -path "$fileName`*" |Sort-Object -Descending |Select-Object -First 1
        
            if($fs.Length -eq 0)
            {
                "No file found named or beginning with:" + $fileName
                #Exit
            }
            else
            {
                # Process any file in the DML
                $parserDMC.ParseDM($fs[0].FullName)
                $reasonForUpdate = ""

                #/dmodule/identAndStatusSection/dmAddress/dmIdent/issueInfo
                $issueNum   = [string] $parserDMC.IssueInfo.issueNumber
                $issueYear  = [string] $parserDMC.Issue_year.'#text'
                $issueMonth = [string] $parserDMC.Issue_month.'#text'
                $issueDay   = [string] $parserDMC.Issue_day.'#text'
                $techName   = [string] $parserDMC.TechName
                $infoName   = [string] $parserDMC.InfoName
                $rfuData    = $parserDMC.ReasonForUpdate
                #$rfuData
                $fileName
                $reasonForUpdate = ""
                if($issueNum -ne "001")
                {
                    foreach ($rfuPara in $rfuData)
                    {
                        $rfuData.simplePara
                        $reasonForUpdate += $rfuData.simplePara + "|"
                    }
                }
                $oldKeyValue = $oldHash.Get_Item($dataModuleCode)
                $oldIssueNum = $null
                $dmcs_withhighlight_Tags[$dataModuleCode] ="$techName|$pub|$ch|$se|$su|$issueNum|$issueYear|$issueMonth|$issueDay|$reasonForUpdate"                
                if($oldKeyValue -ne $null)
                {
                    $dmc = $oldHash.Get_Item($dataModuleCode)
                    [string[]] $values = $dmc.Split("|")
                    $oldIssueNum = $values[4]
                    if($oldIssueNum -ne $issueNum)
                    {
                        $dmc_to_put_in_hilites.Add($dataModuleCode, "$techName|$pub|$ch|$se|$su|$issueNum|$issueYear|$issueMonth|$issueDay|$reasonForUpdate")
                        $dmcs_withhighlight_Tags[$dataModuleCode] = "$techName|$pub|$ch|$se|$su|$issueNum|$issueYear|$issueMonth|$issueDay|$reasonForUpdate"
                    }
                }
                else
                {                    
                    $dmc_to_put_in_hilites.Add($dataModuleCode, "$techName|$pub|$ch|$se|$su|$issueNum|$issueYear|$issueMonth|$issueDay|$reasonForUpdate") 
                } 
            }           
        }            
    }
}

#endregion

#region Process the data collected from the above processing

[string[]] $pubList = @()
[string[]]$pubValues = $dmcs_withhighlight_Tags.GetEnumerator() |  %{$_.Value.Split("|")[1]}
foreach ($pubValue in $pubValues)
{
    if($pubList -notcontains $pubValue)
    {
        $pubList += $pubValue
    }
}

# Produce the applicable highlights data modules
if($dmcs_withhighlight_Tags.Count -gt 0)
{
    $ntArray = @()
    $allTables = $null
    foreach ($pub in ($pubList | Sort-Object) )
    { 
        $tableTitle = "IETM Highlights - " + $highlightsDMNames.get_item($pub).Split("|")[1]
        $pubNumber = $highlightsDMNames.get_item($pub).Split("|")[0]
        $toTemplateTemp = [xml] (Get-Content -Path $highlightTemplatePath) 
        $toTemplate = $toTemplateTemp.Clone()
        $toTemplate.dmodule.identAndStatusSection.dmAddress.dmAddressItems.issueDate.day   = "22"
        $toTemplate.dmodule.identAndStatusSection.dmAddress.dmAddressItems.issueDate.month = "02"
        $toTemplate.dmodule.identAndStatusSection.dmAddress.dmAddressItems.issueDate.year  = "2018"
        $toTemplate.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode.disassyCodeVariant = "A" + $pubNumber + "K"
        $toTemplate.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName = $tableTitle

        $highlightsDMName = $highlightsTemplateNamePrefix + $pubNumber + $highlightsTemplateNameSuffix
        $nt = $tableTemp.Clone()
        $tb = $tbody.Clone()
        $nt.table.title = [string] $tableTitle
        $list = $dmcs_withhighlight_Tags.GetEnumerator() | Sort-Object -Property Name | ?{[string] $_.Value.Split("|")[1] -eq $pub}
        foreach($key in $list)
        {        
            $nr = $newRow.Clone()
            $rowEntryText = ""
            #Get the value of the item to add to the highlights data module
            [string[]] $dmc = $key.Name
            [string[]] $values = $key.Value.Split("|")
            $nr.row.entry[0].para = $values[1]     #<--- The name of the manaul
            $dmcElement = Create-DMCCodeElementFromDMCFileName -dmCode $dmc[0]
            $dmcElement
            # create a dmref
            $tname = ""
            try
            {
                $tname = [string]$values[0]
                $tname = $tname.Replace("&", " and ")
                $nr.row.entry[1].InnerXml = "<para><dmRef><dmRefIdent>" + $dmcElement.OuterXml + "</dmRefIdent><dmRefAddressItems><dmTitle><techName>$tname</techName></dmTitle></dmRefAddressItems></dmRef></para>"     # <--- The DMC Code

                $nr.row.entry[2].para = $values[2]
                $nr.row.entry[3].para = $values[3]
                $nr.row.entry[4].para = $values[4]
                $nr.row.entry[5].para = $values[5]
                $nr.row.entry[6].para = $values[6]
                $nr.row.entry[7].para = $values[7]
                # RFU has to broekn up into multiple para tags
                $nr.row.entry[8].para = $values[8]
            }
            catch
            {
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                $FailedItem
                $ErrorMessage
                "TechName:`t" + $tname
            }
            if($values.Length -gt 9)
            {
                $rowEntryText = ""
                for ($i = 9; $i -lt $values.Length - 1; $i++)
                { 
                    $val = ""
                    $val = [string] $values[$i]
                    if($val.Length -gt 0)
                    {
                        $rowEntryText += "<para>" + $val + "</para>"
                    }
                }
                try
                {                    
                    if($rowEntryText.Length -gt 0)
                    {
                        $nr.row.entry[9].InnerXml = $rowEntryText
                    }
                }
                catch
                {
                    "Catch: " + $Error                   
                }
                finally
                {
                    
                }
            }
            else
            {
                if($values[9].Length -gt 0)
                {
                    $nr.row.entry[9].para = $values[9].ToString()
                }                        
            }            
            $newNode = $tb.ImportNode($nr.ChildNodes[0], $true)
            $tb.DocumentElement.AppendChild($newNode)
        }        
        $nt.table.tgroup.tbody.InnerXml = $tb.tbody.InnerXml
        #$toTemplate.dmodule.content.description.levelledPara.InnerXml
        $toTemplate.dmodule.content.description.levelledPara.InnerXml = $nt.table.OuterXml            
        $toTemplate.Save("$highlightOutputFolder\$highlightsDMName")
        $allTables = $null
        $nt = $null
    }
}

#endregion

$EndDate=(GET-DATE)
"End of processing.  The elapsed time is: {0:G}" -f (NEW-TIMESPAN –Start $StartDate –End $EndDate)
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"