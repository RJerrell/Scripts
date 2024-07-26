cls
#region Variables

$commonRoot = "KC46"
$KC46DataRoot = "F:\KC46 Staging"
$source_BaseLocation = "$KC46DataRoot\production\Manuals"
$outputPath = "C:\KC46 Staging\scripts\Report Generators\Outputs"

$reportNameCSV = "KC46 S1000D Quality Assurance Report - FullDataset.csv"
$reportNameDESC = "KC46 S1000D Quality Assurance Report - Description DMs.csv"
$reportNamePROC = "KC46 S1000D Quality Assurance Report - Procedure DMs.csv"

$objs = @()

[string[]] $PubList   = @("KC46","ACS","AMM","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WUC","WDM")

[string[]] $PubList   = @("AMM")

#endregion

# get an XMLTextWriter to create the XML

foreach ($pub in $PubList)
{
    $pathToDMs = "$source_BaseLocation\$pub\s1000d\sdllive\dmc*.xml"
    
    $files = gci -Path $pathToDMs | Sort-Object
    
    foreach ($file in $files)
    {
        $dm = new-object System.Xml.XmlDocument
        
        $FFN = $file.FullName
        
        $FSN = $file.Name.ToUpper().Replace(".XML", "")
        
        $dm.Load($FFN)
        
        # /dmodule/identAndStatusSection/dmAddress/dmIdent/issueInfo
        $issueNumber = $dm.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.issueNumber
        
        # /dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle/techName
        $techName = $dm.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName

        # /dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle/techName
        $infoName = $dm.dmodule.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName

        # /dmodule/content
        $dmType = $dm.dmodule.content.LastChild.Name       
        
        $Total_COC_Entries = ([regex]::Matches($dm.dmodule.content.OuterXml, "authorityName=`"COC`"" )).count
        
        $BDS_COC_Details = ([regex]::Matches($dm.dmodule.content.OuterXml, "id=`"acr-" )).count 
        
        $CAS_COC_DetailsCount = $CAS_COC_Details - $BDS_COC_Details

        $obj = [pscustomobject][ordered]@{Total_COC_Entries="$Total_COC_Entries";CAS_COC_DetailsCount="$CAS_COC_DetailsCount";BDS_COC_Details="$BDS_COC_Details";Manual=$pub;TechName=$techName;InfoName=$infoName;DMC=$FSN;issueNumber=$issueNumber;Type=$dmType;}
        
        $objs += $obj

        $obj = $null

        $dm = $null
    }     
}

# -- Sort properties for the report
$prop1 = @{Expression ='Manual'; Ascending=$true }
$prop2 = @{Expression ='Type'; Ascending=$true }
$prop3 = @{Expression ='DMC'; Ascending=$true }

# Store the full data report
$objs.GetEnumerator() | Sort-Object -Property $prop1,$prop2,$prop3 | Export-Csv "$outputPath\$reportNameCSV" -NoTypeInformation

# Generate breakdown reports
$proc , $desc = $objs.Where({$_.Type -eq "procedure" },'Split') # creates 2 arrays: 1 for the procedure dms and a 2nd non-procedure dms (IPC included)
$desc.GetEnumerator() | Sort-Object -Property $prop1,$prop2,$prop3 | Export-Csv "$outputPath\$reportNameDESC" -NoTypeInformation
$proc.GetEnumerator() | Sort-Object -Property $prop1,$prop2,$prop3 | Export-Csv "$outputPath\$reportNamePROC" -NoTypeInformation

"$outputPath\$reportNameCSV"
"$outputPath\$reportNameDESC"
"$outputPath\$reportNamePROC"