cls

$outputPath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$reportName = "SE and Consumables per DM - November 2017.csv"
$solventReportName = "Data Modules that require Sealants.csv"
$objs =  @()
$sealants =  @()
$Publist = @("AMM")


foreach ($pub in $Publist)
{
    $inputPath = "F:\KC46 Staging\production\Manuals\$pub\S1000D\SDLLIVE\DMC*.xml"
    $files = gci -Path  $inputPath -Verbose
    foreach ($file in $files)
    {
        $supplyNodes = $null
        $seNodes = $null
        $DMC =  [xml](Get-Content -Path $file.FullName)
        $FN = $file.Name
        #/dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle/techName
        $techName = $DMC.DocumentElement.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
        $infoName = $DMC.DocumentElement.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName
   
        $seNodes = $DMC.SelectNodes("/dmodule/content/procedure/preliminaryRqmts/reqSupportEquips/supportEquipDescrGroup/supportEquipDescr") | Sort-Object -Property Description
        $supplyNodes = $DMC.SelectNodes("/dmodule/content/procedure/preliminaryRqmts/reqSupplies/supplyDescrGroup/supplyDescr") | Sort-Object -Property Description
        
        #Support Equipment
        foreach ($node in $seNodes) 
        {
            $preferred = $null # presets the preffered tag value in the hash table
            if($node.authorityName.Length -gt 0)
            {
                $preferred = 1
            }
            $seToolNum = $node.toolRef.toolNumber
            $seMfgCode = $node.toolRef.manufacturerCodeValue
            $seName = $node.name
            $seReqQty = $node.reqQuantity
            $seShortName = $node.shortName
            $obj = [pscustomobject][ordered]@{Manual=$pub;DMC=$FN;TechName=$techName;InfoName=$infoName;seDescription=$seName;sePN=$seToolNum;sePartName=$seShortName;seQty=$seReqQty;seCage=$seMfgCode;sePreferred=$preferred;SupplyName="";SupplyRqmtNumber="";SupplyReqQty="";seRecord=1;supplyRecord=0;}
            $objs += $obj
            $obj=$null
        }
        <#Consumables
        foreach ($node in $supplyNodes) 
        {
            #Supplies
            $supplyName = $node.name.ToString()
            $supplyName
            $supplyName.ToUpper() -contains "SEALANT"
            $supplyReqRef=$node.supplyRqmtRef.supplyRqmtNumber
            $supplyReqQty=$node.reqQuantity
            $obj = [pscustomobject][ordered]@{Manual=$pub;DMC=$FN;TechName=$techName;InfoName=$infoName;seDescription="";sePN="";sePartName="";seQty="";seCage="";sePreferred="";SupplyName=$supplyName;SupplyRqmtNumber=$supplyReqRef;SupplyReqQty=$supplyReqQty;seRecord=0;supplyRecord=1;}
            $objs += $obj
            if(($supplyName.ToUpper()).Contains("SEALANT"))
            {
                $supplyName
                $sealants += $obj
            }
            $obj=$null
        }#>
    }
}

# Define the sort properties for the report
$prop1 = @{Expression='Manual'; Ascending=$true }
$prop2 = @{Expression='DMC'; Ascending=$true }
$prop3 = @{Expression='seDescription'; Ascending=$true }
$prop4 = @{Expression='sePreferred'; Descending=$true }
$prop5 = @{Expression='sePN'; Ascending=$true }

# Store the report
$objs.GetEnumerator() | Sort-Object -Property $prop1, $prop2,$prop3,$prop4,$prop5 | Export-Csv "$outputPath\$reportName" -NoTypeInformation 
#$sealants.GetEnumerator() | Sort-Object -Property $prop1, $prop2,$prop3,$prop4,$prop5 | Export-Csv "$outputPath\$solventReportName" -NoTypeInformation 