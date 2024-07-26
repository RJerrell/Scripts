cls
# The preferred status of parts cannot be determined until the CSDB has been augmented with SERD DB values.
# Without augmentation there will be SE part numbers but no PREFERRED STATUS notation
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$outputPath = "C:\KC46 Staging\Scripts\Report Generators\Outputs\All SE and Consumables" +  (Get-Date -Format 's').Replace(":","-") + ".csv"
$objECTSs =  @()
#[string[]] $PubList   = @($commonRoot, "ACS","AMM","ARD","FIM","IPB","LOAPS","NDT","SIMR","SPCC","SSM","SWPM", "TC","WUC","WDM")
[string[]] $MS   = @("AMM")
[bool] $processSupplyNodes = $true
[bool] $processSENodes = $true
$dmParser = new-object -TypeName S1000D.DataModule_401
#$files = gci -Path "C:\Users\drc9577\Documents\_CandV\DMC_Workshop\DMC*.XML"

foreach ($M in $MS | Sort-Object -Descending)
{
    $inputPath = "F:\KC46 Staging\production\Manuals\$M\S1000D\SDLLIVE\DMC*.xml"
    $files = gci -Path  $inputPath -Verbose
    foreach ($file in $files)
    {
        Measure-Command{        
            $DMC = $dmParser.GetDM($file.FullName)
        }

        $FN = $file.Name
        #/dmodule/identAndStatusSection/dmAddress/dmAddressItems/dmTitle/techName
        $techName = $DMC.DocumentElement.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.techName
        $infoName = $DMC.DocumentElement.identAndStatusSection.dmAddress.dmAddressItems.dmTitle.infoName
        $seNodes = $DMC.SelectNodes("/dmodule/content/procedure/preliminaryRqmts/reqSupportEquips/supportEquipDescrGroup/supportEquipDescr") | Sort-Object -Property Description
        # /dmodule/content/procedure/preliminaryRqmts/reqSupplies/supplyDescrGroup/supplyDescr
        $supplyNodes = $DMC.SelectNodes("/dmodule/content/procedure/preliminaryRqmts/reqSupplies/supplyDescrGroup/supplyDescr") | Sort-Object -Property Description
        if($processSENodes)
        {
            $nodeCount = $seNodes.ChildNodes.Count
            if($nodeCount -gt 0)
            {                
                foreach ($node in $seNodes) 
                {
                    $preferred = $null # presets the preffered tag value in the hash table
                    #$node.authorityName
                    if($node.authorityName.Length -gt 0)
                    {
                        $preferred = 1
                    }
                    #Support Equipment attributes
                    $seToolNum = $node.toolRef.toolNumber
                    $seMfgCode = $node.toolRef.manufacturerCodeValue
                    $seName = $node.name
                    $seReqQty = $node.reqQuantity
                    $seShortName = $node.shortName
                    $objsec = [pscustomobject][ordered]@{Manual=$M;DMC=$FN;TechName=$techName;InfoName=$infoName;seDescription=$seName;sePN=$seToolNum;sePartName=$seShortName;seQty=$seReqQty;seCage=$seMfgCode;sePreferred=$preferred;SupplyName="";SupplyRqmtNumber="";SupplyReqQty="";seRecord=1;supplyRecord=0;}
                    $objECTSs += $objsec
                    $objsec=$null
                }
            }
        }
        if($processSupplyNodes)
        {
            $nodeCount = $supplyNodes.ChildNodes.Count
            if($nodeCount -gt 0)
            {                
                foreach ($node in $supplyNodes) 
                {
                    #Supplies
                    $supplyName = $node.name
                    $supplyReqRef=$node.supplyRqmtRef.supplyRqmtNumber
                    $supplyReqQty=$node.reqQuantity
                    $objsec = [pscustomobject][ordered]@{Manual=$M;DMC=$FN;TechName=$techName;InfoName=$infoName;seDescription="";sePN="";sePartName="";seQty="";seCage="";sePreferred="";SupplyName=$supplyName;SupplyRqmtNumber=$supplyReqRef;SupplyReqQty=$supplyReqQty;seRecord=0;supplyRecord=1;}
                    $objECTSs += $objsec
                    $objsec=$null
                }
            }
        }
    }
}
$prop1 = @{Expression='Manual'; Ascending=$true }
$prop2 = @{Expression='DMC'; Ascending=$true }
$prop3 = @{Expression='Description'; Ascending=$true }
$prop4 = @{Expression='Preferred'; Descending=$true }
$prop5 = @{Expression='PN'; Ascending=$true }

$objECTSs.GetEnumerator() | Sort-Object -Property $prop1, $prop2,$prop3,$prop4,$prop5 | Export-Csv $outputPath -NoTypeInformation
"The report awaits you: " + $outputPath