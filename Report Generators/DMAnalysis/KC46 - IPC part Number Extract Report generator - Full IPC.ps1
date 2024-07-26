
<#

Title: 
Author: Roger Jerrell
Date Created: Get-Date 
Purpose: 
Description of Operation: 
Description of Use:

#>
cls
$sd = Get-Date
$ErrorActionPreference = "SilentlyContinue"
$error.Clear()
$a = (Get-Host).PrivateData
$a.WarningBackgroundColor = "Yellow"
$a.WarningForegroundColor = "Black"
$rn = "<br/>"
$a.ErrorBackgroundColor = "red"
$a.ErrorForegroundColor = "White"

$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"

#Get-Module -ListAvailable
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force 
# *****************************************************************************************************
$pathToDatamodules = "F:\KC46 Staging\Production\Manuals\IPB\S1000D\SDLLIVE\DMC*.XML"
$dms = gci -Path $pathToDatamodules
$reportpath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$rptBaseName = "IPC Part Number Listing - Provisioning Report"
$rptSuffix1 = "- Sorted by DMC"

$allParts =  @()
$objsec = $null
$dmXML = New-Object System.Xml.XmlDocument

foreach ($dm in $dms)
{
    $dmXML.Load($dm.FullName)
    # /dmodule/content/illustratedPartsCatalog/catalogSeqNumber
    $CSNS = $dmXML.SelectNodes("/dmodule/content/illustratedPartsCatalog/catalogSeqNumber")

    foreach ($CSN in $CSNS)
    {    
        $isns = $CSN.itemSequenceNumber
        $catalogItemNumber = $CSN.catalogItemNumber
        $indenture = $CSN.indenture
        foreach ($isn in $isns)
        {               
            $position = ""
            $equipment = ""
            $cmm = ""
            $hci = ""
            $hci  = $isn.partCharacteristic
            if($hci -ne $null -and $hci -eq "pc06")
            {
                $hci = "Part with electrostatic discharge sensitivity"
            }
            elseif($hci -ne $null -and $hci -eq "pc01")
            {
                $hci = "A hardness critical item"
            }

            $partNumber = ([string] ($isn.partNumber)).Trim()
            $quantityPerNextHigherAssy = $isn.quantityPerNextHigherAssy
            $sourceMaintRecoverability = $isn.locationRcmdSegment.locationRcmd.sourceMaintRecoverability
            $genericPartDataGroup = $isn.genericPartDataGroup
            $CAGE = $isn.manufacturerCode
            $Description = ([string] ($isn.partIdentSegment.descrForPart)).Trim()
            # /dmodule/content/illustratedPartsCatalog/catalogSeqNumber/itemSequenceNumber/applicabilitySegment/usableOnCodeAssy
            $effectivity = ([string] ($isn.applicabilitySegment.usableOnCodeAssy)).PadLeft("000")
            $origEff = $isn.applicabilitySegment.usableOnCodeAssy
            foreach ($genericPartDataValue in $genericPartDataGroup)
            {
                if(($genericPartDataValue.genericPartData.genericPartDataName).ToString().ToUpper().Contains("POSITION DATA"))
                {
                    $position = $genericPartDataValue.genericPartData.genericPartDataValue.ToString().Trim()
                }
                if(($genericPartDataValue.genericPartData.genericPartDataName).ToString().ToUpper().Contains("ELECTRICAL EQUIP NUMBER"))
                {
                    $equipment = $genericPartDataValue.genericPartData.genericPartDataValue.ToString().Trim()
                }
                if(($genericPartDataValue.genericPartData.genericPartDataName).ToString().ToUpper().Contains("COMPONENT MAINT MANUAL REF"))
                {
                    $cmm = $genericPartDataValue.genericPartData.genericPartDataValue.ToString().Trim()
                }
            }

            $dmParts = $dm.Name.Split("-")

            $CH = $dmParts[3]
            $SE = $dmParts[4]
            $SU = [string] $dmParts[5].SubString(0,2)


            $objsec = [pscustomobject][ordered]@{DMC=$dm.Name;CH=$CH;SE=$SE;SU=$SU;CCIN=$catalogItemNumber;Part_Number=$partNumber;CAGE=$CAGE;DESC=$Description;UNITS_PER_ASSY=$quantityPerNextHigherAssy;EFF=$effectivity;Position=$position;Equipment=$equipment;HCI=$HCI;CMM=$cmm}
            $allParts += $objsec
            $objsec = $null        
          }
     }    
}

Remove-Item -Path "$reportpath\$rptBaseName $rptSuffix1.xlsx" -Force -ErrorAction SilentlyContinue
$prop1 = @{Expression='DMC'; Ascending=$true }
$prop2 = @{Expression='CIN'; Ascending=$true }
$prop3 = @{Expression='Part_Number'; Ascending=$true }
$prop4 = @{Expression='CAGE'; Ascending=$true }
$allParts.GetEnumerator() | Sort-Object -Property $prop1, $prop2, $prop3,$prop4 | Export-Csv "$reportpath\$rptBaseName $rptSuffix1.csv" -NoTypeInformation -Encoding UTF8
$allParts | Export-XLSX -Path "$reportpath\$rptBaseName $rptSuffix1.xlsx" -WorksheetName "IPC" 
"$reportpath\$rptBaseName $rptSuffix1.csv"
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"