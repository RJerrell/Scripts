CLS
$ErrorActionPreference = "Stop"

<#
Title: KC46 - Set USAF SecondVerification Values
Author: Roger Jerrell
Date Created: 09/11/2017
Purpose: Set the values for the secondVerification tags to the values supplied by the USAF
Description of Operation: Use a white list of values to drive the values
Description of Use: 
    - Read in an Excel file supplied by the CUSTOMER and set the secondVerification element to the value supplied.
    - If the IssueNumber for a DMC is lower than the Verified value in the spreadsheet, the secondVerification element must be removed.
    - A text log will be kept and furnished to the CUSTOMER for any discrpencies.
#>

# *****************************************************************************************************
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force
Add-Type -Path "F:\S1000D_Parser\Parser.dll" -Verbose

$parserDM = new-object -TypeName S1000D.DataModule_401
$parserPM = new-object -TypeName S1000D.PublicationModule_401
$parserCommon = New-Object -TypeName S1000D.CommonFunctions
# *****************************************************************************************************
# Import data from an XLSX spreadsheet
$spreadsheetName = "BOEING Verification - Release 10.2 dated 9-7-2018 - USAF Response.xlsx"
$spreadsheetLocation = "\\nw\newgen-tanker\Spt_TrngSys\SptSys\Tech Pubs\Verification - AFTO 27 Form - Inputs and Process\Release 10.2 - Verification Request - USAF"
$spreadsheetFullPath = "$spreadsheetLocation\$spreadsheetName"
if(Test-Path -Path "F:\$spreadsheetName")
{
    Remove-Item -Path "F:\$spreadsheetName"
}

Copy-Item -Path $spreadsheetFullPath -Destination "F:\$spreadsheetName" -Force

#$USAFList = Import-XLSX -Path "C:\KC46 Staging\Scripts\Report Generators\Outputs\Verification\In\3. BOEING Verifications-UPDATED.xlsx" -Sheet "Verifications_UPDATED"
# $USAFList = Import-XLSX -Path "\\nw\newgen-tanker\Spt_TrngSys\SptSys\Tech Pubs\Verification - AFTO 27 Form - Inputs and Process\Release 10 - Verification Request - USAF\BOEING Verification - Releases 5 thru 10 - 5-24-2018 - All Books - File Used to update CSDB.xlsx" -Sheet "Verifications"

$USAFList = Import-XLSX -Path "F:\$spreadsheetName"  -Sheet "Verifications"
# Processing log location
$processingLog = "C:\KC46 Staging\Scripts\Report Generators\Outputs\Verification\Out\USAF 2ND VERIFICATIONS - ProcessingLog - $startTime.csv"

# Where are the manuals to update??
$AMM_DataPath = "C:\KC46 Staging\Production\FullCSDB\AMM\S1000D"

$log = @()
$log += "Book `t DMC `t  Location `t FullFileName `t BOE_CertificationType `t USAF_VerificationType"

$boe_ttandoo_CTR = 0
$boe_tt_CTR = 0
$boe_OO_Ctr = 0
$usaf_tt_CTR = 0
$usaf_ttandoo_CTR = 0
$usaf_OO_Ctr = 0
$usaf_BlankorInvalid_CTR = 0
$totalRecords = $USAFList.Count
for ($i = 0; $i -lt $USAFList.Count; $i++)
{
    "$i of $totalRecords"
    $usafValue = ""
    $boeValue = ""
    $book = ""
    $dmc = ""
    $boeValue = "tabtop"
    $usafValue = ""

    # 
    #$Boeing_Cert =  Column "BOE_CertificationType"
    #$USAF_vER = Column "USAF_VerificationType"
    
    $dmc = ($USAFList[$i].DMC.ToString()).ToUpper()
    $book = $parserCommon.GetAssociatedBook($dmc)

    $boeValue = $USAFList[$i].BOE_CertificationType     #.ToString()).ToLower()
    $usafValue = $USAFList[$i].USAF_VerificationType     #.ToString()).ToLower()  -- Removed the code compensating for typos in the USAF values supplied

    $NEWPATH = $AMM_DataPath.Replace("AMM",$book) # This is the historical CSDB path -- not the working Manuals path

    <#
    $dms = gci -Path $NEWPATH -Filter "$dmc`*.xml" -Recurse | Sort-Object -Descending | Select-Object -First 2  # Should be set to 2
    
    $fname1 = "" # path to the S1000D file in the historical CSDB
    $fname2 = "" # path to the S1000D file in the production SDLLIVE folder
    HISTORICAL CSDB PROCESSING
    foreach ($dm in $dms)
    {    
        $fname1 = $dm.FullName
        Set-ItemProperty $dm.FullName -name IsReadOnly -value $false
        $parserDM.ParseDM($fname1)  
        if(($usafValue -eq "tabtop" )-or ($usafValue -eq "onobject") -or  ($usafValue -eq "ttandoo") )
        {
            $parserDM.Set_QualityAssurance($boeValue,$usafValue)               
        }
        else
        {
            $parserDM.Set_QualityAssurance($boeValue,"none")
        }
        $log += "$book `t $dmc `t  FullCSDB `t " +  "$fname1 `t $boeValue `t $usafValue"
    }
    #>

    # PRODUCTION MANUALS PROCESSING
    $prodPath = $AMM_DataPath.Replace("FullCSDB", "Manuals").Replace("AMM", $book) + "\SDLLIVE"
    $prodPath = $prodPath.Replace("C:","F:")
    $dmList = gci -Path $prodPath -Filter "$dmc`*.xml" # Only 1 file
    if ($dmList.Count -ne 1 )
    {
        "Too many files in the production environment with the same name"
        exit
    }

    $fname2 = $dmList[0].FullName
    $fName2_Short = $dmList[0].Name
    Set-ItemProperty $dmList[0].FullName -name IsReadOnly -value $false
    $parserDM.ParseDM($fname2)

    if($boeValue -eq "tabtop")
    {
        $boe_tt_CTR ++
    }
    elseif($boeValue -eq "onobject")
    {
        $boe_OO_Ctr ++
    }
    elseif($boeValue -eq "ttandoo")
    {
        $boe_ttandoo_CTR ++
    }


    if(($usafValue -eq "tabtop" )-or ($usafValue -eq "onobject") -or  ($usafValue -eq "ttandoo") )
    {
        $parserDM.Set_QualityAssurance($boeValue,$usafValue)    
        if($usafValue -eq "tabtop" )
        {
            $usaf_tt_CTR ++
        } 
        elseif($usafValue -eq "onobject" )
        {
            $usaf_OO_Ctr ++
        }
        elseif($usafValue -eq "ttandoo" )
        {
            $usaf_ttandoo_CTR ++
        }          
    }
    else
    {
        $parserDM.Set_QualityAssurance($boeValue,"none")
        $usaf_BlankorInvalid_CTR ++        
    }
    $log += "$book `t $dmc `t  PRODUCTION `t " +  "$fName2_Short `t $boeValue `t $usafValue"
    $ctr = ([int] $i + 1).ToString()
    "$ctr of $totalRecords"
}

Out-File -FilePath $processingLog -Force -InputObject $log

"Boeing ttandoo count: `t " + $boe_ttandoo_CTR
"Boeing tabtop count: `t " + $boe_tt_CTR
"Boeing onobject count: `t " + $boe_OO_Ctr

"USAF ttandoo count: `t " + $usaf_ttandoo_CTR
"USAF tabtop count: `t " + $usaf_tt_CTR
"USAF onobject count: `t " + $usaf_OO_Ctr
"USAF blank or invalid: `t " + $usaf_BlankorInvalid_CTR


# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"$PSCommandPath `r`n" +  "Process completed"