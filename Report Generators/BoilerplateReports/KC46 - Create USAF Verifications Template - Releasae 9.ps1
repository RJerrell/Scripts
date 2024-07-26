<#
Title: 
Author: Roger Jerrell
Date Created: 11/1/2017
Purpose: Create an Excel spreadsheet for distribution internally to Beoing Technical Publications leaders
Description of Operation: Spreadsheet used to track the IssueNumber history for each data module in the AMM.
Description of Use: 

#>
cls
$sd = Get-Date
$ErrorActionPreference = "Stop"
$error.Clear()
$a = (Get-Host).PrivateData
$a.WarningBackgroundColor = "Yellow"
$a.WarningForegroundColor = "Black"
$rn = "<br/>"
$a.ErrorBackgroundColor = "red"
$a.ErrorForegroundColor = "White"
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"

Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force 
# *****************************************************************************************************
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$parser = new-object -TypeName S1000D.DataModule_401


$rel5Path = "R:\2017-01-20-14-18-01 - Non CDRL January 2017 - Release 5\CSDB\Manuals\AMM\S1000D\SDLLIVE"
$rel6Path = "R:\2017-06-06-07-18-23 - Non CDRL June 2017 - Release 6\CSDB\DVD\AMM\S1000D\SDLLIVE"
$rel7Path = "R:\2017-09-18-14-39-31 - Non CDRL Sept 2017 - Release 7\CSDB\DVD\AMM\S1000D\SDLLIVE"
$rel8Path = "R:\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\AMM\S1000D\SDLLIVE"
$rel9Path = "D:\Shared\IDE cd sets\Releases\Feb 2018 - Release 9\CSDB\DVD\AMM\S1000D\SDLLIVE"
$pathToTheGOOOOPile = "\\nw.nos.boeing.com\newgen-tanker\Spt_TrngSys\SptSys\Tech Pubs\Validation\C&V_workbook\DMCertVerWorksheet.xlsx"

if(Test-Path -Path "$env:TMP\DMCertVerWorksheet.xlsx")
{
    Remove-Item -Path "$env:TMP\DMCertVerWorksheet.xlsx" -Force
}

Copy-Item -Path $pathToTheGOOOOPile -Destination $env:TMP -Force
$pathToTheGOOOOPile = "$env:TMP\DMCertVerWorksheet.xlsx"
$GOOOWorkSheetName = "StatusWorksheet"
$verCompListArray = @()
$badDMCArray = @()
$verListArray = @()
$WSRDTasksCERTdArray = @()
$verCompListArray = @()
$badDMCArray = @()
$verreecords = @()
$WSRDRecords = @()
$badDMCrecords = @()
$verCompListTotals = 0
$badDMCtotals = 0
$verListTotals = 0
$WSRDTasksTotals = 0

$wsRows = Import-XLSX -Path $pathToTheGOOOOPile -Sheet $GOOOWorkSheetName

$zz = 1

foreach ($wsRow in $wsRows)
{  
    $pbTask = $wsRow.'PB/Task'
    if($pbTask -ne $null -and $pbTask.Length -gt 0 -and $pbTask.ToString().ToUpper().Contains("TASK"))
    {
        $pbTask = ($wsRow.'PB/Task').ToString().ToUpper().Replace("TASK", "").Replace("PAGEBLOCK", "").Trim()
    }
    else
    {
        $pbTask = ""       
    }
    
    $verList = [string] $wsRow.'Ver List'
    $verListDMC = $wsRow.DMC
    $certComp = $wsRow.'Cert Comp'
    $verComp = $wsRow.'Ver Comp'

    #$verListDMC
    if($verListDMC -eq $null -or $verListDMC.ToString().ToUpper().Contains("-XX") -or $verListDMC -eq "")
    {
        $badDMCrecords += New-Object -TypeName PSObject -Property @{
            Type       = "BadDMC"
            PBTask     = $pbTask;                                      
            certComp   = $certComp;
            verListDMC = $verListDMC;
            verList    = $verList;
            verComp    = $verComp;
        } | Select Type,PBTask,certComp,verListDMC,verList,verComp
    }

    if($verlist -ne $null -and $verList.ToString().ToUpper().Contains("WSRD"))
    {
        $WSRDRecords += New-Object -TypeName PSObject -Property @{
            Type       = "WSRDTask"
            PBTask     = $pbTask;                                      
            certComp   = $certComp;
            verListDMC = $verListDMC;
            verList    = $verList;
            verComp    = $verComp;
        } | Select Type,PBTask,certComp,verListDMC,verList,verComp
    }
    if($verlist -ne $null -and $verList.ToString().ToUpper().Contains("VER LIST"))
    {
        $verreecords += New-Object -TypeName PSObject -Property @{
            Type       = "VerListTask"
            PBTask     = $pbTask;                                      
            certComp   = $certComp;
            verListDMC = $verListDMC;
            verList    = $verList;
            verComp    = $verComp;
        } | Select Type,PBTask,certComp,verListDMC,verList,verComp        
    }

    if($certComp -ne $null)
    {
        $certCompType = $certComp.GetType()
               
        if($certCompType.Name -eq "DateTime" -and $pbTask.Length -gt 0)
        {    
            $shortDate = ""
            $verList -eq $null
            if($verList -ne $null)
            {
                $verList.GetType()
            }
    
            # WSRD Tasks from the GOOOO Pile
            if($verList -ne $null -and $verList.ToString().ToUpper().Contains("WSRD"))
            {
                $WSRDTasksCERTdArray += New-Object -TypeName PSObject -Property @{
                    PBTask = $pbTask;                                 
                    certComp = $certComp.Date.ToShortDateString();
                    verListDMC = $verListDMC;
                    verList = $verList;
                    verComp = $verComp;
                } | Select PBTask,certComp,verListDMC,verList,verComp
            } #Ver List Tasks in the GOOO Pile
            elseif($verList -ne $null -and $verList.ToString().ToUpper().Contains("VER LIST"))
            {
                $verListArray += New-Object -TypeName PSObject -Property @{
                    PBTask = $pbTask;                                  
                    certComp = $certComp;
                    verListDMC = $verListDMC;
                    verList = $verList;
                    verComp = $verComp;
                } | Select PBTask,certComp,verListDMC,verList,verComp
            }
            # Bad DMC codes in the GOOOO Pile
            if($verListDMC -ne $null -and $verListDMC.ToString().ToUpper().Contains("-XX"))
            {
                $badDMCArray += New-Object -TypeName PSObject -Property @{
                        PBTask = $pbTask;                                      
                        certComp = $certComp.Date.ToShortDateString();
                        verListDMC = $verListDMC;
                        verList = $verList;
                        verComp = $verComp;
                } | Select PBTask,certComp,verListDMC,verList,verComp
            }        
        
            # Actually good records that we believe represent a completed CERT
            $shortDate = $certComp.Date.ToShortDateString()

            if((! $shortDate.Contains("1899")) -and $pbTask.Length -gt 0)
            {
                $verCompListArray += New-Object -TypeName PSObject -Property @{
                        PBTask = $pbTask;                                      
                        certComp = $certComp.Date.ToShortDateString();
                        verListDMC = $verListDMC;
                        verList = $verList;
                        verComp = $verComp;
                } | Select PBTask,certComp,verListDMC,verList,verComp
                
            }
            else
            {
                $shortDate
            }
            #$verCompListArray.Count
            #$zz            
            $zz ++
        }
    }
}

$exportFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs\Verification\Out"
# $reportName = "BOEING Verification - Releases 5 thru 7 - Revised.xlsx"
$rptTimeStamp = $sd.Date.Year.ToString() + $sd.Month.ToString()+ $sd.Day.ToString() + " " +  $sd.Hour.ToString() + $sd.Minute.ToString() + $sd.Millisecond.ToString()
$reportName = "BOEING Verification - Releases 5 thru 9 - " + $rptTimeStamp + ".xlsx"
$reportName
$masterList = @() # Carries all the values we need for this report

$pmXml = New-Object System.Xml.XmlDocument
$ttandooCTR = 0
# Path to the most current release of the CSDB
$pms = gci -Path "$rel9Path\PMC*.XML" -File
$pmXml.Load($pms[0].FullName)
$dmRefs = $pmXml.SelectNodes("/pm/content//dmRef")
$pmEntries = $pmXml.SelectNodes("/pm/content//pmEntry")
$ctr = [int] 0
$ttandooCTR = 0

foreach ($dmRef in $dmRefs)
{
    # Release 9
    #$dmXml = New-Object System.Xml.XmlDocument
    #$pbTask = $dmRef.ParentNode
    #/dmodule/identAndStatusSection/dmAddress/dmIdent/issueInfo
    $fileName = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
    $fileOnDisk = gci -Path "$rel9Path\$fileName`*.xml"| Sort-Object -Descending | Select-Object -First 1
    $parser.ParseDM($fileOnDisk[0].FullName)
    $dmXml = $parser.Dmodule
    #$dmXml.Load($fileOnDisk[0].FullName)
    $IssueNum = $parser.IssueInfo.issueNumber
    $infoName =  $parser.InfoName
    $techName =  $parser.TechName



    # Release 8
    $rel8Issue = ""
    $fileOnDisk8 = gci -Path "$rel8Path\$fileName`*.xml"| Sort-Object -Descending | Select-Object -First 1
    if($fileOnDisk8.Count -gt 0)
    {
        $parser.ParseDM($fileOnDisk8[0].FullName)
        $rel8Issue = $parser.IssueInfo.issueNumber
    }

    # ********************************************** IMPORTANT ******************************************
    $previousIssue = $rel8Issue  # <---- Cahnge this value
    
    $issuedUp = $false
    if([int]$IssueNum -gt [int]$previousIssue)
    {
        $issuedUp = $true
    }
    # ********************************************** change to match the previous release ******************************************

    # Release 7
    $rel7Issue = ""
    $fileOnDisk7 = gci -Path "$rel7Path\$fileName`*.xml"| Sort-Object -Descending | Select-Object -First 1
    if($fileOnDisk7.Count -gt 0)
    {
        $parser.ParseDM($fileOnDisk7[0].FullName)
        $rel7Issue = $parser.IssueInfo.issueNumber
    }

    # Release 6
    $rel6Issue = ""
    $fileOnDisk6 = gci -Path "$rel6Path\$fileName`*.xml"| Sort-Object -Descending | Select-Object -First 1
    if($fileOnDisk6.Count -gt 0)
    {
        $parser.ParseDM($fileOnDisk6[0].FullName)
        $rel6Issue = $parser.IssueInfo.issueNumber
    }
    
    # Release 5
    $rel5Issue = ""
    $fileOnDisk5 = gci -Path "$rel5Path\$fileName`*.xml"| Sort-Object -Descending | Select-Object -First 1
    if($fileOnDisk5.Count -gt 0)
    {
        $parser.ParseDM($fileOnDisk5[0].FullName)
        $rel5Issue = $parser.IssueInfo.issueNumber
    }

    $type = $parser.DmType
    
    $pbTask = [string] $dmRef.href
    if($pbTask.length -eq 0)
    {
        $pbTask = "00-00-00"
    }
    
    $verificationComplete = ""
    $VerListValue = ""
    $boeCertType = "tabtop"


	#$pbTask.ToUpper().Replace("PAGEBLOCK","").Replace("TASK","").Trim()

    foreach ($verCompListItem in $WSRDTasksCERTdArray)
    {
        if($verCompListItem.certComp.Length -gt 0 -and $verCompListItem.PBTask -eq $pbTask.ToUpper().Replace("PAGEBLOCK","").Replace("TASK","").Trim())
        {

            $boeCertType = "ttandoo"
            $VerListValue = "WSRD"
            $ttandooCTR ++
            $verificationComplete = "ttandoo"
            break
        }
    }



    # If the $verificationComplete value has already been determined, don't process it again
    if($verificationComplete.Length -eq 0)
    {
        foreach ($verCompListItem in $verCompListArray)
        {
            if($verCompListItem.certComp.Length -gt 0 -and $verCompListItem.PBTask -eq $pbTask.ToUpper().Replace("PAGEBLOCK","").Replace("TASK","").Trim())
            {
                $boeCertType = "ttandoo"
                $VerListValue = "Ver List"
                $ttandooCTR ++
                $verificationComplete = "ttandoo"
                break
            }
        }
    }

    if($VerListValue -ne "")
    {
        Write-Debug $VerListValue
    }

    [string] $rfuEntry = ""
    if($issuedUp)
    {
        $rfuList = $dmXml.dmodule.identAndStatusSection.dmStatus.reasonForUpdate
        
        $c = 0
        foreach ($rfu in  $rfuList.simplePara)
        {                        
            if($rfuList.ChildNodes.Count -eq 1)
            {
                $rfuEntry = [string] $rfu
            }
            elseif($rfuList.ChildNodes.Count -gt 1)
            {             
                $rfuEntry += [string] $rfu + " | "                       
            }
            $c ++
        }
    }

    $masterList += New-Object -TypeName PSObject -Property @{
                VerListValue = $VerListValue;
                Type = $type;
                PBTask = $pbTask;
                DMC = $fileName;
                TechName=$techName;
                InfoName = $infoName;
                Rel_5 = $rel5Issue;
                Rel_6 = $rel6Issue;
                Rel_7 = $rel7Issue;
                Rel_8 = $rel8Issue;
                Rel_9 = $IssueNum;
                BOE_CertificationType = $boeCertType;
                USAF_VerificationType = "";
                RFU_List = $rfuEntry
                
    } | Select VerListValue,Type,PBTask,DMC,TechName,InfoName,Rel_5,Rel_6,Rel_7,Rel_8,Rel_9,BOE_CertificationType,USAF_VerificationType,RFU_List
    $ctr ++
    "Processing $ctr"
}

# Export it

cls

"Executive summary Report :`t" + $sd

"Out of " + $wsRows.Count + " total rows in the DMCertVerWorksheet.xlsx , " + ($masterList.GetEnumerator() | ?{$_.BOE_CertificationType -eq "ttandoo"}).Count + " had recognizable PB/Task values and had a discernable CERT Date and should be recognized as CERT'd"
"Out of " + $wsRows.Count + " total rows in the DMCertVerWorksheet.xlsx , " + $WSRDRecords.Count + " are designated as 'WSRD' tasks.  Of those, " + $WSRDTasksCERTdArray.Count + " had a discernable CERT Date"
"Out of " + $wsRows.Count + " total rows in the DMCertVerWorksheet.xlsx , " + $verreecords.Count + " are designated as 'VER List' tasks.  Of those, " + $verListArray.Count + " had a discernable CERT Date"
"Out of " + $wsRows.Count + " total rows in the DMCertVerWorksheet.xlsx , " + $badDMCrecords.COUNT + " have bad DMC codes."

Remove-Item -Path "$exportFolder\$reportName" -Force -ErrorAction SilentlyContinue
C:\KC46 Staging\Scripts\Report Generators\Outputs\Verification\Out | Export-XLSX -Path "$exportFolder\$reportName" -Header  VerListValue,Type,PBTask,DMC,TechName,InfoName,Rel_5,Rel_6,Rel_7,Rel_8,Rel_9,BOE_CertificationType,USAF_VerificationType,RFU_List -WorksheetName "Verifications"
$badDMCrecords | Export-XLSX -Path "$exportFolder\$reportName" -Header Type,PBTask,certComp,verListDMC,verList,verComp -WorksheetName "Bad DMCs" 
$verreecords | Export-XLSX -Path "$exportFolder\$reportName" -Header Type,PBTask,certComp,verListDMC,verList,verComp -WorksheetName "Ver List Tasks" 
$WSRDRecords | Export-XLSX -Path "$exportFolder\$reportName" -Header  Type,PBTask,certComp,verListDMC,verList,verComp -WorksheetName "WSRD Tasks" 

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
$ed
"Report now available at this location:`r`n$exportFolder\$reportName"