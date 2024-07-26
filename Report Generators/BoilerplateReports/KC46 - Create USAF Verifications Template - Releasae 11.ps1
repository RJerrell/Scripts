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

$parserDM = new-object -TypeName S1000D.DataModule_401
$parserPM = new-object -TypeName S1000D.PublicationModule_401
$parserCommon = New-Object -TypeName S1000D.CommonFunctions

$rel5Path = "D:\Shared\IDE cd sets\Releases\2017-01-20-14-18-01 - Non CDRL January 2017 - Release 5\CSDB\Manuals\AMM\S1000D\SDLLIVE"
$rel6Path = "D:\Shared\IDE cd sets\Releases\2017-06-06-07-18-23 - Non CDRL June 2017 - Release 6\CSDB\DVD\AMM\S1000D\SDLLIVE"
$rel7Path = "D:\Shared\IDE cd sets\Releases\2017-09-18-14-39-31 - Non CDRL Sept 2017 - Release 7\CSDB\DVD\AMM\S1000D\SDLLIVE"
$rel8Path = "D:\Shared\IDE cd sets\Releases\2017-10-30-11-47-35 - Non CDRL Nov 2017 - Release 8\CSDB\DVD\AMM\S1000D\SDLLIVE"
$rel9Path = "D:\Shared\IDE cd sets\Releases\Feb 2018 - Release 9\CSDB\DVD\AMM\S1000D\SDLLIVE"
$rel10Path = "D:\Shared\IDE cd sets\Releases\May 2018 - Release 10\CSDB\DVD\AMM\S1000D\SDLLIVE"
$rel11Path = "D:\Shared\IDE cd sets\Releases\Sep 2018 - Release 11\CSDB\DVD\AMM\S1000D\SDLLIVE"

$pathToTheGOOOOPile = "\\nw.nos.boeing.com\newgen-tanker\Spt_TrngSys\SptSys\Tech Pubs\Validation\C&V_workbook\DMCertVerWorksheet.xlsx"
$exportFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs\Verification\Out"

$rptTimeStamp = $sd.ToShortDateString().Replace("/","-")
$reportName = "BOEING Verification - Releases 5 thru 11 - " + $rptTimeStamp + ".xlsx"
$reportName
$masterList = @() # Carries all the values we need for this report
$ammDMCList = New-Object System.Collections.Generic.List[String]

[string[]] $PubList   = @("AMM","KC46","ABDR","ACS","ARD","ASIP","FIM","IPB","LOAPS","NDT","SIMR","SPCC","TC","WUC","SSM","SWPM", "WDM")
#[string[]] $PubList   = @("SWPM")

foreach ($Pub in $PubList)
{
    if($pub -eq "AMM")
    {       
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
    }


    $ttandooCTR = 0

    # Path to the most current release of the CSDB
    $rel11NewPath = $rel11Path.Replace("AMM",$Pub)
    $pms = gci -Path "$rel11NewPath\PMC*.XML" -File
    $parserPM.ParsePM($pms[0].FullName)
    $dmRefs = $parserPM.DmRefs

    $pmEntries = $parserPM.PmEntries
    $ctr = [int] 0
    $ttandooCTR = 0
    $new2 = @()
    foreach ($dmRef in $dmRefs)
    {
        # Release 10
        $fileName = Get-FilenameFromDMRef -dmRef $dmRef -filePref "DMC"
        
        # Extract the correct manual name from the $filename
        
        $book = $parserCommon.GetAssociatedBook($fileName)
        if($book -ne $Pub)
        {
            "stop"
        }
        $dmPath = $rel11Path.Replace("AMM",$book)
        $fileOnDisk = gci -Path "$dmPath\$fileName`*.xml"| Sort-Object -Descending | Select-Object -First 1
        
        if($fileOnDisk.Count -gt 0)
        {
        
            $parserDM.ParseDM($fileOnDisk[0].FullName)
            $type = $parserDM.DmType
            #$type
            $dmXml = $parserDM.Dmodule
            $IssueNum = $parserDM.IssueInfo.issueNumber
            $infoName =  $parserDM.InfoName
            $techName =  $parserDM.TechName

            if(($IssueNum -eq "001") -and ($parserDM.IdentAndStatusSection.dmStatus.issueType -eq "new") )
            {
                $y = $parserDM.Issue_year
                $m = $parserDM.Issue_month
                $d = $parserDM.Issue_day
                #$y.ToString() + $m.ToString() + $d.ToString()
                if (($y.ToString() +$m.ToString() + $d.ToString()) -eq "20180422")
                {
                    $new2 += $fileOnDisk[0].Name
                }
            }
        }

        # Release 10
        $rel10Issue = ""
        $rel10NewPath = $rel10Path.Replace("AMM",$Pub)
        $fileOnDisk9 = gci -Path "$rel10NewPath\$fileName`*.xml"| Sort-Object -Descending | Select-Object -First 1
        if($fileOnDisk9.Count -gt 0)
        {
            $parserDM.ParseDM($fileOnDisk9[0].FullName)
            $rel10Issue = $parserDM.IssueInfo.issueNumber
        }

        # ********************************************** IMPORTANT ******************************************
        $previousIssue = $rel10Issue  # <---- Change this value
    
        $issuedUp = $false
        if($previousIssue.Length -gt 0)
        {
            if([int]$IssueNum -gt [int]$previousIssue)
            {
                $issuedUp = $true
            }
        }
        # ********************************************** # End of Release 10 ******************************************


        # Release 9
        $rel9Issue = ""
        $rel9NewPath = $rel9Path.Replace("AMM",$Pub)
        $fileOnDisk9 = gci -Path "$rel9NewPath\$fileName`*.xml"| Sort-Object -Descending | Select-Object -First 1
        if($fileOnDisk9.Count -gt 0)
        {
            $parserDM.ParseDM($fileOnDisk9[0].FullName)
            $rel9Issue = $parserDM.IssueInfo.issueNumber
        }

        
        $previousIssue = $rel9Issue  # <---- Change this value
    
        $issuedUp = $false
        if($previousIssue.Length -gt 0)
        {
            if([int]$IssueNum -gt [int]$previousIssue)
            {
                $issuedUp = $true
            }
        }

    
        # Release 8
        $rel8Issue = ""
        $rel8NewPath = $rel8Path.Replace("AMM",$Pub)
        $fileOnDisk8 = gci -Path "$rel8NewPath\$fileName`*.xml" -ErrorAction SilentlyContinue| Sort-Object -Descending | Select-Object -First 1
        if($fileOnDisk8.Count -gt 0)
        {            
            try
            {
                   $parserDM.ParseDM($fileOnDisk8[0].FullName)
                   $rel8Issue = $parserDM.IssueInfo.issueNumber
            }
            catch
            {}
        }

        # Release 7
        $rel7Issue = ""
        $rel7NewPath = $rel7Path.Replace("AMM",$Pub)
        $fileOnDisk7 = gci -Path "$rel7NewPath\$fileName`*.xml" -ErrorAction SilentlyContinue| Sort-Object -Descending | Select-Object -First 1
        if($fileOnDisk7.Count -gt 0)
        {            
            try
            {
                $parserDM.ParseDM($fileOnDisk7[0].FullName)
                $rel7Issue = $parserDM.IssueInfo.issueNumber
            }
            catch
            {}          
        }

        # Release 6
        $rel6Issue = ""
        $rel6NewPath = $rel6Path.Replace("AMM",$Pub)
        $fileOnDisk6 = gci -Path "$rel6NewPath\$fileName`*.xml" -ErrorAction SilentlyContinue| Sort-Object -Descending | Select-Object -First 1
        if($fileOnDisk6.Count -gt 0)
        {            
            try
            {
                $parserDM.ParseDM($fileOnDisk6[0].FullName)
                $rel6Issue = $parserDM.IssueInfo.issueNumber
            }
            catch
            {}
        }
    
        # Release 5
        $rel5Issue = ""
        $rel5NewPath = $rel5Path.Replace("AMM",$Pub)
        $fileOnDisk5 = gci -Path "$rel5NewPath\$fileName`*.xml" -ErrorAction SilentlyContinue| Sort-Object -Descending | Select-Object -First 1

        if($fileOnDisk5.Count -gt 0)
        {
            try
            {
                $parserDM.ParseDM($fileOnDisk5[0].FullName)
                $rel5Issue = $parserDM.IssueInfo.issueNumber 
            }
            catch
            {}          
        }

        

        $pbTask = [string] $dmRef.href
        if($pbTask.length -eq 0)
        {
            If($Pub -eq "AMM")
            {
                $pbTask = "00-00-00"
            }
        }
    
        $verificationComplete = ""
        $VerListValue = ""
        $boeCertType = "tabtop"

        if($pub -eq "AMM")
        {
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
        if($issuedUp -and $rfuEntry.Length -eq 0)
        {
            $IssueNum += "*"
        }

        if($Pub.ToUpper() -ne "AMM")
        {
            foreach ($item in $masterList)
            {
                if($item.DMC -eq $fileName)
                {                    
                   $boeCertType =  $item.BOE_CertificationType
                   $VerListValue = $item.VerListValue
                   break
                }
            }        
        }
        
        $masterList += New-Object -TypeName PSObject -Property @{
            Manual = $pub;
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
            Rel_9 = $rel9Issue;
            Rel_10 = $IssueNum;
            BOE_CertificationType = $boeCertType;
            USAF_VerificationType = "";
            RFU_List = $rfuEntry
                
        } | Select Manual,VerListValue,Type,PBTask,DMC,TechName,InfoName,Rel_5,Rel_6,Rel_7,Rel_8,Rel_9,Rel_10,BOE_CertificationType,USAF_VerificationType,RFU_List
        
        $ctr ++
        "Processing $ctr" + " in the $Pub manual"
    }
}
# Export it

cls

"Executive summary Report :`t" + $sd

"Out of " + $wsRows.Count + " total rows in the DMCertVerWorksheet.xlsx , " + ($masterList.GetEnumerator() | ?{$_.BOE_CertificationType -eq "ttandoo"}).Count + " had recognizable PB/Task values and had a discernable CERT Date and should be recognized as CERT'd"
"Out of " + $wsRows.Count + " total rows in the DMCertVerWorksheet.xlsx , " + $WSRDRecords.Count + " are designated as 'WSRD' tasks.  Of those, " + $WSRDTasksCERTdArray.Count + " had a discernable CERT Date"
"Out of " + $wsRows.Count + " total rows in the DMCertVerWorksheet.xlsx , " + $verreecords.Count + " are designated as 'VER List' tasks.  Of those, " + $verListArray.Count + " had a discernable CERT Date"
"Out of " + $wsRows.Count + " total rows in the DMCertVerWorksheet.xlsx , " + $badDMCrecords.COUNT + " have bad DMC codes."

Remove-Item -Path "$exportFolder\$reportName" -Force -ErrorAction SilentlyContinue
$masterList    | Export-XLSX -Path "$exportFolder\$reportName" -Header  Manual,VerListValue,Type,PBTask,DMC,TechName,InfoName,Rel_5,Rel_6,Rel_7,Rel_8,Rel_9,Rel_10,BOE_CertificationType,USAF_VerificationType,RFU_List -WorksheetName "Verifications"
#$badDMCrecords | Export-XLSX -Path "$exportFolder\$reportName" -Header  Type,PBTask,certComp,verListDMC,verList,verComp -WorksheetName "Bad DMCs" 
#$verreecords   | Export-XLSX -Path "$exportFolder\$reportName" -Header  Type,PBTask,certComp,verListDMC,verList,verComp -WorksheetName "Ver List Tasks" 
#$WSRDRecords   | Export-XLSX -Path "$exportFolder\$reportName" -Header  Type,PBTask,certComp,verListDMC,verList,verComp -WorksheetName "WSRD Tasks" 
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
$ed
"Report now available at this location:`r`n$exportFolder\$reportName"