cls

$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
Import-Module -Name "KC46Common" -Verbose -Force

# CHANGE THIS LOCATION FOR THE REPORT FILE
$reportPath = "C:\TEMP\VALIDATEDMREFS.CSV"
<# Set the basic path to the data with the data in a folder structure like this
 
    C:\KC46 Staging\Production\Manuals\AMM\S1000D\SDLLIVE (DMC and PMC files here
                                    \ARD\S1000D\SDLLIVE (DMC and PMC files here
                                    \FIM\S1000D\SDLLIVE (DMC and PMC files here
                                    \NDT\S1000D\SDLLIVE (DMC and PMC files here
                                    \SRM\S1000D\SDLLIVE (DMC and PMC files here
                                    \SSM\S1000D\SDLLIVE (DMC and PMC files here                                            
                                    \WDM\S1000D\SDLLIVE (DMC and PMC files here
#>

$basePath = "C:\KC46 Staging\Production\Manuals"
 [string[]] $PubList = @("AMM","ARD", "FIM", "IPB", "LOAPS", "MOM", "NDI", "NDT", "SPCC", "SRM", "SSM", "SWPM", "WUC")
[string[]] $PubList = @("ABDR","WUC")

foreach( $pub in $PubList)
{
    $a1=@()
    $b1=@()

    $manualBasePath = "$basePath\$pub\S1000D\SDLLIVE"

    $pmcPath = "$manualBasePath\pmc*.xml"

    $pmcFiles = gci -Path $pmcPath | Sort-Object -Descending | Select-Object -First 1

    $pmc = New-Object System.Xml.XmlDocument

    $pmc.Load($pmcFiles[0].Fullname)

    $dmRefs = $pmc.SelectNodes("//dmRef");


    foreach( $dmRef in $dmRefs )
    {
        $filePref = "DMC";
        $RefFilename = Get-FilenameFromDMRef -dmRef $dmRef -filePref $filePref
        $bookReferenced = Get-DocTypeFromDMC -dc $RefFilename
        $fullFilePath2ToRefDoc = "$basePath\$bookReferenced\S1000D\SDLLIVE"

        $fs = gci -path "$fullFilePath2ToRefDoc\$RefFilename`*" |Sort-Object -Descending |Select-Object -First 1

        if($a1 -notcontains $fs[0].FullName.ToUpper())
        {        
            $a1 += $fs[0].FullName.ToUpper()
        }
        
        
        if($fs.Count -eq 0)
        {       
            "Missing data module that begins with this name :`t$basename"
        }
    }
    
    <# 
        Now, we will process each S1000D folder and compare the dmCode to the PMC for that
        manual to see if there are files
        that exist on disk that are not referenced.
    #>
    $filePath = "$basePath\$pub\S1000D\SDLLIVE\DMC*.XML"

    $files = Get-ChildItem -Path $filePath | Sort-Object -Descending
    foreach($file in $files)
    {
        $idx = $file.FullName.IndexOf("_")
        $bName = $file.FullName.Substring(0, $idx)
        if($b1 -notcontains $bName)
        {            
            $b1 += $bName
        }
    }
" ********************************  START OF $pub REPORT          **********************************`n"
    "TOTAL PMC ENTRIES`t`t`t`t: " + $a1.Count
    "TOTAL DATA MODULES ON DISK IN CSDB`t`t`t`t: " + $b1.Count
    " DIFFERENCE REPORT BETWEEN THE 2 LISTS: PMC Entries versus CSDB entries"
    "<= indicates a reference in the PMC that does not have an equivalent file on disk`n=> indicates a file on disk that is not referenced in the PCM`n"    


Compare-Object -ReferenceObject ($a1 | Sort-Object) -DifferenceObject ($b1 | Sort-Object) |Export-Csv $reportPath
   
		"TOTAL PMC ENTRIES`t`t`t`t: " + $a1.Count
    "TOTAL DATA MODULES ON DISK IN CSDB`t`t`t`t: " + $b1.Count
    " DIFFERENCE REPORT BETWEEN THE 2 LISTS: PMC Entries versus CSDB entries"
    "<= indicates a reference in the PMC that does not have an equivalent file on disk`n=> indicates a file on disk that is not referenced in the PCM`n"

" ********************************    END OF $pub REPORT          **********************************`n"
}