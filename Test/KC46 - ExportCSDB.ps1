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

$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\KC46 Staging\Scripts\Modules;"
$env:PSModulePath

#Get-Module -ListAvailable
Import-Module -Name "KC46Common" -Verbose -Force

$release = "RELEASE 4"
<#
    IMPORTANT NOTE BEFORE EACH EXECUTION OF THIS SCRIPT
    INSURE THE RELEASE NUMBER IS ACCURATE AND WHAT YOU WANT IT TO BE CALLED BEFORE PROCEEDING
    BACKUP ANY CURRENT DAT IN THE TEMP FOLDER YOU NEED BEFORE RUNNING THE SCRIPT.
    THE TEMP FOLDER WILL BE DROPPED AND CREATED EACH TIME WITH A POTENTIAL LOSS OF EXPORTED DATA FROM
    PREVIOUS RUNS OF THE SCRIPT.
#>


$DM_Folder = "F:\LiveContentData\source\KC46\source_xml"
$Illustrations_Folder1 = "F:\LiveContentData\publications\KC46\figures"
$Illustrations_Folder2 = "F:\LiveContentData\publications\WIRING\figures"
$exportFolderBase = "C:\CSDB_EXPORTS"
$tempFolder = "$exportFolderBase\TEMP"
if(!(Test-Path -Path $exportFolderBase))
{
    md $exportFolderBase    
}

#Remove-Item -Recurse -Path $tempFolder -Force
#md $tempFolder

Workflow Push-CSDB
{
Param([string] $DM_Folder , [string] $tempFolder , [string] $release, [string] $Illustrations_Folder1, [string] $Illustrations_Folder2)
    parallel
    {
        Copy-FolderContents -source $DM_Folder -dest "$tempFolder\$release"
        Copy-FolderContents -source $Illustrations_Folder1 -dest "$tempFolder\$release\Illustrations"
        Copy-FolderContents -source $Illustrations_Folder2 -dest "$tempFolder\$release\Illustrations"
    }
}
Measure-Command{
    Push-CSDB -DM_Folder $DM_Folder -tempFolder $tempFolder -release $release -Illustrations_Folder1 $Illustrations_Folder1 -Illustrations_Folder2 $Illustrations_Folder2
}


# Cleanup the Illustrations by removing older versions of the ICN files
$files = gci -Path "$tempFolder\$release\Illustrations" -Exclude *.txt | Sort-Object -Property Name
foreach ($file in $files)
{
    # Inwork
    $fileName = $file.Name.Replace(".cgm", "")
    $fileParts = $fileName.Split("-")

    $rootName = $fileParts[0],$fileParts[1],$fileParts[2] -join "-"
    $p = $file.Directory.FullName + "\$rootname*"
    $fileset = gci -Path $p | Sort-Object -Property LastWriteTime -Descending
    if($fileset.Count -gt 1)
    {
       for ($i = 1; $i -lt $fileset.Count; $i++)
       { 
           Remove-Item -Path $fileset[$i].FullName -Verbose
       }
    }
    
    # Versions
    #Reset-Version -file $file
}


#Stop-Transcript
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"$PSCommandPath Process completed"