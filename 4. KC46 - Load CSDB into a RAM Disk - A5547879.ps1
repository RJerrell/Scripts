cls
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
Import-Module -Name "KC46Common" -Verbose -Force
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'

$destDrive = "F:"
$ramdrive = "F:"
$fldr = "$destDrive\KC46 Staging\Production\Manuals"
Function Store-FileInRamDrive
{
    param([string] $filename, [string] $destination)
    Copy-Item -Force -Path $filename -Destination $destination -ErrorAction Stop 
}

WorkFlow Clear-Everything
{
    Param([string] $path)

    $dirs = gci -Path $path -Directory

    foreach -parallel ($dir in $dirs)
    {
        $path = $dir.FullName            
        if(Test-Path $path)
        {
            Remove-Item -Force -Recurse -Path $dir.FullName -ErrorAction Stop
        }
    }
}

Function Load-Data
{
    param([string] $destDrive, [string] $ramdrive)

    $sourceBase = "C:\KC46 Staging\Production\Manuals"
    $destinationBase = "$destDrive\KC46 Staging\Production\Manuals"
    
    [string[]] $PubList = @("KC46","ABDR", "ACS","ARD","ASIP","AMM","FIM","IPB","LOAPS","NDT", "NDTS","SIMR","SPCC","SSM","SWPM","TC","WUC","WDM")
    
    foreach ($pub in ($PubList |Sort-Object))
    {
        $sourceDM = "$sourceBase\$pub"
        $destinationDM = "$destinationBase\$pub"
        $Spath = "$sourceDM\S1000D\SDLLIVE"
        $ISpath = "$sourceDM\Illustrations\Illustrations"
        $Dpath = "$destinationDM\S1000D\SDLLIVE"
        $IDpath = "$destinationDM\Illustrations\Illustrations"

        if (!(Test-Path -Path $IDpath))
        {
            md $IDpath
        }
        if (!(Test-Path -Path $Dpath))
        {
            md $Dpath
        }

        Set-RamDriveFiles -SPath $Spath -Dpath $Dpath
        Set-RamDriveIllustrations -ISpath $ISpath -IDPath $IDpath
        
    }

    $sfmPath = "C:\KC46 Staging\production\manuals\kc46\s1000d\frontmatter"
    $dfmPath = "$destDrive\KC46 Staging\production\manuals\kc46\s1000d\frontmatter"
    if(!(Test-Path -Path $dfmPath))
    {
        md $dfmPath
    }

    Copy-Item -Path "$sourceBase\ACRONYM_MasterList.txt"  -Destination $destinationBase -Verbose -Force
    Copy-Item -Path "$sfmPath\DMC-1KC46-A-00-00-0000-00A0K-010A-A.xml"  -Destination $dfmPath -Verbose -Force
    Copy-Item -Path "$sfmPath\DMC-1KC46-A-00-00-0000-01A0K-018A-A.xml"  -Destination $dfmPath -Verbose -Force

    robocopy /e /mir "C:\Program Files\XyEnterprise" "$destDrive\Program Files\XyEnterprise"
    robocopy /e /mir "$sfmPath\release dml set\Highlights" "$dfmPath\release dml set\Highlights"
    robocopy /e /mir "C:\KC46 Staging\Production\Manuals\KC46\S1000D\Master PM" "$destDrive\KC46 Staging\Production\Manuals\KC46\S1000D\Master PM"
    robocopy /e /mir "C:\KC46 Staging\Production\Manuals\KC46\Illustrations\Illustrations" "$destDrive\KC46 Staging\Production\Manuals\KC46\Illustrations\Illustrations"
    robocopy /e /mir "C:\LiveContentData\common\config - KC46 Only" "$destDrive\LiveContentData\common\config - KC46 Only"
    robocopy /e /mir "C:\LiveContentData\publications" "$destDrive\LiveContentData\publications"
    
    Reset-wietmsdFiles -destDrive $destDrive -ramdrive $ramdrive
}

Function Set-RamDriveIllustrations
{
Param([string]  $ISpath , [string] $IDPath )
    $fileListIllustrations = gci -Path $ISpath -Filter *.cgm
    foreach ($f in $fileListIllustrations)
    {
        Copy-Item -Force -Path $f.FullName -Destination $IDPath  -ErrorAction Stop 
    }
}

Function Set-RamDriveFiles
{
Param([string]  $SPath , [string] $Dpath )
    $fileBases = @()
    $filesToCopy = @()

    #Deal with the publication modules first
    $pmcFiles = gci -Path $Spath -Filter PMC*.xml | Sort-Object -Descending | Select -First 1
    Copy-Item -Force -Path $pmcFiles[0].Fullname -Destination $Dpath  -ErrorAction Stop 
    #Store-FileInRamDrive -filename $pmcFiles[0].Fullname -destination $Dpath
    
    # Now manage the DMs
    $allFiles = gci -Path $Spath -Filter DMC*.xml
    foreach ($file in $allFiles)
    {
        $startPos = $file.Name.IndexOf("_")
        if($startPos -lt 1)
        {
            $startPos = $file.Name.ToLower().IndexOf(".xml")
        }

        $fileBase = $file.Name.Substring(0,$startPos)

        if(!($fileBases -contains $fileBase))
        {                
            $fileBases += $fileBase
            $filterString = $fileBase + "*"
            #$fileCount = gci -Path $Spath -Filter $fileBase
            $fl = gci -Path $Spath -Filter $filterString |Sort-Object -Descending | Select -First 1
            $filesToCopy += $fl[0].FullName
        }
    }
    foreach ($f in $filesToCopy)
    {
        Copy-Item -Force -Path $f -Destination $Dpath  -ErrorAction Stop    
    }
    
}

Function Reset-wietmsdFiles
{
    param([string] $destDrive, [string] $ramdrive)
    $folder = "c:\LiveContentData\common\config - KC46 Only"
    $wietmsdFiles = gci -Path "$folder\wietmsd*.xml" -Recurse
    if( !(Test-Path -Path $folder ))
    {
        md $folder
    }  
    foreach ($wietmsdFile in $wietmsdFiles)
    {
        #attrib -r $wietmsdFile.FullName
        Set-ItemProperty -Path $wietmsdFile.FullName -Name IsReadOnly -Value $false -Force -Verbose
        $content = [System.IO.File]::ReadAllText($wietmsdFile.FullName)
        $X = $content.Replace("C:",$ramdrive)
        $xml1 = New-Object -TypeName System.Xml.XmlDataDocument
        $xml1.LoadXml($X)
        $dFName = $wietmsdFile.FullName.Replace("C:",$ramdrive)
        Set-ItemProperty -Path $dFName -Name IsReadOnly -Value $false -Force -Verbose
        $xml1.Save($dFName)
        $dFName
    }
}

if( !(Test-Path -Path $fldr ))
{
    md $fldr
}

Clear-Everything  -path $fldr

Load-Data -destDrive $destDrive -ramdrive $ramdrive

$propFiles = gci -Path "C:\LiveContentData\publications\properties.xml" -Recurse

foreach ($propFile in $propFiles)
{
        $propFile.Attributes = "Archive"
        $fdriveName = $propFile.FullName.Replace("C:",$ramdrive)
        $dirName = $propFile.DirectoryName.Replace("C:",$ramdrive)
        if( !(Test-Path -Path $dirName ))
        {
            md $dirName
        }
        $content = [System.IO.File]::ReadAllText($propFile.FullName)
        $X = $content.Replace("C:","$ramdrive")

        $destinationName = $propFile.FullName.Replace("C:",$destDrive)

        $xml = New-Object -TypeName System.Xml.XmlDataDocument
        $xml.LoadXml($X)
        $xml.Save($destinationName)
    }
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"