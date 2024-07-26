Clear-Host

# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
# *****************************************************************************************************
function Update-ProductionCSDB
{
    Param([string] $sourcePath , [string] $targetPath, [string] $fullCSDBPath)
    [string[]] $PubList   = @("ABDR","ACS","ASIP","LOAPS","NDTS","SIMR","SPCC","WUC")
    foreach($Pub in $PubList)
    {
        $sdmPath   = "$sourcePath\$Pub\S1000D\SDLLive"
        $sIllPath  = "$sourcePath\$Pub\Illustrations\Illustrations"        
                
        $tdmPath1   = "$targetPath\$Pub\S1000D\SDLLive"
        $tdmPath2   = "$fullCSDBPath\$Pub\S1000D\SDLLive"

        $tIllPath1  = "$targetPath\$Pub\Illustrations\Illustrations"       
        $tIllPath2  = "$fullCSDBPath\$Pub\Illustrations\Illustrations" 
        if(!(Test-Path -Path $tdmPath1))
        { New-Item -Path $tdmPath1 -ItemType Directory}
        if(!(Test-Path -Path $tdmPath2))
        { New-Item -Path $tdmPath2 -ItemType Directory}
        if(!(Test-Path -Path $tIllPath1))
        { New-Item -Path $tIllPath1 -ItemType Directory}
        if(!(Test-Path -Path $tIllPath2))
        { New-Item -Path $tIllPath2 -ItemType Directory}


        $files = Get-ChildItem -Path $sdmPath -Filter *DMC*.XML | Sort-Object -Descending
        $pmcfiles = Get-ChildItem -Path $sdmPath -Filter *PMC*.XML | Sort-Object -Descending | Select-Object -First 1
        
        Remove-Item -Force -Recurse -Path "$tdmPath1\PMC*.XML"

        Copy-Item -Path $pmcfiles[0].FullName -Destination $tdmPath1  -ErrorAction Stop 
        Copy-Item -Force -Path $pmcfiles[0].FullName -Destination $tdmPath2  -ErrorAction Stop 
        $processedFiles = @()

        foreach ($file in $files)
        {
            $fileParts = $file.Name.ToLower().Replace(".xml","").Split("_")
            $basePart = $fileParts[0].ToUpper()
            if(! $processedFiles.Contains($basePart))
            {
                # Remove all the files with the same base name from the current CSDB and add the latest version
                Remove-Item -Force -Recurse -Path "$tdmPath1\$basePart`*.xml"
                $basefiles = Get-ChildItem -Path $sdmPath -Filter "$basePart`*.xml" | Sort-Object -Descending | Select -First 1
                Copy-Item -Path $basefiles[0].Fullname -Destination $tdmPath1  -ErrorAction Stop 
            
                # Overwrite (-Force) same named files in the CSDB
                Copy-Item -Force -Path $basefiles[0].Fullname -Destination $tdmPath2  -ErrorAction Stop 
                $processedFiles += $basePart
            }
        }

        # Just copy all the Illustrations to both destinations
        

        if(!(Test-path -Path $tIllPath1 ))
        {
            mkdir $tIllPath1
        }
        else
        {
            Remove-Item -Force -Path $tIllPath1 -Recurse -Verbose
            mkdir $tIllPath1
        }

        $allGraphics = Get-ChildItem -Path "$sIllPath\*" -Include *.cgm,*.png
        foreach ($Graphic in $allGraphics)
        {
            Copy-Item -Force -Path $Graphic.FullName -Destination $tIllPath1
            Copy-Item -Force -Path $Graphic.FullName -Destination $tIllPath2
        }

        # Remove the doctype information from the header
        Reset-DoctypesPriorToBuild -pathtoXml $tdmPath1
        Reset-DoctypesPriorToBuild -pathtoXml $tdmPath2
$Pub
        # Applies acronyms and data markings
        Update-1Folder -basepath $tdmPath1
        Update-1Folder -basepath $tdmPath2
    }
}

$devPath = "C:\KC46 Staging\Dev\Manuals"
$prodPath = "C:\KC46 Staging\Production\Manuals"
$fullCSDBPath = "C:\KC46 Staging\Production\FullCSDB"

Update-ProductionCSDB -sourcePath $devPath -targetPath $prodPath -fullCSDBPath $fullCSDBPath

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"