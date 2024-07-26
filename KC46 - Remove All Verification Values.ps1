CLS
$ErrorActionPreference = "Stop"
$error.Clear()
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force					
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
# *****************************************************************************************************
$parserDM  = New-object -TypeName S1000D.DataModule_401
$parserDML  = New-Object -TypeName S1000D.DataModuleList
$parserCOM = New-Object -TypeName S1000D.CommonFunctions
$parserPM    = New-object -TypeName S1000D.PublicationModule_401
# Add your code here
# $parserDM.Set_QualityAssurance($boeValue,$usafValue)    
$path = "c:\KC46 Staging\Production\Manuals"
#$path = "C:\KC46 Staging\Dev\Manuals"
$boeValue = "tabtop"
$usafValue = ""
$I = 1

[string[]] $PubList   = @("KC46","ABDR","ACS","AMM","ARD","ASIP","FIM","IPB","LOAPS","NDT","NDTS","SIMR","SPCC","TC","WUC","SSM","SWPM", "WDM") | Sort-Object

Function Set-Values
{
    Param([System.IO.FileInfo[]] $files)
    $fc = $files.Count
    foreach ($file in $files)
    {
        $FName = $file.FullName
        $FName
        Set-ItemProperty -Path $FName -Name IsReadOnly -Value $false -Force
        
        $parserDM.ParseDM($FName)
        $Book = $parserDM.AssociatedBook

        $boeCurrentValue = ($parserDM.IdentAndStatusSection.dmStatus.qualityAssurance.firstVerification.verificationType).Trim()
        if($boeCurrentValue.Length -gt 0)
        {
            $parserDM.Set_QualityAssurance($boeCurrentValue,$usafValue) 
        }
        else
        {
            $parserDM.Set_QualityAssurance($boeValue,$usafValue)
        }
        $I ++
        "$Book : $I of $fc"
    }
}

foreach ($Pub in $PubList)
{    
    $pathToManual1 = "$path\$pub\s1000d\"
    $pathToManual2 = "$path\$pub\s1000d\sdllive"
    $files1 = gci -Path $pathToManual1  -Filter DMC*.XML
    Set-Values -files $files1
    
    $files2 = gci -Path $pathToManual2  -Filter DMC*.XML
    Set-Values -files $files2
}
# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"					
					