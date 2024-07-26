
cls
$sd = Get-Date
$ErrorActionPreference = "Stop"
$error.Clear()
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
# *****************************************************************************************************
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force 
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'

Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$parserDML = new-object -TypeName S1000D.DataModuleList
$parserDMC = new-object -TypeName S1000D.DataModule_401
# *****************************************************************************************************
# Get a litsing of every deleted data module
$path2DMLs = "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set"
$dmlArray =@()
$dmlArray += "$path2DMLs\Release 5 DMLs\DML-1KC46-AAA0A-P-2016-00004.xml"
$dmlArray += "$path2DMLs\Release 6 DMLs\DML-1KC46-AAA0A-P-2017-00001_001-00_SX-US.xml"
$dmlArray += "$path2DMLs\Release 7 DMLs\DML-1KC46-AAA0A-P-2017-00002_001-00_SX-US.xml"
$dmlArray += "$path2DMLs\Release 8 DMLs\DML-1KC46-AAA0A-P-2017-00003_001-00_SX-US.xml"
$dmlArray += "$path2DMLs\Release 9 DMLs\DML-1KC46-AAA0A-P-2018-00001_001-00_SX-US.xml"
$dmlArray += "$path2DMLs\Release 10 DMLs\DML-1KC46-AAA0A-P-2018-00002_001-00_SX-US.xml"
$dml = New-Object System.Xml.XmlDocument

$deletedItemArray = @()

foreach ($dmlpath in $dmlArray)
{   
    $dml = $parserDML.GetModule($dmlpath)
    #$dml.Load( $dmlpath )
    $deletedItems = $dml.SelectNodes("/dml/dmlContent/dmEntry[@dmEntryType=`"d`"]")
    foreach ($deletedItem in $deletedItems)
    {
        $tName = ""
        $iName = ""
        $fn = Get-FilenameFromDMRef -dmRef $deletedItem.dmRef -filePref "DMC"
        $files= gci -Path "C:\KC46 Staging\Production\Manuals\AMM\S1000D\SDLLIVE\$fn`*.xml" | Sort-Object | select -First 1
        $fullName = $files[0].FullName
        $parserDMC.ParseDM($fullName)

        $tname = $parserDMC.TechName
         $iName = $parserDMC.InfoName

        $deletedItemArray += New-Object -TypeName PSObject -Property @{
            DMC       = $fn
            TaskName  = $deletedItem.dmRef.title;                                      
            InfoName  = $iName;
            Release   = $dmlpath.Replace("C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set", "");
        } | Select DMC,TaskName,InfoName,Release
    }
}
$outputFullPath = "C:\KC46 Staging\Scripts\Report Generators\Outputs\BoilerPlateReports\All Deleted Data Modules.xlsx"
Remove-Item -Path $outputFullPath -Force -ErrorAction SilentlyContinue
$deletedItemArray | Sort-Object -Property Release,DMC | Export-XLSX -Path $outputFullPath -ClearSheet -WorksheetName "DeletedTasks"

$deletedItemArray.Count
# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"