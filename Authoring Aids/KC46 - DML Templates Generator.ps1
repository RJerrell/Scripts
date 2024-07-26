
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
Import-Module -Name "KC46Common" -Verbose -Force
# *****************************************************************************************************
# Generate 3 DML files per book for 20 years in advance
<#  

    Taken from the top of a DML as a sample of how each dmlCode should look for each book
    <dmlCode dmlType="p" modelIdentCode="1KC46" senderIdent="AAA0##bookLetter##" seqNumber="##seqNumber##" yearOfDataIssue="##yearOfDataIssue##"/>

#>

$dmlTemplatePath = "C:\KC46 Staging\Scripts\Templates" 
$templateName = "DML-1KC46-81205-P-2017-00001_001-00_SX-US.xml"
$outputPath = "C:\KC46 Staging\Dev\Templates &  Guides\DML Templates"

$letters =@{ABDR="G";ACS="H";ASIP="C";SIMR="J";SPCC="Q";SWPM="U";WUC="D";LOAPS="M";NDTS="N";}

$years = (2017..2050)
# Clone the template to this machines %TEMP% location
Copy-Item -Path "$dmlTemplatePath\$templateName" -Destination $env:TEMP
remove-item -path $outputPath  -Force
md $outputPath 
$newBasename = ""
foreach ($letter in $letters.Values)
{
    $letter
    $j = 2017
    $seq = 00001
    $newBasename = $templateName.Replace("81205" , "AAA0" + $letter)
    $lastseq = ""
    for ($i = $j; $i -lt 2041; $i++)
    {        
        $newBasename = $newBasename.Replace($lastyear,$i.ToString())
        for($x = 1; $x -lt 4; $x++)
        { 
            $newSeq = $x.ToString().PadLeft(5,"0")
            $newBasename = $newBasename.Replace($lastseq , $newSeq)
            $xmlDML = New-Object System.Xml.XmlDocument
            $xmlDML.Load(" $env:TEMP\$templateName")
            
            $xmlDML.dml.identAndStatusSection.dmlAddress.dmlIdent.dmlCode.senderIdent = "AAA0$letter"
            $xmlDML.dml.identAndStatusSection.dmlAddress.dmlIdent.dmlCode.seqNumber = $newSeq
            $xmlDML.dml.identAndStatusSection.dmlAddress.dmlIdent.dmlCode.yearOfDataIssue = $i.ToString()
            #/dml/identAndStatusSection/dmlAddress/dmlAddressItems/issueDate
            $xmlDML.dml.identAndStatusSection.dmlAddress.dmlAddressItems.issueDate.year = $i.ToString()
            $xmlDML.dml.identAndStatusSection.dmlAddress.dmlAddressItems.issueDate.day = "01"
            if($x -eq 1)
            {
                $xmlDML.dml.identAndStatusSection.dmlAddress.dmlAddressItems.issueDate.month = "01"
            }
            elseif($x -eq 2)
            {
                 $xmlDML.dml.identAndStatusSection.dmlAddress.dmlAddressItems.issueDate.month = "05"
            }
            elseif($x -eq 3)
            {
                $xmlDML.dml.identAndStatusSection.dmlAddress.dmlAddressItems.issueDate.month = "09"
            }
            else
            {
                "Add code for values above 3"
                exit
            }
            Save-PrettyXML -FName "$outputPath\$newBasename" -xmlDoc $xmlDML

            $lastSeq = $x.ToString().PadLeft(5,"0")
        }
        $lastyear = $i
    }
}

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Total Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"