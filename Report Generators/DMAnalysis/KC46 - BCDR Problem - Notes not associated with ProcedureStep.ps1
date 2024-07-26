<#
    Title: 
    Author: Roger Jerrell
    Date Created: Get-Date 
    Purpose: 
    Description of Operation: 
Description of Use:
#>
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
# *****************************************************************************************************
$error.clear()
$reportpath = "C:\KC46 Staging\scripts\Report Generators\Outputs"
$environment = "Production"
$KC46DataRoot = "F:\KC46 Staging"
$source_BaseLocation = "$KC46DataRoot\$environment\Manuals"
#[string[]] $books   = @("KC46","LOAPS","AMM","ACS", "ARD","FIM","IPB","NDT","TC","SSM","SWPM","WUC","SPCC","SIMR")
[string[]] $books   = @("KC46","AMM","ARD","FIM","NDT","TC")
$objsec = @()
foreach ($book in $books)
{
    $path = "$source_BaseLocation\$book\S1000D\SDLLIVE\DMC*.*"
    
    $files = gci -Path $path

    foreach ($file in $files)
    {
        $dm = New-Object System.Xml.XmlDocument
        $dm.Load($file.FullName)
        $psnotes = $dm.DocumentElement.SelectNodes("//proceduralStep/note")
        foreach ($psnote in $psnotes)
        {            
            if($psnote.ParentNode.ChildNodes.Count -eq 1)
            {
                if($psnote.ParentNode.ChildNodes[0].Name -eq 'note')
                {
                    $noteText = $psnote.InnerText
                    $obj = [pscustomobject][ordered]@{Manual=$book;DMC=$file.Name;NoteText=$noteText;}
                    $objsec += $obj
                    $obj = $null
                }
            }
        }
    }
}

$prop1 = @{Expression='Manual'; Ascending=$true }
$prop2 = @{Expression='DMC'; Ascending=$true }

$objsec.GetEnumerator() | Sort-Object -Property $prop1, $prop2| Export-Csv "$reportpath\UnassociatedNotes.csv" -NoTypeInformation -Encoding UTF8
"$reportpath\UnassociatedNotes.csv"

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"