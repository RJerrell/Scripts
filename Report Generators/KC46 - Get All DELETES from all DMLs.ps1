cls
$sd = Get-Date
$ErrorActionPreference = "Stop"
$error.Clear()

# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
$parserDML = new-object -TypeName S1000D.DataModuleList

$deletes = @()


for ($i = 5; $i -lt 12; $i++)
{ 
    $dmls = gci -Path "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set\Release $i DMLs" -Filter DML*.* | Sort-Object
    foreach ($dml in $dmls)
    {
        $parserDML.ParsePM($dml.FullName)
        $e = $parserDML.DMEntries

        foreach ($entry in $e)
        {
            if($entry.dmEntryType -eq "d")
            {
                $fname = Get-FilenameFromDMRef -dmRef $entry.dmRef -filePref "DMC-"
                if(!($deletes.Contains($fname)))
                {
                    $deletes += $fname + "`t|`t" + $dml.Name
                }
            }
        }        
    }
}

$deletes | Sort-Object -Descending