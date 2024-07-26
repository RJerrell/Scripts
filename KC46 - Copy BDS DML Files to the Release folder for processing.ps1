cls

[string[]] $PubList   = @("ABDR","ACS","LOAPS","SIMR","SSM","WUC")

$destination = "C:\KC46 Staging\Production\Manuals\KC46\S1000D\FrontMatter\Release DML Set\Release 9 DMLs"

foreach ($Pub in $PubList)
{
    #copy the DML for this book to the destination

    $from = "C:\KC46 Staging\Production\Manuals\$pub\S1000D\SDLLIVE"
    $files = gci -Path $from -Filter DML*.XML
    foreach ($file in $files)
    {
        $fileBasename = $file.BaseName
        $dml = gci -Path $file.Directory -Filter DML*.XML |  Sort-Object -Descending | Select-Object -First 1
        Copy-Item -Path $dml[0].FullName -Destination $destination -Force
    }
}