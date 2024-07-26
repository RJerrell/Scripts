cls
$perBook = @()
$perBook += "Manual|DM Count|ICN Count|DM Ref Total|ICN Ref Total|Unique ICN Refs"
[string[]] $PubList   = @("KC46", "ABDR", "ACS", "AMM", "ARD", "ASIP", "FIM", "IPB", "LOAPS", "NDT", "SIMR",  "SSM", "SRM", "SWPM", "TC", "WUC", "WDM")

foreach ($Pub in $PubList)
{
    $bookDMRefcounter = 0
    $bookICNcounter = 0    
    $dmcCollection = gci -Path "C:\KC46 Staging\Production\Manuals\$Pub\S1000D\SDLLIVE\DMC*.XML"
    foreach ($dm in $dmcCollection)
    {
        $dmXML = New-Object System.Xml.XmlDocument
        $dmXML.Load($dm.FullName)
        
        $dmRefsDescription = $dmXML.SelectNodes("/dmodule/content/description//dmRef")
        $dmRefsProcedure = $dmXML.SelectNodes("/dmodule/content/procedure//dmRef")
        $dmGraphicsCollection = $dmXML.SelectNodes("/dmodule/content//graphic")

        # Get unique ICN information
        foreach ($dmGraphic in $dmGraphicsCollection)
        {
            if($bookICNArray.Contains($dmGraphic.infoEntityIdent))
            {

            }
            else
            {
                $bookICNArray += $dmGraphic.infoEntityIdent
            }
        }

        # Add to the entire IETM ICN Totals
        $TotalICNCounter += $dmGraphicsCollection.Count

        # Add DM count for this PUB to the book Totals
        $bookDMRefcounter += $dmRefsDescription.Count
        $bookDMRefcounter += $dmRefsProcedure.Count
        
        # Add ICN count for this PUB to the book Totals
        $bookICNcounter += $dmGraphicsCollection.Count        
    }

    $uCount = $bookUNIQUEICNArray.Count
    $bookDMCount = $dmcCollection.Count
    
    # $perBook += "Manual|DM Count|ICN Count|DM Ref Total|ICN Ref Total|Unique ICN Refs"
    $perBook += "$Pub,$bookDMCount|$bookICNcounter,$uCount"

}
$perBook | Out-File -FilePath "C:\KC46 Staging\Scripts\Report Generators\Outputs\CSDB Metrics.csv"