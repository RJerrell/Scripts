CLS
        [string[]] $PubList   = @("KC46","AMM","ARD","FIM","NDT","SSM", "TC","WDM")
        
        $pmcArray = @()

        foreach ($Pub in $PubList |Sort-Object)
        {
            $path2PMC = "C:\KC46 Staging\production\Manuals\$pub\S1000D\SDLLIVE\PMC*.xml"

            $pmcS = gci -Path $path2PMC | Sort-Object -Descending | Select -First 1
            $pmcArray += $pmcS.NAME
        }
         $pmcArray