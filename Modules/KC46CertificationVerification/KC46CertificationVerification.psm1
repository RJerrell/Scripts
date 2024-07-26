#region Public Functions / Workflows
$cvStatusList = @{}
function Reset-AllVerficationTags
{
   Param([string[]] $pubList , [string] $basePath) 
   foreach ($pub in $pubList)
   {
       $path1 = "$basePath\$pub\s1000d\dmc*.xml"
       $path2 = "$path1\sdllive\dmc*.xml"

   }
}
function Set-Certification
{
    Param([string] $dmc)
    Set-Verification
}

#endregion

#region Private Functions / Workflows
function Set-Verification
{
    Param([string] $dmc)
    "done"
}

function Get-AllCVStatus
{
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    $siteUrl = "https://collab.web.boeing.com/sites/KC46TankerTechPubs/reporting_techpubs"
    $listName = "Task List Merged"

    $context = New-Object -TypeName Microsoft.SharePoint.Client.ClientContext($siteUrl)
    $context.Credentials = New-Object System.Net.NetworkCredential("SVCKC46_SP_LOGGER", 'Boeing$1', "NW")

    $itemsesults = @()

    $list = $context.Web.Lists.GetByTitle($listName)
    $context.Load($list.Fields)
    $context.ExecuteQuery()

    foreach ($field in $list.Fields)
    {
        $field.InternalName
    }


    #$q = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(5)
    $q = New-Object Microsoft.SharePoint.Client.CamlQuery
    
    $q.ViewXml = "<View><ViewFields><FieldRef Name='Ver_x0020_Comp'/><FieldRef Name='DMC'/><FieldRef Name='PB_x002f_Task'/></ViewFields><RowLimit>2000</RowLimit></View>"
   
    
    # An array to hold all of the ListItems
    $items = $list.getItems($q)
    $context.Load($items)
    $context.ExecuteQuery()

    $position = $items.ListItemCollectionPosition
    $cvDate = ""
    $dmc = ""
    $pbTask =""
    $csv = New-Object System.Text.StringBuilder
    $reportname = "c:\temp\CandVStatus-PS.csv"
    Remove-Item -Path $reportname -Verbose -ErrorAction SilentlyContinue
    # Get Items from the List until we reach the end
    do
    {
        $q.ListItemCollectionPosition = $position;
        $items = $list.GetItems($q);
        $context.Load($items)
        $context.ExecuteQuery()

        foreach($item in $items)
        {
            if ($item["PB_x002f_Task"] -ne $null -and (! $item["PB_x002f_Task"].ToString().Contains("PAGEBLOCK")))
            {
                $pbTask = $item["PB_x002f_Task"].ToString();
                $pbTask = $pbTask.Replace("Task ", "");
                if ($item["DMC"] -ne $null)
                {
                    $dmc = $item["DMC"].ToString();
                }

                if ($item["Ver_x0020_Comp"] -ne $null)
                {
                    $cvDate = $item["Ver_x0020_Comp"].ToString();
                    $newLine = "$pbTask`t$dmc`t$cvDate"
                   
                    $csv.AppendLine($newLine);
                }


                
            }
        }
    }
    While($position -ne $null);
        
    try 
	{
        System.IO.File.Delete($reportname);
            
	}
	finally
	{
        System.IO.File.WriteAllText($reportname, $csv.ToString());
	}

}
<#

#>
<#

#>
<#

#>
<#

#>
#endregion

<# Exports #>
#export-modulemember -function Reset-AllVerficationTags, Get-AllCVStatus