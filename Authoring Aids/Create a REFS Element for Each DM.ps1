
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
$env:PSModulePath

#Get-Module -ListAvailable
Import-Module -Name "KC46Common" -Verbose -Force
# *****************************************************************************************************
[string[]] $PubList   = @("ABDR","ACS","LOAPS","SPCC", "SIMR","WUC")
[string[]] $PubList   = @("ABDR")
foreach($pub in $PubList)
{
    $folderPath = "C:\KC46 Staging\Dev\Manuals\$pub\S1000D\SDLLIVE\DMC*.XML"
    $files = gci -Path $folderPath
    $refsArray = @()

    foreach ($file in $files)
    {
        $dirty = $true
        $FName = $file.FullName
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.Load($file.FullName)
        $contentDMRefCollection = $xmlDoc.SelectNodes("/dmodule/content//dmRef")
        $contentRefsNode = $xmlDoc.SelectNodes("/dmodule/content/refs/dmRef")
        $contentRefsNodeRoot = $xmlDoc.SelectNodes("/dmodule/content/refs")
        $contentNode = $xmlDoc.SelectNodes("/dmodule/content")

        #"Refs node  " + $contentRefsNode.Count
        #"All dmRefs  " + $contentDMRefCollection.Count
        if($contentRefsNode.Count -gt 0)
        {
            # Drop the contents of the refs node
            foreach($node in $contentRefsNodeRoot.ChildNodes)
            {
                $contentRefsNodeRoot.RemoveChild($node)
            }
        }
        $xmlDoc.dmodule.content.RemoveChild($contentRefsNodeRoot)

        if($contentDMRefCollection.Count -gt 0)
        {
            $dirty = $true
            # Add entries to the REFS node
            $rnode = $xmlDoc.SelectSingleNode("/dmodule/content/refs")
            if($rnode -eq $null)
            {
                "ref node is null"
                $rnode = $xmlDoc.CreateElement("refs")
                $contentNode.AppendChild($rnode)
            }

            foreach ($contentDMRef in $contentDMRefCollection)
            {
                if($refsArray -notcontains $contentDMRef.OuterXml)
                {
                    $clonedNode = $contentDMRef.Clone()
                    $rnode.AppendChild($clonedNode)
                    $refsArray += $contentDMRef.OuterXml
                }
            }
            $xmlDoc.dmodule.content.InsertBefore($rnode, $xmlDoc.dmodule.content.FirstChild)
            $xmlDoc.dmodule.content.refs.OuterXml
            

        }
        if($dirty)
        {
            # attrib -R $file.FullName;
            Set-ItemProperty -Path $file.FullName -Name IsReadOnly -Value $false -Force -Verbose
            Save-PrettyXML -FName $FName -xmlDoc $xmlDoc
            attrib +R $file.FullName;
        }        
    }
}

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.TotalDays
"Total Hours to complete:`t" + $x.TotalHours
"Total Minutes to complete:`t" + $x.TotalMinutes
"Process completed"
