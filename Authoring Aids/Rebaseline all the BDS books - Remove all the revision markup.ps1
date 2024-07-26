<#
Title: 
Author: Roger Jerrell
Date Created: Get-Date 
Purpose: 
1. Baseline the issue numbers in all the BDS authored manuals
2. Sync the name to the Issue Number and Inwork
3. Remove all Reasons for Update all authoring highlights and revision markup.
 
Description of Use:
    Description: The element <issueInfo> contains the issue number of the data module.
    Markup element: <issueInfo> (M)
    Attributes:
    − issueNumber (M). Every approved issue of a data module must be allocated an
    incremented issue number which, with the data module code, uniquely identifies that
    instance of the data module. The initial issue must be numbered with the value "001",
    which must be incremented with every approved release of a data module.
    − inWork (M). This attribute gives the "inwork" number of the unreleased data module. It
    can be used for monitoring and control of intermediate drafts within a project. The initial
    inwork number is set to the value "01", and is incremented with every change to the
    unreleased data module.

This script: 

1. Sets the Issue Number to 001 and the InWork to 01 on all data modules in the path including the PMC.
2. Renames all the files in the path to reflect the naming convention that includes the name + IssueNumber + InWork + Language + Country
3. Remove all revision markup in all files.
4. Checks all the files into TFS as a start point for future edits and authoring.

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
# *****************************************************************************************************
cls
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
$env:PSModulePath
Import-Module -Name "KC46Common" -Verbose -Force

$env:Path = $env:Path + ";C:\PROGRAM FILES (X86)\MICROSOFT VISUAL STUDIO 12.0\COMMON7\IDE;"
if(!($env:Path.ToUpper() -contains "C:\PROGRAM FILES (X86)\MICROSOFT VISUAL STUDIO 12.0\COMMON7\IDE;"))
{
    $env:Path = $env:Path + ";C:\PROGRAM FILES (X86)\MICROSOFT VISUAL STUDIO 12.0\COMMON7\IDE;"
}
if ( (Get-PSSnapin -Name Microsoft.TeamFoundation.PowerShell -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PSSnapin Microsoft.TeamFoundation.PowerShell
}

[string[]] $PubList   = @("ABDR","ACS","LOAPS","SIMR","WUC")
#[string[]] $PubList   = @("WUC")
$basePath = "C:\KC46 Staging\Production\Manuals"

$issueNumber = "001"

$inWork = "01"

$languageIsoCode = "SX"

$countryIsoCode = "US"

Function Reset-DataModuleRevisionMarkup
{
    Param([System.IO.FileInfo] $file)
    $dirty = $false
    $oName  = $file.Name
    $FName  = $file.FullName
    $folder = $file.DirectoryName
    
    $dmXML = New-Object System.Xml.XmlDocument

    $dmXML.Load($FName)
    
    $dmStatusNode = $dmXML.SelectSingleNode("/dmodule/identAndStatusSection/dmStatus")
    
    $rfus = $dmStatusNode.SelectNodes('reasonForUpdate')       
    <# Remove all the reasonForUpdate tags in the header #>
    foreach ($rfu in $rfus)
    {
        $null = $dmStatusNode.RemoveChild($rfu)
        $dirty = $true
    }
        # Now get rid of all the attributes called "reasonForUpdateRefIds"
        $dmXML.OuterXml.ToString().Contains("changeMark")
        $dmXML.OuterXml.ToString().Contains("reasonForUpdateRefIds")
    if($dmXML.OuterXml.ToString().Contains("reasonForUpdateRefIds"))
    {
        $rfuIDNodes = $dmXML.dmodule.content.SelectNodes("//*[@reasonForUpdateRefIds!='']")
        foreach($node in $rfuIDNodes)
        {
            foreach($att in $node.Attributes)
            {
                if($att.Name -eq "reasonForUpdateRefIds")
                {
                    $null = $node.Attributes.Remove($att)
                }
            }
        }
        $dirty = $true
    }
    
    $cleanStringOriginal = $dmXML.OuterXml.ToString()
    $cleanString = ""
    if($dmXML.OuterXml.ToString().Contains("changeMark"))
    {        
        $cleanString = $cleanStringOriginal.Replace("changeMark=`"1`"","")
            $cleanString = $cleanString.Replace("changeMark=`"0`"","")
                $cleanString = $cleanString.Replace("changeType=`"modify`"","")
                    $cleanString = $cleanString.Replace("changeType=`"add`"","")
                                $cleanString = $cleanString.Replace("changeType=`"delete`"","")                                    
    }
       
    # Lastly, remove all the inline revision tags - whewww!
    if(($cleanString -ne "") -and ($cleanString -ne $cleanStringOriginal))
    {
        #Load a cleaned up version of the document
        $dmXML.LoadXml($cleanString)
        $dirty = $true
    }

    if($dirty)
    {
        attrib -R $file.FullName;
        Save-PrettyXML -FName $FName -xmlDoc $dmXML
        attrib +R $file.FullName;
    }
    
    <# Process the revision markup and remove it all #>
    #$sr = New-Object System.IO.StreamReader($FName)
    #$strDoc = $sr.ReadToEnd()
}
foreach ($Pub in $PubList)
{
    $path1 = "$basePath\$Pub\S1000D"
    $path2 = "$basePath\$Pub\S1000D\SDLLIVE"

    $files1 = gci -Path $path1 -Filter DMC*.xml | Sort-Object
    $files2 = gci -Path $path2 -Filter DMC*.xml | Sort-Object
    

    if($files1.count)
    {        
        foreach ($file in $files1)
        {
            Reset-DataModuleRevisionMarkup -file $file
        }
    }   
    if($files2.count)
    {
   
        foreach ($file in $files2)
        {
            Reset-DataModuleRevisionMarkup -file $file
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