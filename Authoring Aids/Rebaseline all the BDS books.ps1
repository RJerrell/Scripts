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
$env:Path = $env:Path + ";C:\PROGRAM FILES (X86)\MICROSOFT VISUAL STUDIO 12.0\COMMON7\IDE;"
if(!($env:Path.ToUpper() -contains "C:\PROGRAM FILES (X86)\MICROSOFT VISUAL STUDIO 12.0\COMMON7\IDE;"))
{
    $env:Path = $env:Path + ";C:\PROGRAM FILES (X86)\MICROSOFT VISUAL STUDIO 12.0\COMMON7\IDE;"
}
if ( (Get-PSSnapin -Name Microsoft.TeamFoundation.PowerShell -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PSSnapin Microsoft.TeamFoundation.PowerShell
}

[string[]] $PubList   = @("ABDR","ACS","LOAPS","SIMR","SPCC", "WUC")
[string[]] $PubList   = @("SWPM")
$basePath = "C:\KC46 Staging\Production\Manuals"
$issueNumber = "001"
$inWork = "00"
$languageIsoCode = "sx"
$countryIsoCode = "US"
Function Reset-DataModuleIssueAndInwork
{
    Param([System.IO.FileInfo] $file)
    $oName = $file.Name
    $FName = $file.FullName
    $folder = $file.DirectoryName
    $dmXML = New-Object System.Xml.XmlDocument
    $dmXML.Load($FName)
    $dmXML.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.issueNumber = $issueNumber
    $dmXML.dmodule.identAndStatusSection.dmAddress.dmIdent.issueInfo.inWork = $inWork
    $dmXML.dmodule.identAndStatusSection.dmAddress.dmIdent.language.languageIsoCode = $languageIsoCode
    $dmXML.dmodule.identAndStatusSection.dmAddress.dmIdent.language.countryIsoCode = $countryIsoCode
       
    # Set the issude dates to a standard date -- /dmodule/identAndStatusSection/dmAddress/dmAddressItems/issueDate
    $dmXML.dmodule.identAndStatusSection.dmAddress.dmAddressItems.issueDate.year  = "2017"
    $dmXML.dmodule.identAndStatusSection.dmAddress.dmAddressItems.issueDate.month = "04"
    $dmXML.dmodule.identAndStatusSection.dmAddress.dmAddressItems.issueDate.day   = "01"

    if($dmXML.dmodule.identAndStatusSection.dmStatus.HasAttributes -eq $false)
    {        
        $att = $dmXML.CreateAttribute("issueType")
        $att.Value = "new"
        $dmXML.dmodule.identAndStatusSection.dmStatus.Attributes.Append($att)
    }
    $indexof = $oName.ToUpper().IndexOf("_")
    if($indexof -lt 1)
    {
        $indexof = $oName.ToUpper().IndexOf(".XML")
    }
           
    $newName = $oName.Substring(0,$indexof) + "_" + $issueNumber + "-" + $inWork + "_" + $languageIsoCode.ToUpper() + "-" + $countryIsoCode.ToUpper() + ".xml"
    $newFullName = "$folder\$newName"

    attrib -R $file.FullName;
    #"Saving : `t" + $file.FullName  
    # Insure that no Unicode Byte Order Mark precedes the preamble in the file when saved to disk            
    # ************************ Critical that this encoding is used ****************************** 
    $encoding = New-Object System.Text.UTF8Encoding($False)
    # *******************************************************************************************
    $settings = New-Object System.Xml.XmlWriterSettings;
    $settings.Indent = $true;

    $settings.OmitXmlDeclaration = $false;
    $settings.NewLineOnAttributes = $false;
    $settings.Encoding = $encoding
    $settings.WriteEndDocumentOnClose = $false
    $settings.CheckCharacters = $true
    $settings.DoNotEscapeUriAttributes = $false

    #Create an XmlWriter to insure the output XML conforms to the settings above
    $writer = [System.Xml.XmlWriter]::Create($newFullName,$settings)
    # "Saving : `t" + $file.FullName 
    $dmXML.Save($writer);
    $newFullName
    $writer.Flush();
    $writer.Close();
    $writer.Dispose();
}
foreach ($Pub in $PubList)
{
    $path1 = "$basePath\$Pub\S1000D"
    $path2 = "$basePath\$Pub\S1000D\SDLLIVE"

    $files1 = gci -Path $path1 -Filter DMC*.xml
    $files2 = gci -Path $path2 -Filter DMC*.xml

    if($files1.count)
    {
        
        foreach ($file in $files1)
        {
            Reset-DataModuleIssueAndInwork -file $file
        }
    }
   
    if($files2.count)
    {
   
        foreach ($file in $files2)
        {
            Reset-DataModuleIssueAndInwork -file $file
        }
    }
}

# *****************************************************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)
"Publist: `t$PubList"
"Total Days to complete:`t" + $x.Days
"Total Hours to complete:`t" + $x.Hours
"Total Minutes to complete:`t" + $x.Minutes
"Total Seconds to complete:`t" + $x.Seconds
"Process completed"