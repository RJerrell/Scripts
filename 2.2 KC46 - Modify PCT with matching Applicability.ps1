CLS
$ErrorActionPreference = "Stop"
$error.Clear()
# Include all common variables
. 'C:\KC46 Staging\Scripts\Common\KC46CommonVariables.ps1'
Import-Module -Name "KC46Common" -Verbose -Force
Import-Module -Name "PSExcel" -Verbose -Force					
Add-Type -Path "C:\KC46\S1000D\Deployments\S1000D_Parser\Parser.dll" -Verbose
# *****************************************************************************************************
$parserDML   = New-Object -TypeName S1000D.DataModuleList
$parserCOM   = New-Object -TypeName S1000D.CommonFunctions
$parserPM    = New-object -TypeName S1000D.PublicationModule_401

$outputFolder = "C:\KC46 Staging\Scripts\Report Generators\Outputs\Effectivities"
$reportName = "Effectivity exceptions - those not in our CCT - $startTime.xlsx"
$outputPath = "$outputFolder\$reportName"
# Applicability Array
$applicArray = @()
$applicList = new-Object System.Collections.Generic.List[string]

#CCT Information
$cctBasePath = "F:\KC46 Staging\production\Manuals\KC46\S1000D\SDLLIVE"
$cctDM = gci -Path $cctBasePath -Filter DMC-1KC46-A-00-00-0000-04A0Z-00QA-D*.xml
$cccFullFilePath = $cctDM[0].FullName
$cctXML = New-Object System.Xml.XmlDocument
$cctXML.Load($cccFullFilePath)
$conditions = $cctXML.SelectNodes("//cond")

#PTC Information
$pctBasePath = "F:\KC46 Staging\production\Manuals\KC46\S1000D\SDLLIVE"
$pctDM = gci -Path $pctBasePath -Filter DMC-1KC46-A-00-00-0000-03A0Z-00PA*.xml
$pctFullFilePath = $pctDM[0].FullName

$exceptionalConditions = @()
$condElementTemplate = [xml] @"
<assign applicPropertyValue="##PRE_POST_WITH_WITHOUT##" applicPropertyType="condition" applicPropertyIdent="##CONDITION##"/> 
"@
$dirty = $false
function Get-ConditionID
{
    Param([string] $condition)  
    #$c = $cctXML.SelectNodes("/dmodule/content/condCrossRefTable/condList/cond[@id='$condition']")
    $c = $cctXML.SelectNodes("//cond[name='$condition']")
    if($c -ne $null )
    { return $true }
    else
    {  return $false }
}

function Get-Product
{
    Param([string] $effectivity)  
    return $pctXML.SelectNodes("/dmodule/content/productCrossRefTable/product[./assign/@applicPropertyValue='$effectivity']")
}

function Clear-ProductCrossRefTable
{
    Param([string] $pathToPCT)
    $pctTmp = New-Object System.Xml.XmlDocument
    $pctTmp.Load($pctFullFilePath)
    $products = $pctTmp.SelectNodes("//product")

    foreach ($product in $products)
    {
        $nodes2RemoveColl = $product.SelectNodes("assign[@applicPropertyType='condition']")
        foreach ($node in $nodes2RemoveColl)
        {
            $null = $product.RemoveChild($node)
        }

        $ce = $condElementTemplate.InnerXml.Replace("##CONDITION##", "WARPS INSTALLED").Replace("##PRE_POST_WITH_WITHOUT##","without")
        $product.InnerXml = $product.InnerXml + $ce
    }
    Save-PrettyXML -FName $pathToPCT -xmlDoc  $pctTmp
}
function Get-LastTailNumberInPCT
{
    $allProducts = $pctXML.SelectNodes("//product")
    $maxNum = ""
    foreach ($Product in $allProducts)
    {
        $id =  $Product.id.ToString().SubString($Product.id.Length-3,3)
        $isNum = 0
        $x = [int32]::TryParse($id , [ref] $isNum )
        if($isNum -gt $maxNum)
        {
            $maxNum = $isNum
        }
    }
    return $maxNum.ToString().PadLeft(3,"0")
}


$maxTailNum = Get-LastTailNumberInPCT

$parserDMC   = New-object -TypeName S1000D.DataModule_401

Clear-ProductCrossRefTable -pathToPCT $pctFullFilePath

$pctXML = New-Object System.Xml.XmlDocument
$pctXML.Load($pctFullFilePath)
$products = $pctXML.SelectNodes("//product")

# Path to 
$dmFiles = gci -Path $ManualsBasePath -Recurse -Filter DMC*.XML
#$dmFiles = gci -Path $ManualsBasePath -Recurse -Filter DMC*20-55-4400-23A0A*.XML



foreach ($dmFile in $dmFiles)
{
    $dmFile.FullName

    $parserDMC.ParseDM($dmFile.FullName)
    #$applics = $parserDMC.Dmodule.dmodule.SelectNodes("//applic")
    $applics = $parserDMC.IdentAndStatusSection.SelectNodes("//applic")
    if($applics.ChildNodes.Count -gt 0)
    {
        foreach ($applic in $applics)
        {
            if($applic.displayText.InnerXml.Length -gt 0)
            {
                $applicParts = $applic.displayText.InnerText.Split(";")
                foreach ($applicPart in $applicParts)
                {
                    if($applicPart.Contains(" WITH ") -or ($applicPart.Contains(" WITHOUT ")) -or ($applicPart.Contains(" PRE ")) -or ($applicPart.Contains(" POST ")))
                    {
                        if($applicPart.Contains(" WITH "))
                        {
                            $condElement = [xml] $condElementTemplate.InnerXml.Replace("##PRE_POST_WITH_WITHOUT##", "without")
                            $INX = $applicPart.IndexOf("WITH")
                        }
                        elseif($applicPart.Contains(" WITHOUT "))
                        {
                            $condElement =  [xml] $condElementTemplate.InnerXml.Replace("##PRE_POST_WITH_WITHOUT##", "without")
                            $INX = $applicPart.IndexOf("WITHOUT")
                        }
                        elseif($applicPart.Contains(" PRE "))
                        { 
                            $condElement =  [xml] $condElementTemplate.InnerXml.Replace("without", "pre")
                            $INX = $applicPart.IndexOf("PRE")
                        }
                        elseif($applicPart.Contains(" POST "))
                        {
                            $condElement =  [xml] $condElementTemplate.InnerXml.Replace("without", "pre")
                            $INX = $applicPart.IndexOf("POST")
                        }

                        $condition = [string] ($applicPart.Substring($INX + 5,$applicPart.Length - ($INX + 5))).Trim()
                        if($condition-eq "ELEVATOR FEEL COMPUTER WITH FOUR INPUTS")
                        {
                            "stop"
                        }

                        $isConditionInCCT = Get-ConditionID -condition $condition
                        if($isConditionInCCT -eq $false -and (! $condition.Contains("WARPS")))
                        {
                            $parserDMC.ParseDM($dmFile.FullName)
                            $dmcParts = $dmFile.Name.Split("-")
                            $book = $parserDMC.AssociatedBook
                            $CH = $dmcParts[3]
                            $SE = $dmcParts[4]
                            $SU = $dmcParts[5]
                            $SName = $dmFile.Name
                            $TName = $parserDMC.TechName
                            $IName = $parserDMC.InfoName
                            $emodBuild = $parserDMC.EmodBuild
                            $exceptionalConditions += New-Object -TypeName PSObject -Property @{
                                Book = $book;
                                DMC = $SName;
                                Condition = $applicPart;
                                CH = $CH;                                      
                                SE = $SE;
                                SU = $SU;
                                TechName    = $TName;
                                InfoName    = $IName;
                                EMODBuild = $emodBuild;
                            } | Select Book,DMC,Condition,CH,SE,SU,TechName,InfoName,EMODBuild
                        }
                        if($isConditionInCCT)
                        {
                            $range = $false
                            $tails = $applicPart.Substring(0, $INX-1)
                            if ($tails.Contains(","))
                            {
                                $tailParts = $tails.Replace(" ", "").Split(",")
                            }
                            elseif($tails.Contains("-"))
                            {
                                $range = $true
                                $tailParts = $tails.Replace(" ", "").Split("-")    
                            }
                            

                            $conditionToStore = $condElement.InnerXml.Replace("##CONDITION##", $condition)
                            if(($tailParts.Count -gt 0) -and  ($range -eq $true))
                            {   
                                $low = [int] $tailParts[0]

                                $hi =   [int] $tailParts[$tailParts.Count -1]  
                                if($hi -eq 999)
                                {
                                    $hi = $maxTailNum
                                } 
           
                                for ($i = $low; $i -le $hi; $i++)
                                { 
                                    $tail = $i.ToString().PadLeft(3,"0")
                                    
                                    $pctTailRecord = Get-Product -effectivity $tail
                                    if(! ($pctTailRecord.InnerXml.Contains($conditionToStore)))
                                    {
                                        $pctTailRecord.InnerXml = $pctTailRecord.InnerXml + $conditionToStore
                                        $dirty = $true
                                    }                                    
                                }
                            }
                            elseif( ($tailParts.Count -gt 0) -and  ($range -eq $false))  
                            {                              
                                foreach ($tail in $tailParts)
                                {
                                    $isNum = 0
                                    $x = [int32]::TryParse($tail , [ref] $isNum )
                                    if($isNum -gt 0)
                                    {
                                        $pctTailRecord = Get-Product -effectivity $tail
                                        if(! ($pctTailRecord.InnerXml.Contains($conditionToStore)))
                                        {
                                            $pctTailRecord.InnerXml = $pctTailRecord.InnerXml + $conditionToStore
                                            $dirty = $true
                                        }
                                    }
                                }
                            }
                            else
                            {
                                foreach ($product in $products)
                                { 
                                    if(! ($product.InnerXml.Contains($conditionToStore)))
                                    {
                                        $product.InnerXml = $product.InnerXml + $conditionToStore
                                        $dirty = $true
                                    }
                                }                    
                            }                      
                       }
                    }
                }
            } 
     
        }
    }

}
if($dirty)
{
    Save-PrettyXML -FName $pctFullFilePath -xmlDoc $pctXML
}  
# Select DMC,CH,SE,SU,TechName,InfoName 
$prop1 = @{Expression='Book'; Ascending=$true }
$prop2 = @{Expression='CH'; Ascending=$true }
$prop3 = @{Expression='SE'; Ascending=$true }
$prop4 = @{Expression='SU'; Ascending=$true }
if(Test-Path -Path $outputFolder)
{
    Remove-Item -Path $outputFolder -Force -Verbose -ErrorAction SilentlyContinue -Recurse
    md $outputFolder
}
$exceptionalConditions.GetEnumerator() | Sort-Object -Property $prop1, $prop2 , $prop3, $prop4 | Export-XLSX -Path $outputPath -Header Book,DMC,Condition,CH,SE,SU,TechName,InfoName,EMODBuild -WorksheetName "Exceptions"

# ********************************************************************
$ed = Get-Date
$x = $ed.Subtract($sd)					
"`r`nTotal Time to complete:`t" + $x.Hours + ":" + $x.Minutes + ":" +  $x.Seconds  + ":" +  $x.Milliseconds
"Process completed"					
					