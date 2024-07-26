CLS
"This is DDNGenerator, Version 1. Coded by Peter Schilling, July 2015."

#Prompt user for source data until a valid path is entered
$path = read-host "Provide valid source data folder path"
while (!(test-path $path))
    {
    $path = read-host "Invalid path; provide valid source data folder path"
    }
#Check for trailing backslash and append if absent
if (!($path[$path.length-1] -eq "\"))
    {
    $path = ($path+"\")
    }

#Get sender NCAGE code
$senderNCAGEdefault = "81205"
$senderNCAGE = read-host "Provide sender NCAGE or press enter for 81205"
$senderNCAGE = ($senderNCAGEdefault,$senderNCAGE)[[bool]$senderNCAGE]

#Get reciever NCAGE code
$receiverNCAGEdefault = "80049"
$receiverNCAGE = read-host "Provide reciever NCAGE or press enter for 80049"
$receiverNCAGE = ($receiverNCAGEdefault,$receiverNCAGE)[[bool]$receiverNCAGE]

#Prompt user for information to determine a data dispach note number (DDNN)
$DDpath = read-host "Provide valid folder path for previous dispach or press enter to manually input a data dispach note number (DDNN)"
if ($DDpath)
    {
    #Make sure a valid path is entered
    while (!(test-path $DDpath))
        {
        $DDpath = read-host "Invalid path; provide valid folder path for previous dispach"
        }
        #Check for trailing backslash and append if absent
    if (!($DDpath[$DDpath.length-1] -eq "\"))
        {
        $DDpath = ($DDpath+"\")
        }
    #Gather all previous DDNs
    $DDNlist = get-childitem $DDpath -name -include DDN*xml
    #Check for existence of DDNs in path
    while (!($DDNlist))
        {
        $DDpath = read-host "No DDNs found; submit a new path or press enter to manually input a DDNN"
        }
    #Initialize a data dispach note number (DDNN) for comparison
    $DDNNtemp = 1
    #Find the largest previous DDNN
    foreach ($DDNfile in $DDNlist)
        {
        #Load the DDN
        $DDNtemp = [xml](get-content ($DDpath+$DDNfile))

        #Get DDNN from dmCode
        $seqNumber = [int]$DDNtemp.ddn.identAndStatusSection.ddnAddress.ddnIdent.ddnCode.seqNumber
        #Compare DDNNs to find largest
        if ($seqNumber -gt $DDNNtemp) {$DDNNtemp = $seqNumber}
        }      
    #Define data dispach note number
    $DDNN = ($DDNNtemp+1).ToString().PadLeft(5,'0')
    }
else
    {
    #Initialize input check variable
    $check = $false
    #Check for integer input
    do
        {
        try
            {
            [int]$DDNNtemp = read-host "Enter integer DDNN of previous dispach"
            $check = $true
            }
        catch
            {
            "Not an integer."
            }
        }
    until ($check)
    #Define final DDNN
    $DDNN = ($DDNNtemp+1).ToString().PadLeft(5,'0')
    }

#Begin progress bar
write-progress -activity "Building DDN"

#Create list of all non-DDN and non-DML files
$MClist = get-childitem $path -name -exclude DDN*,DML* | sort-object

#Create parallel workflow for building DMC list
workflow build-list()
    {
    #Take input of *MC list and filepath
    param ($MClistin, [string] $pathin)
    #Initialize a list of filenames
    $filenamesXML = @()
    #Initialize a counter for DMCs
    $DMCcount = 0
    #Initialize a counter for PMCs
    $PMCcount = 0
    #Initialize a counter for other files
    $othercount = 0
   
    #Build a list of filenames with proper xml tags
    foreach -parallel ($filename in $MClistin)  
        {
       $inline = inlinescript
            {
            #Bring external variables into inlinescript
            $filenamesXMLinline = $using:filenamesXML

            $DMCcountinline = $using:DMCcount
            $PMCcountinline = $using:PMCcount
            $othercountinline = $using:othercount
            #Differentiate among PMC xml, DMC xml, and any other files not previously filtered for
            switch -wildcard ($filename)
                {
                "DMC*xml" 
                    {
                    #Load xml file
                    $DMC = [xml](get-content ($pathin+$filename))

                    #Get properties from dmCode
                    $dmCode = $DMC.dmodule.identAndStatusSection.dmAddress.dmIdent.dmCode

                    #Assign individual properties to objects
                    $assyCode = $dmCode.assyCode
                    $disassyCode = $dmCode.disassyCode
                    $disassyCodeVariant = $dmCode.disassyCodeVariant
                    $infoCode = $dmCode.infoCode
                    $infoCodeVariant = $dmCode.infoCodeVariant
                    $itemLocationCode = $dmCode.itemLocationCode
                    $modelIdentCode = $dmCode.modelIdentCode
                    $subSubSystemCode = $dmCode.subSubSystemCode
                    $subSystemCode = $dmCode.subSystemCode
                    $systemCode = $dmCode.systemCode
                    $systemDiffCode = $dmCode.systemDiffCode

                    #Build DMC filename
                    $DMCtag = ("<deliveryListItem><dispatchFileName>DMC"+"-"+$modelIdentCode+"-"+$systemDiffCode+"-"+$systemCode+"-"+$subSystemCode+$subSubSystemCode+"-"+$assyCode+"-"+$disassyCode+$disassyCodeVariant+"-"+$infoCode+$infoCodeVariant+"-"+$itemLocationCode+".xml</dispatchFileName></deliveryListItem>")

                    #Add filename to array
                    $filenamesXMLinline = $DMCtag
                    #Increment DMC counter
                    $DMCcountinline++
                    }
                "PMC*xml"
                    {
                    #Load xml file
                    $PMC = [xml](get-content ($pathin+$filename))

                    #Get properties from pmCode
                    $pmCode = $PMC.pm.identAndStatusSection.pmAddress.pmIdent.pmCode

                    #Assign individual properties to objects
                    $modelIdentCode = $pmCode.modelIdentCode
                    $pmIssuer = $pmCode.pmIssuer
                    $pmNumber = $pmCode.pmNumber
                    $pmVolume = $pmCode.pmVolume

                    #Build PMC filename
                    $PMCtag = ("<deliveryListItem><dispatchFileName>PMC"+"-"+$modelIdentCode+"-"+$pmIssuer+"-"+$pmNumber+"-"+$pmVolume+".xml</dispatchFileName></deliveryListItem>")

                    #Add filename to array
                    $filenamesXMLinline = $PMCtag

                    #Increment PMC counter
                    $PMCcountinline++
                    }
                  default
                    {
                    #Ignore all properties; pass through filename with DDN xml markup
                    $filenamesXMLinline = ("<deliveryListItem><dispatchFileName>"+$filename+"</dispatchFileName></deliveryListItem>")
                    
                    #Increment other counter
                    $othercountinline++
                    }
                }
            #Return inline variables
            return $filenamesXMLinline, $DMCcountinline, $PMCcountinline, $othercountinline, $modelIdentCode
            }
        #Assign returned inline variables to workflow variables
        $workflow:filenamesXML += $inline[1]
        $workflow:DMCcount = $inline[2]
        $workflow:PMCcount = $inline[3]
        $workflow:othercount = $inline[4]
        $modelIdentCodeflow = $inline[5]
        }
    #Return workflow variables
    $filenamesXML
    return $filenamesXML, $DMCcount, $PMCcount, $othercount, $modelIdentCodeflow
    }
#Run and time the list-building workflow
$timer = measure-command { $workflowout = build-list -MClist $MClist -path $path }

#Deconstruct the workflow output
$filenamesXML, $DMCcount, $PMCcount, $othercount, $modelIdentCode = $workflowout
$filenamesXML
#Create DDN text
$DDN = [xml]@"
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE ddn>
<ddn xsi:noNamespaceSchemaLocation="http://www.s1000d.org/S1000D_4-0-1/xml_schema_flat/ddn.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink">
<identAndStatusSection>
<ddnAddress>
<ddnIdent>
<ddnCode modelIdentCode="$modelIdentCode" senderIdent="$senderNCAGE" receiverIdent="$receiverNCAGE" yearOfDataIssue="2015" seqNumber="$DDNN"/>
</ddnIdent>
<ddnAddressItems>
<issueDate year="$((get-date).year)" month="$((get-date).tostring("MM"))" day="$((get-date).tostring("dd"))"/>
<dispatchTo>
<dispatchAddress>
<enterprise>
<enterpriseName>Headquarters, Department of The Air Force</enterpriseName>
</enterprise>
<address>
<city>Seattle</city>
<country>USA</country>
</address>
</dispatchAddress>
</dispatchTo>
<dispatchFrom>
<dispatchAddress>
<enterprise>
<enterpriseName>Boeing Commercial Airplanes</enterpriseName>
</enterprise>
<address>
<city>Seattle</city>
<country>USA</country>
</address>
</dispatchAddress>
</dispatchFrom>
</ddnAddressItems>
</ddnAddress>
<ddnStatus>
<security securityClassification="01" commercialClassification="cc52"/>
<dataRestrictions>
<restrictionInstructions>
<dataDistribution>Distribution Statement "D" - Distribution authorized to the Department of Defense and U.S. DoD contractors only (U.S. DoDcontractors must be qualified with assigned duties in direct support of the KC-46 Tanker Modernization Directorate IAW AFI 61-204AFGM2 and comply with all applicable contractual license requirements); (administrative or operational use) (Date ofDetermination equals PAGE DATE below). Other requests for this document shall be referred to KC-46 Tanker ModernizationDirectorate, 2590 Loop Road West, WPAFB, OH 45433-7142.</dataDistribution>
<exportControl>
<exportRegistrationStmt>
<simplePara>WARNING - This document contains technical data whose export is restricted by the Arms Export Control Act (Title 22, U.S.C., Sec2751 et seq) or the Export Administration Act of 1979, as amended, Title 50, U.S.C., App. 2401 et seq. Violations of these export lawsare subject to severe criminal penalties. Disseminate in accordance with provisions of DoD Directive 5230.25.</simplePara>
</exportRegistrationStmt>
</exportControl>
<dataHandling></dataHandling>
<dataDestruction></dataDestruction>
</restrictionInstructions>
<restrictionInfo>
<copyright>
<copyrightPara>BOEING PROPRIETARY, CONFIDENTIAL, AND/OR TRADE SECRET</copyrightPara>
<copyrightPara>Copyright &#169; $((get-date).year) The Boeing Company</copyrightPara>
<copyrightPara>Unpublished Work - All Rights Reserved</copyrightPara>
<copyrightPara>Boeing claims copyright in each page of this document only to the extent that the page contains copyrightable subject matter.</copyrightPara>
<copyrightPara>Boeing also claims copyright in this document as a compilation and/or collective work.</copyrightPara>
<copyrightPara>Boeing, the Boeing signature, the Boeing symbol, 707, 717, 727, 737, 747, 757, 767, 777, 787, Dreamliner, BBJ, DC-8, DC-9, DC-10, KC-10, KC-46, KDC-10, MD-10, MD-11, MD-80, MD-88, MD-90, P-8, Poseidon and the Boeing livery are all trademarks owned by The Boeing Company; and no trademark license is granted in connection with this document unless provided in writing by Boeing.</copyrightPara>
</copyright>
<dataConds>GOVERNMENT PURPOSE RIGHTS Contract No.:  FA8625-11-C-6600 Contractor:  The Boeing Company Contractor Address:  PO BOX 3707, Seattle WA 98124 Expiration Date:  24 February 2016. The Government's right to use, modify, reproduce, release, perform, display, or disclose these technical data or computer software/documentation are restricted by paragraph (b)(2) of the Rights in Technical Data - Noncommercial Items or DFARS 252.227-7014(b)(2) of Non-Commercial Rights in Computer Software and Computer Software Documentation clause contained in the above identified contract. No restrictions apply after the expiration date shown above. Any reproduction of technical data or computer software or portions thereof marked with this legend must also reproduce the marking.</dataConds>
</restrictionInfo>
</dataRestrictions>
<authorization/>
</ddnStatus>
</identAndStatusSection>
<ddnContent>
<deliveryList>
$filenamesXML
</deliveryList>
</ddnContent>
</ddn>
"@

#Build DDN filename
$DDNfilename = ("DDN-BDSKC-"+$senderNCAGE+"-"+"$receiverNCAGE"+"-"+(get-date).year+"-"+$DDNN+".xml")

#Save DDN to XML
$DDN.save("$path$DDNfilename")

#Display ending message
"$DDNfilename was created in $path with $DMCcount DMC(s), $PMCcount PMC(s), and $othercount other file(s) in $($timer.totalseconds) seconds."


<#
FROM: http://stackoverflow.com/questions/13416651/passing-updating-hashtables-and-arrays-by-reference-in-powershell


function UpdateArray {  
    param( [ref]$ArrayNameWithinFunction )  
    $ArrayNameWithinFunction.Value += 'xyzzy'  
}  
$MyArray = @('a', 'b', 'c')  
UpdateArray ([ref]$MyArray)  
$MyArray
#>