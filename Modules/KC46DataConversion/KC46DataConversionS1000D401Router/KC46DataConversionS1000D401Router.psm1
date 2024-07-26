function ConvertTo-S1000D4_0_1
{
Param([Xml.XmlElement[]] $EL2Convert, [string] $ELType )
    $EL2ConvertParentNode = $EL2Convert.ParentNode.Name
    <# SGML TAGS PER THE WDM/SWPM Style Guide
    CHANGE RECORD
    INTRODUCTION
    %deleted - Deleted
    %ein - Equipment Identification Number
    %revatt - Revision Attributes
    %wcn - Warnings, Cautions and Notes
    %text - Text Structures of an Element
    %yesorno - Yes or No Choice
    aclist - Aircraft List
    acn - Aircraft Registration Number
    actwire - Active Wire
    alprtnbr - Airline Part Number (Airline Stock Number)
    altentry - Alternate Entry
    atanbr - ATA Number
    benbr - Basic Engineering Number
    bI - Buttock Line Number
    bunddesc - Bundle Description
    bundnbr - Bundle Number
    bundpnr - Bundle Part Number
    cac - Customer Alternate Code
    X - caution - Caution
    cec - Customer Effectivity Code
    ch2oref - Chapter 20 Reference
    X - chapter - Chapter
    X - chgdesc - Change Description
    cocdata - Customer Originated Change Data
    coceff - Customer-originated Change Effectivity
    coclist - Customer Originated Change List
    cocnbr - Customer Originated Change Number
    cocrev - Customer Originated Change Revision Number
    cocst/cocend - Customer-Originated Change Start/Customer-Originated Change End
    X - colspec - Column Specification
    condid - Conduit Identification
    condlist - Conduit List
    condrow - Conduit Row
    connect - Connector
    cpyrght - Copyright
    cus - Customer Code
    cusaid - Customer Assigned Airplane Identification
    cusrevrec - Customer Revision Record
    custeff - Customer Effectivity
    def - Definition
    deleted - Deleted
    dia - Dimensional Data
    diagref - Diagram Reference
    dwg - Drawing
    ect - Effectivity Code Termination
    effdata - Effectivity Cross Reference Data
    effect - Effectivity
    effrc - Effectivity Range Codes
    effxref - Effectivity Cross (X) Reference
    em - Equipment Identification Number
    einmfr - Equipment Identification Number Manufacturer
    entry - Entry in a Table
    eonbr - Engineering Order Number
    eqiplist - Equipment List
    eqivpnr - Equivalent Part Number
    eqrow - Equipment Row
    esnbr - Engine Set Number
    etype - Electrical Type
    exprtcl - Export Control
    extwlist - Extended Wire List
    extwrow - Extended Wire Row
    fam - Family
    feedthru - Feed Through
    from - Wire From
    fslgnbr - Fuselage Number
    ftnote - Footnote
    fullstmt - Full Statement
    gdesc - Graphic Description
    geninfo - General Information
    glosdata - Glossary Data
    glossary - Glossary
    X - graphic - Graphic
    gridnbr - Grid Number
    grphcref - Graphic Reference
    grsymbol - Inline Graphic
    hazmt - Hazardous Material
    hazmtlst - Hazardous Material List
    hdiagnbr - Home Diagram Number
    holder - Holder
    hotlink - Hotlink
    X - hwinfo - Hardware (Part Number) Information
    hwname - Hardware (Part Number) Name
    icdate - Incorporation Date
    ics - Incorporation Status
    increv - Incremental Revision
    intro - Introduction
    isempty - Empty Element
    legalntc - Legal Notice
    length - Length
    linenbr - Line Number
    X - listX - List of Level X (1 - 7)
    location - Location
    X - IXitem-Itemofa List of LeveIX
    mad - Manufacturer Address
    matewith - Mate With
    matewmfr - Mate With Vendor
    matlinfo - Material Information
    matlname - Material Name
    matlpnr - Material Part Number
    maxpos - Maximum Position
    mdnbr - Modification Number
    mfmatr - Manual Front Matter
    mfr - Manufacturer Code
    mfrevrec - Manufacturers Revision Record
    X - mfrname - Manufacturer Name
    modcode - Modification Code
    modtype - Model Type
    msnbr - Manufacturing Serial Number
    nactwire - Non-Active Wire
    X - note - Note
    numlist - Numbered List
    numlitem - Numbered List Item
    operator - Operator
    optcode - Option Code
    pan - Panel
    pandesc - Panel Description
    paptrnbr - Paper Temporary Number
    X - para - Paragraph Block
    partdesc - Part Description
    partstmt - Partial Statement
    pnr - Part Number
    pnrnha - Next Higher Assembly Part Number
    position - Position
    proptary - Proprietary Statement
    pretopic - Preliminary Topic
    refblock - Reference Block
    refext - Reference (External)
    X - refint - Reference (Internal)
    refmedia - Reference (Media)
    regulation - Regulation
    regulatory - Regulatory
    revst/revend - Revision Start/Revision End
    row - Row
    sbdata - Service Bulletin Data
    sbeff - Service Bulletin Effectivity
    sblist - Service Bulletin List
    sbnbr - Service Bulletin Number
    sbnrmfr - Service Bulletin Number and Manufacturer
    sbrev - Service Bulletin Revision
    sc - Started/Comoleted
    section - Section
    sensep - Sensitivity/Separation
    sheet - Sheet
    shunt - Shunt Code
    spanspec - Span Specification
    sparepin - Spare Pin
    spchap - Standard Practices Chapter
    spsect - Standard Practices Section
    spsubj - Standard Practices Subject
    stnnbr - Station Number
    sub - Subscript
    subject - Subject
    super - Superscript
    X - table - Table
    tbody - Table Body
    term - Term
    termcode - Termination Code
    terminfo - Termination Information
    termname - Termination Name
    termnbr - Termination Number
    termpnr - Termination Part Number
    tfoot - Table Footer
    tgroup - Table Group
    thead - Table Header
    X - title - Title
    to - To
    X - toolinfo - Tool Information
    toolname - Tool Name
    toolpnr - Tool Part Number
    tqa - Total Quantity per AircraftJEngine
    tr - Temporary Revision
    transltr - Transmittal Letter
    trdata - Temporary Revision Data
    trfmatr - Temporary Revision Front Matter
    trinfo - Temporary Revision Information
    trlist - Temporary Revision List
    trloc - Temporary Revision Location
    trnbr - Temporary Revision Number
    trstatus - Temporary Revision Status
    trxref - Temporary Revision Cross Reference
    ttl page - Title Page
    tctgrphc - Text Graphic
    txtline - Text Line
    X - unlist - Unnumbered List
    X - unlitem - Unnumbered List Item
    venaddr - Vendor Address
    venbr - Variable Engineering Number
    vendata - Vendor List Data
    vendlist - Vendor Name and Address List
    veninfo - Vendor Information
    venname - Vendor Name
    venrpl - Vendor Replacement
    X - warning - Warning
    wire - Wire
    wireawg - Wire Guage
    wireclr - Wire Color
    wirecode - Wire Code Identification
    wirenbr - Wire Number
    wirerte - Wire Route
    wiretype - Wire Type
    wI - Waterline
    wm - Wiring Diagram Manual
    wmlist - Wiring Manual List
    wnbrmfr - Wire Number Grouping
    year - Year
    zone - Zone Number
#>

    $list_xml_return = ""
    $listItem_xml_text = ""
    # Parent LIST Node

    $key = $EL2Convert.KEY
    $listKey = ""
    $listItem_xml_text = ""
    $EL2ConvertChildNodeText = ""

    # CHILD node of a parent LIST node
    $childNode_text = ""
    switch -Wildcard ($ELType.ToUpper())
    {

     'CAUTION' {           
        $childNode_text_out = Assert-WCN -EL2Convert $EL2Convert  -ELType "CAUTION"
        break      
    }
     'CHGDESC' {
        $childNode_text_out = ""          
        #$childNode_text_out = Process-WCN -EL2Convert $EL2Convert  -ELType "CAUTION"
        break
    }        
    'GRAPHIC' {           
        if($EL2Convert.ParentNode.Name.ToUpper() -ne "PARA" )
        {
            $graphic_xml_text = ConvertFrom-GRAPHIC -EL2Convert $EL2Convert -ELType $ELType -createIRef $true
        }
        else
        {
            $graphic_xml_text = ConvertFrom-GRAPHIC -EL2Convert $EL2Convert -ELType $ELType 
        }        
        $childNode_text_out = $graphic_xml_text
        break    
    }
     'HWINFO'
    {
        "yes"
        break
    } 
    'MATLINFO'
    {
        break
    }
    'TOOLINFO'
    {
        break
    }
    'TERMINFO' {    
        # A random list in S1000D can be of types pf01-99 but commonly only 1-3 are supported.  
        # The default will be pf02 unless otherewise specified using the listBulletType argument
        $childNode_text_out = Assert-RandomList -EL2Convert $EL2Convert -ELType $ELType -listBulletType "pf01"
        break
    }
    'GRPHCREF'{
        $childNode_text_out = ConvertFrom-GRPHCREF -EL2Convert $EL2Convert -ELType $ELType        
        break    }
    'LIST*'
    {
        <# DTD Element Definitions for each of the numeric EL2Convert types
        <!ELEMENT l1item - - ((chgdesc*, title?, regulatory*, (%text;)?,list2?) | %deleted;) 
        <!ELEMENT l2item - - ((chgdesc*, title?, regulatory*, (%text;)?,list3?) | %deleted;) 
        <!ELEMENT l3item - - ((chgdesc*, title?, regulatory*, (%text;)?,list4?) | %deleted;) 
        <!ELEMENT l4item - - (title?, regulatory*, (%text;)?, list5?) > 
        <!ELEMENT l5item - - (title?, regulatory*, (%text;)?, list6?) >
        <!ELEMENT l6item - - (title?, regulatory*, (%text;)?, list7?) >
        <!ELEMENT l7item - - (title?, regulatory*, %text;) >
        #>
        $childNode_text_out = ConvertFrom-LIST -EL2Convert $EL2Convert -ELType $ELType        
        break        
    }
    'MFRNAME' {
        $childNode_text_out = $EL2Convert.InnerText
        break
    }	  
    'NOTE' {
        $childNode_text_out = ConvertFrom-NOTE -EL2Convert $EL2Convert -ELType $ELType        
        break
    }
    'NUMLIST' {
        $childNode_text_out = "<$ELType> * * * * * * * * " + $EL2Convert.Name  + "* * * * * * * * *</$ELType>"  
        break
    }		     
    'PARA' {        
        $childNode_text_out = ConvertFrom-PARA -EL2Convert $EL2Convert -ELType $ELType
        break  
    }
    'REFINT'{
        $childNode_text_out = ConvertFrom-REFINT -EL2Convert $EL2Convert -ELType $ELType
        break          
    }
    'REFEXT'
    {
        # <externalPubRef><externalPubRefIdent><externalPubTitle>TO 00-25-172</externalPubTitle></externalPubRefIdent></externalPubRef>
        $childNode_text_out = '<externalPubRef><externalPubRefIdent><externalPubTitle>' + $EL2Convert.InnerText + '</externalPubTitle></externalPubRefIdent></externalPubRef>'
        break
    }
    {$_ -in 'REVST','REVEND'}{
        # Add code to handle this
        break
    }
    'TABLE' {
        # Workaround for the SDL Viewer limitations on large tables
        #cls
        [int] $maxRows = 1500
        [int] $rowCount =  $EL2Convert.TGROUP.TBODY.ChildNodes.Count

        [int] $iterations = $rowCount / $maxRows
        [int] $mod = $rowCount % $maxRows
        if( $mod -gt 0) 
        {
            $iterations ++
        }
        $startRowNum = 0        
        if($rowCount -lt $maxRows)
        {
            $endRowNum = $rowCount
            $childNode_text_out = ConvertFrom-TABLE -EL2Convert $EL2Convert -startRowNum $startRowNum -endRowNum $endRowNum
            break
        }
        else
        {
            $endRowNum = $maxRows
            for ($i = 1; $i -lt $iterations; $i++)
            {
                [int] $morerows = 0
                $childNode_text_out += ConvertFrom-TABLE -EL2Convert $EL2Convert -startRowNum $startRowNum -endRowNum $endRowNum -morerows ([ref] $morerows)
                #$rowCount
                $startRowNum = $endRowNum + $morerows
                $endRowNum   = $startRowNum + $maxRows
            }
            break
        }
    }
    'TITLE' {        
        $childNode_text_out = ConvertFrom-TITLE -EL2Convert $EL2Convert -ELType $ELType
        break  
    }
    'UNLIST' {
        $childNode_text_out = ConvertFrom-UNLIST -EL2Convert $EL2Convert -ELType $ELType
        break       
    }
    'WARNING' {			
		$childNode_text_out = Assert-WCN -EL2Convert $EL2Convert  -ELType "WARNING"
        break      
    }
    ''
    {}
    default {
            "<$ELType> * * * * * * * * " + $EL2Convert.Name  + "* * * * * * * * *</$ELType>"
        }
    }
    return $childNode_text_out
}