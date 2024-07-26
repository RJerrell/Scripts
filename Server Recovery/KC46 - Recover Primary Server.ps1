
# *****************************************************************************************************
Workflow MirrorFolders
{
    $XyEnterprise = "Program Files\XyEnterprise\LiveContent"

    sequence
    {
        parallel
        {
            robocopy /e /mir /NS /NC /NFL /NDL /NP   "C:\KC46 Staging\Dev"  "\\a5778954\C$\KC46 Staging\Dev"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\KC46 Staging\Production"  "\\a5778954\C$\KC46 Staging\Production"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\KC46 Staging\Scripts"  "\\a5778954\C$\KC46 Staging\Scripts"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\KC46 Staging\Resources"  "\\a5778954\C$\KC46 Staging\Resources"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\LiveContentData" "\\a5778954\C$\LiveContentData"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\$XyEnterprise" "\\a5778954\C$\$XyEnterprise"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\KC46" "\\a5778954\C$\KC46"
        }
        
        parallel
        {
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\KC46 Staging\Dev"  "D:\KC46 Staging\Dev"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\KC46 Staging\Production"  "D:\KC46 Staging\Production"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\KC46 Staging\Scripts"  "D:\KC46 Staging\Scripts"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\KC46 Staging\Resources"  "D:\KC46 Staging\Resources"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\LiveContentData" "D:\LiveContentData"
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\$XyEnterprise" "D:\$XyEnterprise"            
            robocopy /e /mir /NS /NC /NFL /NDL /NP  "C:\KC46" "D:\KC46"
        }
    }
}

cls

MirrorFolders