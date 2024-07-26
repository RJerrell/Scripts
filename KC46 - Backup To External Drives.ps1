# *****************************************************************************************************
Function MirrorFolders_Local
{
    $XyEnterprise = "Program Files\XyEnterprise\LiveContent"
    
    {
        robocopy /e /mir "C:\KC46" "D:\KC46"
        robocopy /e /mir "C:\KC46 Staging\Dev"  "D:\KC46 Staging\Dev"
        robocopy /e /mir "C:\KC46 Staging\Scripts"  "D:\KC46 Staging\Scripts"
        robocopy /e /mir "C:\KC46 Staging\Resources"  "D:\KC46 Staging\Resources"
        robocopy /e /mir "C:\KC46 Staging\Production"  "D:\KC46 Staging\Production"
        robocopy /e /mir "C:\LiveContentData" "D:\LiveContentData"
        robocopy /e /mir  "C:\$XyEnterprise" "D:\$XyEnterprise"
    }       
}

cls
$PSVersionTable
MirrorFolders_Local
