
# *****************************************************************************************************
WorkFlow MirrorFolders
{
    $XyEnterprise = "Program Files\XyEnterprise\LiveContent"
    parallel
    {
        # 5.6 Manuals on the Primary machine

        robocopy /e /mir /NP  "C:\KC46 Staging\Production\Manuals"  "\\a5547879\C`$\KC46 Staging\Production\Manuals"
        
        <# 4.0.2 Manuals locally #>
        robocopy /e /mir "C:\KC46" "\\a5778954\C$\KC46"
        robocopy /e /mir "C:\KC46 Staging\Dev"  "\\a5778954\C$\KC46 Staging\Dev"
        robocopy /e /mir "C:\KC46 Staging\Production"  "\\a5778954\C$\KC46 Staging\Production"
        robocopy /e /mir "C:\KC46 Staging\Scripts"  "\\a5778954\C$\KC46 Staging\Scripts"
        robocopy /e /mir "C:\KC46 Staging\Resources"  "\\a5778954\C$\KC46 Staging\Resources"
        robocopy /e /mir "D:\AllTankerIllustrations" "\\a5778954\C$\AllTankerIllustrations"
        robocopy /e /mir "C:\LiveContentData" "\\a5778954\C$\LiveContentData"
        robocopy /e /mir  "C:\$XyEnterprise" "\\a5778954\C$\$XyEnterprise"
        robocopy /e /mir  "D:\Shared" "\\A5778954\d$\Shared"
        
    }       
}

cls
$PSVersionTable
$env:COMPUTERNAME
if(($env:COMPUTERNAME -eq "A5723772") -and $env:COMPUTERNAME -ne "A5778954")
{
    MirrorFolders
}