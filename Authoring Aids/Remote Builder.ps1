cls
#Set-Item wsman:\localhost\client\trustedhosts *
#Restart-Service WinRM
# $server = "kc46-lc-sdl"
$server = "a5723772"
# This is a comma separated list of the mil manuals: (ABDR,ACS,LOAP,MOM,SIMR,WUC)
# Note the below command must always start with KC46,add your mil manual code after without spaces. 
$cred = Get-Credential

Invoke-command { powershell.exe -ExecutionPolicy Unrestricted -NoProfile  -Command "& {& 'C:\KC46 Staging\Scripts\KC46DataManagement - RAM Drive Version.ps1' -publicationList KC46,LOAPS}" } -computername $server -Credential $cred

"DONE"