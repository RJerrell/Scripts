#region Monitoring and Distribution
Set-Variable -Name "pathSuffix" -Value "_11_2018_Rel11.1" -Description "IPR NAME"
Set-Variable -Name "outputFolderBaseName" -Value "D:\Shared\BCDR\IPR\100 Percent IPR - Nov 2018\KC46_100_IPR_" -Description "Top level folder name for the output of our BCDR processing"
Set-Variable -Name "inputPathBaseFolder" -Value "D:\Shared\IDE cd sets\Releases\2018-11-19-11-48-30 Release 11.1\CSDB\DVD" -Description "Source data location of the data to put into BCDR"
#endregion
#region BASE GLOBALS
$global:sd = Get-Date
$global:startTime = $global:sd.Year.ToString()
$global:startTime += "-" + $global:sd.Month.ToString().PadLeft(2,"0")
$global:startTime += "-" + $global:sd.Day.ToString().PadLeft(2,"0")
$global:startTime += "-" + $global:sd.Hour.ToString().PadLeft(2,"0")
$global:startTime += "-" + $global:sd.Minute.ToString().PadLeft(2,"0")
$global:startTime += "-" + $global:sd.Second.ToString().PadLeft(2,"0")
$env:PSModulePath = "C:\Program Files (x86)\PowerShell Community Extensions\Pscx3\;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;c:\KC46 Staging\Scripts\Modules;"
#region Common Items
$a = (Get-Host).PrivateData
$a.WarningBackgroundColor = "Yellow"
$a.WarningForegroundColor = "Black"
$rn = "<br/>"
$a.ErrorBackgroundColor = "red"
$a.ErrorForegroundColor = "White"
#endregion