$cgmDest="C:\KC46 Staging\BetaTR\Manuals\FIGURES"
Remove-Item -Path $cgmDest -Recurse -Force

# Process all the figures

$batfile = [diagnostics.process]::Start("C:\Program Files\XyEnterprise\LiveContent\KC46-ICNBranding.bat")
$batfile.WaitForExit()