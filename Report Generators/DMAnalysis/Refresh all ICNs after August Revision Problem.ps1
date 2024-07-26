WorkFlow RefreshICNs
{
    parallel
    {
        Copy-Item -verbose -Path "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DDN-1KC46-AAAZZ-81205-2015-00003-ICN Refresh\ICN-81205-KA*"  -Destination "C:\KC46 Staging\Production\Manuals\AMM\ILLUSTRATIONS\ILLUSTRATIONS"
        Copy-Item -verbose -Path "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DDN-1KC46-AAAZZ-81205-2015-00003-ICN Refresh\ICN-81205-KE*"  -Destination "C:\KC46 Staging\Production\Manuals\ARD\ILLUSTRATIONS\ILLUSTRATIONS"
        Copy-Item -verbose -Path "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DDN-1KC46-AAAZZ-81205-2015-00003-ICN Refresh\ICN-81205-KF*"  -Destination "C:\KC46 Staging\Production\Manuals\FIM\ILLUSTRATIONS\ILLUSTRATIONS"
        Copy-Item -verbose -Path "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DDN-1KC46-AAAZZ-81205-2015-00003-ICN Refresh\ICN-81205-KR*"  -Destination "C:\KC46 Staging\Production\Manuals\SSM\ILLUSTRATIONS\ILLUSTRATIONS"
        Copy-Item -verbose -Path "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DDN-1KC46-AAAZZ-81205-2015-00003-ICN Refresh\ICN-81205-KS*"  -Destination "C:\KC46 Staging\Production\Manuals\SRM\ILLUSTRATIONS\ILLUSTRATIONS"
        Copy-Item -verbose -Path "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DDN-1KC46-AAAZZ-81205-2015-00003-ICN Refresh\ICN-81205-KV*" -Destination "C:\KC46 Staging\Production\Manuals\NDT\ILLUSTRATIONS\ILLUSTRATIONS"
        Copy-Item -verbose -Path "C:\KC46 Staging\Production\Archives\Source\UnpackHere\DDN-1KC46-AAAZZ-81205-2015-00003-ICN Refresh\ICN-81205-KW*"  -Destination "C:\KC46 Staging\Production\Manuals\WDM\ILLUSTRATIONS\ILLUSTRATIONS"
    }
}

RefreshICNs