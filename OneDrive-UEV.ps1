#Stamp the registry with updated value for UEV - OneDrive
WriteLog -Text "Updating Registry with UEV - OneDrive Value"
New-ItemProperty -Path "HKLM:\Software\Microsoft\UEV\Agent\Configuration" -Name "ApplyExplorerCompatFix" -PropertyType DWORD -Value 1 -Force