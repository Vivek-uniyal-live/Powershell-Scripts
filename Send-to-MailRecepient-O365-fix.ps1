New-item -path "HKLM:\SOFTWARE\Microsoft\ClickToRun" -Name OverRide -force

New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ClickToRun\OverRide" -Name "AllowJitvInAppvVirtualizedProcess" -value 1 -PropertyType "DWord" -force
