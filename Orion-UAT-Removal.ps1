write-host "Cleaning instance of Stellar UAT2"
Remove-Item "C:\Program Files (x86)\Data Action\UAT\Orion\orion.UAT2.exe" -Force
Remove-Item "C:\Program Files (x86)\Data Action\UAT\Orion\orion.UAT2.exe.config" -Force
remove-item "C:\Users\Public\Desktop\Orion (UAT2 V1.0.19.58).lnk" -Force
Remove-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Data Action\UAT\Orion (1.0.19.58)\Orion (UAT2 V1.0.19.58).lnk" -Force
write-host  "Completed Cleanup successfully"