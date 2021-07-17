write-host "Cleaning instance of Stellar UAT2"
Remove-Item "C:\Program Files (x86)\Data Action\UAT\Stellar\Stellar.UAT2.exe" -Force
Remove-Item "C:\Program Files (x86)\Data Action\UAT\Stellar\Stellar.UAT2.exe.config" -Force
remove-item "C:\Users\Public\Desktop\Stellar (UAT2 V1.0.19.58).lnk" -Force
Remove-Item "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Data Action\UAT\Stellar (1.0.19.58)\Stellar (UAT2 V1.0.19.58).lnk" -Force
write-host  "Completed Cleanup successfully"