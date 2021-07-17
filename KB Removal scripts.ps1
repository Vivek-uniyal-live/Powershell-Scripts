$UpdateVersion = "17763.1158"
$SearchUpdates = dism /online /get-packages | findstr "Package_for"
$updates = $SearchUpdates.replace("Package Identity : ", "") | findstr $UpdateVersion
DISM.exe /Online /Remove-Package /PackageName:$updates /quiet /norestart
