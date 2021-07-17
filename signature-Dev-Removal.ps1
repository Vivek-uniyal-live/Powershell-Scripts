$path1 = "C:\Users\Public\desktop\Signature - DEV.lnk"
$path2 = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\J Walk Windows Client\Signature - DEV.lnk"

if(test-path $path1)
{
remove-item -Path "C:\Users\Public\desktop\Signature - DEV.lnk" -Force -Recurse
Write-Output "Removed signature Dev shortcut link on the computer" | Out-File "c:\windows\temp\sigdev.txt"
}
else
{
Write-Output "unable to find signature Dev shortcut link on the computer" | Out-File "c:\windows\temp\sigdev.txt"
}

if(Test-Path $path2)
{
Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\J Walk Windows Client\Signature - DEV.lnk" -Force -Recurse
Write-Output "Removed signature Dev shortcut link on the computer" | Out-File  -Append "c:\windows\temp\sigdev.txt"
}
else
{
Write-Output "unable to find signature Dev shortcut link on the computer" | Out-File -Append "c:\windows\temp\sigdev.txt"
}