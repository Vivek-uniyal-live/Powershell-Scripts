# Script to fix Signature issues on users Outlook.



$LogFile = "C:\Windows\Temp\signature-fix_$((Get-Date).ToString("yyyyMMdd")).log"

Function Main{
$path = "C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures\*"

WriteLog "Testing the local signature path on users device"

if((test-path $path) -eq $true)
        {
writelog "Copy the Signature files to c:\windows\temp\signature-backup location on users device "
copy-item $path "c:\windows\temp\signature-backup" -Force 
remove-item $path -force -recurse

writelog "Performing GPUPDATE in force mode "
invoke-command -scriptblock{gpupdate /force}
writelog "Running Sig.exe command on the device"
\\accu.local\netlogon\EmailSign\V6\sign.exe /database=MailSigs /type=1 /server=DBMAILSIGSPRD\SQL2017 

invoke-command -scriptblock{gpupdate /force /boot}

        }
        else
        {
writelog "Unable to find signature files on users device, Hence exiting the script"
        }


}
Function WriteLog{
    Param(
    [String]$Text,
    [ValidateSet('info','error','success', 'warning')][System.String]$Type="info"
    )
    $Text = "$(Get-Date -Format 'G') $Text" 
    Switch($Type)
    {
        "info"{$Text | Add-Content -path $Logfile -PassThru | Write-Host -ForegroundColor White}
        "error"{$Text | Add-Content -path $Logfile -PassThru | Write-Host -ForegroundColor Red}
        "success"{$Text | Add-Content -path $Logfile -PassThru | Write-Host -ForegroundColor Green}
        "warning"{$Text | Add-Content -path $Logfile -PassThru | Write-Host -ForegroundColor Yellow}
    }
}

WriteLog -Text "----Starting Fixing Signature Issue----"
Main
WriteLog -Text "----Finished Applying the fix----"