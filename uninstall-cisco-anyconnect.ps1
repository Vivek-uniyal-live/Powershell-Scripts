<#
.SYNOPSIS


.DESCRIPTION
  

.PARAMETER
	
.INPUTS
  None

.OUTPUTS
  Log file written to %TEMP%, called <ScriptName>_<ReverseDate>.log

.NOTES
  Version:        2.0
  Author:         Vivek Uniyal
  Creation Date:  30/04/2021


.EXAMPLE
  
#>

#Log Stuff
$MyDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$MyScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$LogFile = "$($env:Temp)\$MyScriptName`_$((Get-Date).ToString("yyyyMMdd")).log"

Function Main{

    WriteLog -Text "---Start CiscoAnyconnect UnInstall---"

    #Check for apps that we know break the install and uninstall them
    UninstallApp -DisplayName "Cisco AnyConnect Secure Mobility Client"
    UninstallApp -DisplayName "Cisco AnyConnect Start Before Login Module"
    UninstallApp -DisplayName "Cisco AnyConnect Posture Module"
    UninstallApp -DisplayName "Cisco AnyConnect Diagnostics and Reporting Tool"

}

Function UninstallApp{
    Param(
    [String]$DisplayName
    )

    $UninstallPath1 = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    $UninstallPath2 = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' 

    WriteLog -Text "Looking for installations of $DisplayName"
    $UniqueIDs = Get-ChildItem -Path $UninstallPath1,$UninstallPath2 | Get-ItemProperty | ?{$_.DisplayName -match $DisplayName} | Select PSChildName

    ForEach($ID in $UniqueIDs)
        {
        WriteLog -Text "Uninstalling $($ID.PSChildName)"
        $MSIArguments = @(
               "/x"
               ('"{0}"' -f $ID.PSChildName)
               "/qn"
               "/norestart"
            )
        WriteLog -Text "Launching Uninstall using arguments: $MSIArguments"
        $Process = Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow -PassThru
        WriteLog -Text "Process completed with exit code $($Process.ExitCode)"
        }
    remove-item "C:\ProgramData\Cisco\Cisco AnyConnect Secure Mobility Client" -recurse -force	
    WriteLog -Text "Finished check/uninstall for $DisplayName"

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

Main
