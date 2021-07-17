<#
.SYNOPSIS


.DESCRIPTION
  

.PARAMETER
	
.INPUTS
  None

.OUTPUTS
  Log file written to %TEMP%, called <ScriptName>_<ReverseDate>.log

.NOTES
  Version:        1.0
  Author:         Vivek Uniyal
  Creation Date:  26/05/2020

  1.0 - 24/02/2020 - Initial script


.EXAMPLE
  
#>

#Log Stuff
$MyDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$MyScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$LogFile = "$($env:Temp)\$MyScriptName`_$((Get-Date).ToString("yyyyMMdd")).log"

Function Main{

    WriteLog -Text "---Start Zoom UnInstall---"

    #Check for apps that we know break the install and uninstall them
    UninstallApp -DisplayName "zoom"

    start-sleep -Seconds 4
<#
    #Do the install
    $Arguments = @("INSTALL_SILENT=Enable AUTO_UPDATE=Disable WEB_ANALYTICS=Disable EULA=Disable REBOOT=Disable NOSTARTMENU=Enable REMOVEOUTOFDATEJRES=0 SPONSORS=Disable")
    $x86Installer = "$($MyDir)\jre-8u251-windows-i586.exe"
    $x64Installer = "$($MyDir)\jre-8u251-windows-x64.exe"
    
    WriteLog -Text "Installing Java x86"
    WriteLog -Text "Launching: $x86Installer $Arguments"
    $Process = Start-Process $x86Installer -ArgumentList $Arguments -Wait -NoNewWindow -PassThru
    WriteLog -Text "Process completed with exit code $($Process.ExitCode)"

    WriteLog -Text "Installing Java x64"
    WriteLog -Text "Launching: $x64Installer $Arguments"
    $Process = Start-Process $x64Installer -ArgumentList $Arguments -Wait -NoNewWindow -PassThru
    WriteLog -Text "Process completed with exit code $($Process.ExitCode)"

    WriteLog -Text "---Finished Java Install---"
#>
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
