<#
.SYNOPSIS
Remval of Zoom files from  C:\Users\Username\AppData\Roaming\Zoom\bin location.

.DESCRIPTION
  

.PARAMETER
	
.INPUTS
  None

.OUTPUTS
  Log file written to %TEMP%, called <ScriptName>_<ReverseDate>.log

.NOTES
  Version:        1.0
  Author:         Vivek Uniyal
  Creation Date:  04/06/2020




.EXAMPLE
  
#>

#Log Stuff
$MyDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$MyScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
$LogFile = "$($env:Temp)\$MyScriptName`_$((Get-Date).ToString("yyyyMMdd")).log"

Function Main{

    WriteLog -Text "---Start Zoom Files remedaiation---"
    #Do the install
    $users = Get-childitem "C:\users"
    $username = $users.name
    WriteLog -Text "Searching Zoom installers"
    
    foreach($user in $username)
    {
    $path = "C:\users\$user\AppData\Roaming\Zoom"
	if(test-path $path)
	{
	WriteLog -Text "Deleting Zoom folder for $user"
	remove-item $path -Force -Recurse
	}
	else
	{
	WriteLog -Text "Did not find zoom files for $user"
	}      
    }
    WriteLog -Text "---Completed Zoom Files remedaiation---"
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
