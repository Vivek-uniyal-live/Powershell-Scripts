$path = "\\mayor\psource`$\Scripts\Monitor4.ps1"
$trigger = New-ScheduledTaskTrigger  -AtStartup 
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $path
Register-ScheduledTask -TaskName "Monitor4_JOB" -Trigger $trigger -Action $Action -RunLevel Highest -Force


