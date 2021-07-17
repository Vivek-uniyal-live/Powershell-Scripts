# Script: SendEmail-Upgrade-Successful.ps1
# Description: Sends an email the  OSD Task Sequence is successful.
# Execution: PS> PowerShell.exe -ExecutionPolicy ByPass ".\SendEmail-Upgrade-Successful.ps1" -Wait -WindowStyle Hidden
# Author: Vivek Uniyal
 
# Set Connection Variables
    $From = "SCCMReport@peopleschoicecu.com.au"
    $To = "DesktopSupport@peopleschoicecu.com.au"
    $CC = "vuniyal@peopleschoicecu.com.au", "AChamberlain@peopleschoicecu.com.au"
    $SMTPServer = "InternalMail.ACCU.local"

# Gather Machine Variables
    $ComputerName = $env:COMPUTERNAME
    $Manufacturer = (Get-WmiObject -Class:Win32_ComputerSystem).Manufacturer
    $Model = (Get-WmiObject -Class:Win32_ComputerSystem).Model
    $UUID = (Get-WmiObject -Class:Win32_ComputerSystemProduct).UUID
    $SerialNumber = (Get-WmiObject -Class:Win32_BIOS).SerialNumber
    $currentTime = Get-Date -Format "yyyy-MM-dd HH:mm"

# Gather Operating System Variables
    $OSCaption = (Get-WmiObject -Class Win32_OperatingSystem).Caption
    $OSVersion = (Get-WmiObject -Class Win32_OperatingSystem).Version

# Set Subject and Message Body Variables
    $Subject = "SCCM Task: Windows OSD TaskSequence | $ComputerName"
    $MessageBody = @"
<body style="font-family:Calibri;">
$ComputerName completed the task sequence successfully.<br><br />
<hr /><br />
<b>Execution Time</b>: $currentTime <br />
<b>Computer Name</b>: $ComputerName <br />
<b>Manufacturer</b>: $Manufacturer <br />
<b>Model</b>: $Model <br />
<b>Serial Number</b>: $SerialNumber <br />
<b>UUID/BIOS GUID</b>: $UUID <br />
<b>Operating System</b>: $OSCaption - Version $OSVersion <br />
</body>
"@

# Send E-Mail to SMTP Server
Send-MailMessage -Subject $Subject -Body $MessageBody -BodyAsHtml -From $From -To $To -Cc $CC –SMTPServer $SMTPServer