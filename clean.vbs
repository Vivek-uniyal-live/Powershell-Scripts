wscript.echo "Stopping CCMEXEC service"
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "net stop ccmexec" ,0, true


wscript.echo "Checking if CCMEXEC service is stopped....."

Set objSWbemServices = GetObject("winmgmts:\\.\root\cimv2")

set ccmsvc = objSWbemServices.Get("win32_service.name='ccmexec'")

Do While ( ccmsvc.State <> "Stopped")
    wscript.echo "ccmexec service state is " + ccmsvc.State
    wscript.sleep(2000)
    set ccmsvc = objSWbemServices.Get("win32_service.name='ccmexec'")
Loop

wscript.echo "ccmexec service state is " + ccmsvc.State

wscript.echo "Removing stuck application CI tasks"

Set objSWbemServices = GetObject("winmgmts:\\.\root\ccm")
Set colSWbemObjectSet = objSWbemServices.InstancesOf("SMS_MaintenanceTaskRequests") 
For Each objSWbemObject In colSWbemObjectSet 
objSWbemObject.Delete_ 
Next 

Set objSWbemServices = GetObject("winmgmts:\\.\root\ccm\xmlstore")
Set colSWbemObjectSet = objSWbemServices.InstancesOf("xmldocument") 
For Each objSWbemObject In colSWbemObjectSet 
objSWbemObject.Delete_ 
Next 


Set objSWbemServices = GetObject("winmgmts:\\.\root\ccm\citasks")
Set colSWbemObjectSet = objSWbemServices.InstancesOf("CCM_citask") 
For Each objSWbemObject In colSWbemObjectSet 
objSWbemObject.Delete_ 
Next 

wscript.echo "Removing stuck update jobs"

Set objSWbemServices = GetObject("winmgmts:\\.\Root\CCM\SoftwareUpdates\DeploymentAgent")
Set colSWbemObjectSet = objSWbemServices.InstancesOf("CCM_AssignmentJobEx1") 
For Each objSWbemObject In colSWbemObjectSet 
objSWbemObject.Delete_ 
Next 

Set objSWbemServices = GetObject("winmgmts:\\.\Root\CCM\SoftwareUpdates\DeploymentAgent")
Set colSWbemObjectSet = objSWbemServices.InstancesOf("ccm_updatesjob") 
For Each objSWbemObject In colSWbemObjectSet 
objSWbemObject.Delete_ 
Next 

wscript.echo "starting ccmexec service"

WshShell.Run "net start ccmexec", 0, true