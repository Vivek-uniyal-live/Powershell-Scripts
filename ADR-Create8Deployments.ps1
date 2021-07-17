#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows PowerShell and will connect to the site.
#
# This script was auto-generated at '8/04/2020 4:04:20 PM'.

# Uncomment the line below if running in an environment where script signing is 
# required.
#Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# Site configuration
$SiteCode = "PRD" # Site code 
$ProviderMachineName = "admiral.pcprd.com.au" # SMS Provider machine name

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams




# Get the ADR
$ADRName = "ADR.Windows.10.Office.365"
$ADR = Get-CMAutoDeploymentRule -Fast -Name $ADRName

    $DeploymentDetails = @(
        [pscustomobject]@{CollectionName="SU.Workstations.Dynamic.0-1";AvailableHours="291";RequiredHours="3"},
        [pscustomobject]@{CollectionName="SU.Workstations.Dynamic.2-3";AvailableHours="315";RequiredHours="3"},
        [pscustomobject]@{CollectionName="SU.Workstations.Dynamic.4-5";AvailableHours="339";RequiredHours="3"},
        [pscustomobject]@{CollectionName="SU.Workstations.Dynamic.6-7";AvailableHours="363";RequiredHours="3"},
        [pscustomobject]@{CollectionName="SU.Workstations.Dynamic.8-9";AvailableHours="435";RequiredHours="3"},
        [pscustomobject]@{CollectionName="SU.Workstations.Dynamic.A-B";AvailableHours="459";RequiredHours="3"},
        [pscustomobject]@{CollectionName="SU.Workstations.Dynamic.C-D";AvailableHours="483";RequiredHours="3"},
        [pscustomobject]@{CollectionName="SU.Workstations.Dynamic.E-F";AvailableHours="507";RequiredHours="3"}

    ) 


# Create the deployments
Foreach ($Deployment in $DeploymentDetails)
{
    If(Get-CMAutoDeploymentRuleDeployment -Name $ADRName | where {$_.CollectionName -eq $Deployment.CollectionName})
        {Write-Host "$($Deployment.CollectionName) Already exists - skip!"; Continue}
        else
        {Write-Host "Creating Deployment $($Deployment.CollectionName)"}
    
    # Create the deployment
    $Params = @{
        CollectionName = $Deployment.CollectionName
        EnableDeployment = $true
        SendWakeupPacket = $false
        VerboseLevel = 'OnlyErrorMessages'
        UseUtc = $False
        AvailableTime = $Deployment.AvailableHours
        AvailableTimeUnit = 'Hours'
        DeadlineTime = $Deployment.RequiredHours
        DeadlineTimeUnit = 'Hours'
        UserNotification = 'DisplaySoftwareCenterOnly'
        AllowSoftwareInstallationOutsideMaintenanceWindow = $false
        AllowRestart = $false
        SuppressRestartServer = $false
        SuppressRestartWorkstation = $false
        WriteFilterHandling = $false
        NoInstallOnRemote = $false 
        NoInstallOnUnprotected = $false
        UseBranchCache = $false
    }
    $null = $ADR | New-CMAutoDeploymentRuleDeployment @Params

    # Update the deployment with some additional params not available in the cmdlet
    $ADRDeployment = Get-CMAutoDeploymentRuleDeployment -Name $ADRName | where {$_.CollectionName -eq $Deployment.CollectionName}
    [xml]$DT = $ADRDeployment.DeploymentTemplate
    # If software updates are not available on distribution point in current, neighbour or site boundary groups, download content from Microsoft Updates
    $DT.DeploymentCreationActionXML.AllowWUMU = "true"
    # Allow download over metered connections
    $DT.DeploymentCreationActionXML.AllowUseMeteredNetwork = "true"
    # If any update in this deployment requires a system restart, run updates deployment evaluation cycle after restart
    If ($DT.DeploymentCreationActionXML.RequirePostRebootFullScan -eq $null)
    {
        $NewChild = $DT.CreateElement("RequirePostRebootFullScan")
        [void]$DT.SelectSingleNode("DeploymentCreationActionXML").AppendChild($NewChild)
    }
    $DT.DeploymentCreationActionXML.RequirePostRebootFullScan = "Checked" 
    $ADRDeployment.DeploymentTemplate = $DT.OuterXml
    $ADRDeployment.Put()
}
