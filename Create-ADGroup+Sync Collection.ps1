<#
.SYNOPSIS
  Designed to assist with creation of AD groups and collections for application deployment
.DESCRIPTION
  Will create:
  - AD Group based on Application Name (replacing APP. with APP-)
  - AD Group will have the applicaiton name detailed in the 'notes' section
  - SCCM Application Deployment Collection folder (if required)
  - SCCM Collection based on Application Name
  - SCCM Collection Query pointing to above AD group
  - Required application deployment for application to the new collection
  
  The script will then distribute the content to the required Distribution Point (or group)

.PARAMETER AppName
  Specifies the application name to use for creation of all objects   
	
.INPUTS
  None

.OUTPUTS
  Log file written to the same folder as the script, called <ScriptName>_<ReverseDate>.log

.NOTES
  Version:        1.0
  Author:         Leon Zippel
  Creation Date:  21/05/2019
  Purpose/Change: Initial script development

  1.# - ##/##/#### - <Description>

 - The SCCM console needs to be installed on the PC where this script is run
 - The AD Powershell module needs to be installed on the PC where this script is run
  
.EXAMPLE
  Create-AppResources.ps1 -AppName APP.Adobe.Reader.DC.14
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '606,359'
$Form.text                       = "SCCM App Resource Creator"
$Form.TopMost                    = $false

$Panel1                          = New-Object system.Windows.Forms.Panel
$Panel1.height                   = 43
$Panel1.width                    = 589
$Panel1.Anchor                   = 'top,right,left'
$Panel1.location                 = New-Object System.Drawing.Point(7,8)

$txt_appname                     = New-Object system.Windows.Forms.TextBox
$txt_appname.multiline           = $false
$txt_appname.width               = 274
$txt_appname.height              = 20
$txt_appname.Anchor              = 'top'
$txt_appname.location            = New-Object System.Drawing.Point(193,12)
$txt_appname.Font                = 'Microsoft Sans Serif,10'

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "App Name"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.Anchor                   = 'top'
$Label1.location                 = New-Object System.Drawing.Point(103,17)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$Panel2                          = New-Object system.Windows.Forms.Panel
$Panel2.height                   = 241
$Panel2.width                    = 587
$Panel2.Anchor                   = 'top,right,left'
$Panel2.location                 = New-Object System.Drawing.Point(8,56)

$txt_log                         = New-Object system.Windows.Forms.TextBox
$txt_log.multiline               = $true
$txt_log.width                   = 568
$txt_log.height                  = 222
$txt_log.Anchor                  = 'top,right,left'
$txt_log.location                = New-Object System.Drawing.Point(9,6)
$txt_log.Font                    = 'Microsoft Sans Serif,10'

$Panel3                          = New-Object system.Windows.Forms.Panel
$Panel3.height                   = 49
$Panel3.width                    = 586
$Panel3.Anchor                   = 'top,right,left'
$Panel3.location                 = New-Object System.Drawing.Point(8,303)

$btn_Run                         = New-Object system.Windows.Forms.Button
$btn_Run.text                    = "Create!"
$btn_Run.width                   = 97
$btn_Run.height                  = 30
$btn_Run.Anchor                  = 'top'
$btn_Run.location                = New-Object System.Drawing.Point(244,11)
$btn_Run.Font                    = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($Panel1,$Panel2,$Panel3))
$Panel1.controls.AddRange(@($txt_appname,$Label1))
$Panel2.controls.AddRange(@($txt_log))
$Panel3.controls.AddRange(@($btn_Run))

$btn_Run.Add_Click({ Main })
$txt_appname.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter')
        {
            $btn_Run.PerformClick()
        }
 })

#Log Stuff
$MyDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$MyScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand)
#$LogFile = "$MyDir\$MyScriptName`_$((Get-Date).ToString("yyyyMMdd")).log"
$LogFile = "C:\Temp\$MyScriptName`_$((Get-Date).ToString("yyyyMMdd")).log"


# Site configuration
$SiteCode = "PRD" # Site code 
$ProviderMachineName = "admiral.pcprd.com.au" # SMS Provider machine name
$CurrLoc = Get-Location #Powershell location prior to running the script
$ConsoleScripts = "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"

# Global Variables
$ApplicationPrefix = "APP." #Prefix used for Applications in SCCM
$ADGroupPrefix = "APP-" #Prefix used for Application groups in AD
$LimitingCollection = "All Windows Workstations" #Top-level workstations collection
$RefreshHours = 1 #Number of hours application collections should refresh after
$CollectionFolderRootPath = ".\DeviceCollection" #Root folder path
$CollectionFolderName = "Applications - Deployment" #Folder in SCCM to put Application Deployment collections
$GroupsOU = "OU=Groups,OU=PRD,DC=pcprd,DC=com,DC=au" #OU to create application AD groups in

#Distribution of content - Can specify DP or DP Group.
$DPName = "All DPs"

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors



Function Main{

    $AppName = $txt_appname.Text

    #Make sure the AD module is installed/imported
    If((Get-Module -ListAvailable -Name 'ActiveDirectory'))
        {If(!(Get-Module -Name 'ActiveDirectory'))
            {WriteLog -Text "Importing AD Powershell Module" -Type success
            Import-Module ActiveDirectory}
         Else
            {WriteLog -Text "AD Powershell Module already imported" -Type success}
        }
    Else
        {WriteLog -Text "Active Directory Powershell Module not installed!" -Type error; Return} 

    # Make sure the ConfigurationManager.psd1 module is installed/imported
    If(!(Get-Module ConfigurationManager)){
        If(Test-Path $ConsoleScripts)
            {WriteLog -Text "Importing Configuration Manager Module" -Type success
            Import-Module $ConsoleScripts @initParams}
        else
            {WriteLog -Text "Configuration Manager Module not installed!" -Type error
            Return}
    }

    #Validate the Group Name to make sure it's valid for AD Groups
    $AppName = $AppName.Trim()
    $RegEx = [Regex]::new('[#,+"\\<>;]')
    If(($RegEx.Matches($AppName)).Value){WriteLog -Text "App name contains characters not valid for AD Groups!" -Type error; Return}

    # Connect to the site's drive if it is not already present
    if(!(Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams
    
    WriteLog -Text "Proceessing Application: $AppName" -Type Success
       
    #Try create the AD group
    If(CreateADGroup -GroupName $ADGroupName -AppName $AppName)
        {WriteLog -Text "AD Group created/updated successfully" -Type success}
    else
        {WriteLog -Text "Failed to create/update AD group!" -Type error; return}
               
    #Try create the Collection
    If(CreateCollection -AppName $AppName `
                        -ScheduleHours $RefreshHours `
                        -ADGroupName $ADGroupName `
                        -FolderName $CollectionFolderName `
                        -FolderRootPath $CollectionFolderRootPath `
                        -LimitCollectionName $LimitingCollection)
        {WriteLog -Text "Collection/Folder Path created successfully" -Type success}
    else
        {WriteLog -Text "Failed to create collection/folder path!" -Type error; return}   
        
}

Function CreateCollection{
    param([string]$AppName,
          [int]$ScheduleHours,
          [string]$ADGroupName,
          [string]$FolderName,
          [string]$FolderRootPath,
          [string]$LimitCollectionName)

    $FolderPath = "$FolderRootPath\$FolderName"
    $Domain = (Get-WmiObject -Class Win32_ComputerSystem).Domain.split(".")[0]
    $Query = "select * from SMS_R_System where SMS_R_System.SystemGroupName = `"$Domain\\$ADGroupName`""

    If(!(Get-Item -Path $FolderPath -ErrorAction SilentlyContinue))
        {WriteLog -Text "Creating base $FolderName folder"
        Try{New-Item -Name $FolderName -Path $FolderRootPath}
        Catch{WriteLog -Text "Failed to create collection folder! Error: $($_.Exception.Message)" -Type error; return $False}}

    If(Get-CMDeviceCollection -Name $AppName)
        {WriteLog -Text "Existing collection found for $AppName" -Type warning}
    else
        {
        Try{WriteLog -Text "Creating Collection $AppName"
            $Schedule = New-CMSchedule -Start (Get-Date).ToString("G") -RecurInterval Hours -RecurCount $ScheduleHours
            $NewCol = New-CMDeviceCollection -Name $AppName -RefreshSchedule $Schedule -RefreshType 6 -LimitingCollectionName $LimitCollectionName
            Add-CMDeviceCollectionQueryMembershipRule `
                -CollectionName $AppName `
                -RuleName "Query-$ADGroupName" `
                -QueryExpression $Query
            Move-CMObject -FolderPath $FolderPath -InputObject $NewCol
            $Schedule = $Null
            $NewCol = $Null}
        Catch{WriteLog -Text "Unable to create collection $AppName`! Error: $($_.Exception.Message)" -Type error; return $false}

        }
    Return $True
}


Function CreateADGroup{
    Param(
    [String]$GroupName,
    [String]$AppName
    )
    
    If(Get-ADGroup -Filter {SamAccountName -eq $GroupName})
        {
        #Exists, check if we need to update the Notes field
        $CurrentNote = (Get-ADGroup $GroupName -Properties Info).Info
        If((Get-ADGroup $GroupName -Properties Info).Info -ine $AppName)
            {WriteLog -Text "Existing AD Group note field is incorrect. Updating..."
            Try{Set-ADGroup $GroupName -Replace @{info=$AppName}}
            Catch{WriteLog - Text "Failed to update AD group notes field! Error: $($_.Exception.Message)" -Type error; return $false}
            }
        }
    Else
        {
        WriteLog -Text "Creating AD group with Notes field: $AppName"
        Try{;New-ADGroup -Name $GroupName `
                -SamAccountName $GroupName `
                -DisplayName $GroupName `
                -OtherAttributes @{info=$AppName} `
                -GroupScope Global `
                -GroupCategory Security `
                -Path $GroupsOU}
        Catch{WriteLog -Text "Failed to create AD group! Error: $($_.Exception.Message)" -Type error; return $false}
       }
   return $true

}

Function WriteLog{
    Param(
    [String]$Text,
    [ValidateSet('info','error','success', 'warning')][System.String]$Type="info"
    )
    $Text = "$(Get-Date -Format 'G') $Text"
    $Text | Add-Content -path $Logfile
    $txt_log.Text = $txt_log.Text + $Text + [Environment]::NewLine

    #Switch($Type)
    #{
    #    "info"{$Text | Add-Content -path $Logfile -PassThru | Write-Host -ForegroundColor White}
    #    "error"{$Text | Add-Content -path $Logfile -PassThru | Write-Host -ForegroundColor Red}
    #    "success"{$Text | Add-Content -path $Logfile -PassThru | Write-Host -ForegroundColor Green}
    #    "warning"{$Text | Add-Content -path $Logfile -PassThru | Write-Host -ForegroundColor Yellow}
    #}
}




#Write your logic code here
[void]$Form.ShowDialog()