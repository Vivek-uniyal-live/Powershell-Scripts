<#
.SYNOPSIS
  Designed to create a bunch of default collections in SCCM
.DESCRIPTION
  Will create:
  - Collections for each Hardware Manufacturer in the SCCM DB
  - Collections for each Hardware Model in the SCCM DB
  - A collection for 'All Workstations'
  - A collection for 'All Servers'
  - Collections for each Operating System (Server/Client) limited to the above
  - Collections for each Operating System in the SCCM DB (limited to the above)

.PARAMETER <Parameter_Name>
    Create-SCCMCollections
	
.INPUTS
  None

.OUTPUTS
  Log file written to the same folder as the script, called Create-Collections_<ReverseDate>.log

.NOTES
  Version:        1.5
  Author:         Leon Zippel
  Creation Date:  12/02/2019
  Purpose/Change: Initial script development

  1.1 - 27/09/2018 - Optimised base operating system collection creation function, Added additional queries
  1.2 - 07/01/2019 - Updated OS version numbers
  1.3 - 22/01/2019 - Bunch of changes and optimisations
  1.4 - 12/02/2019 - Updated Models to include manufacturer in collection name
  1.5 - 23/08/2019 - Updated OS version numbers
  1.6 - 16/06/2020 - Updated OS version numbers

 - The SCCM console needs to be installed on the PC where this script is run
 - The script will use Role Based Access Control, so the user running it doesn't need 'Administrator' access to the SCCM DB
  
.EXAMPLE
  Create-SCCMCollections

#>

#Log Stuff
$MyDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
#$LogFile = "$MyDir\Create-Collections_$((Get-Date).ToString("yyyyMMdd")).log"
$LogFile = "C:\Windows\Temp\Create-Collections_$((Get-Date).ToString("yyyyMMdd")).log"


# Site configuration
$SiteCode = "PRD" # Site code 
$ProviderMachineName = "admiral.pcprd.com.au" # SMS Provider machine name
$CurrLoc = Get-Location
$ConsoleScripts = "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors



Function Main{

    # Import the ConfigurationManager.psd1 module 
    if(!(Get-Module ConfigurationManager)) 
        {
        If(Test-Path $ConsoleScripts)
            {
            WriteLog -Text "Importing Configuration Manager Module" -Type success
            Import-Module $ConsoleScripts @initParams
            }
        else
            {
            WriteLog -Text "Configuration Manager Module not installed!" -Type error
            Return
            }
        }

    # Connect to the site's drive if it is not already present
    if(!(Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams

    DoManufacturers
    DoModels
    DoBaseOperatingSystemCollections
    DoOperatingSystemVersions
	Set-Location $CurrLoc
}

Function DoBaseOperatingSystemCollections{

    WriteLog -Text "Creating base operating system collections"
    
    $FolderName = "Operating Systems - Base"
    $FolderRootPath = ".\DeviceCollection"

    $StaticCollections = @(
        [pscustomobject]@{Name="All Windows Workstations";LimitCollection="All Systems";Query="select *  from  SMS_R_System where SMS_R_System.OperatingSystemNameandVersion like `"%Workstation%`""},
        [pscustomobject]@{Name="All Windows Servers";LimitCollection="All Systems";Query="select *  from  SMS_R_System where SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 1909 (SAC)";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Caption like `"%1909%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 1903 (SAC)";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Caption like `"%1903%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 2019 (LTSC)";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Caption like `"%2019%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 1809 (SAC)";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Caption like `"%1809%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 1803 (SAC)";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Caption like `"%1803%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 1709 (SAC)";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Caption like `"%1709%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 2016 (LTSC)";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Caption like `"%2016%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 2012 R2";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"6.3%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 2012";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"6.2%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 2008 R2";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"6.1%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 2008";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"6.0%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Server 2003";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"5.2%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Server%`""},
        [pscustomobject]@{Name="All Windows 10";LimitCollection="All Windows Workstations";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"10%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Workstation%`""},
        [pscustomobject]@{Name="All Windows 8.1";LimitCollection="All Windows Workstations";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"6.3%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Workstation%`""},
        [pscustomobject]@{Name="All Windows 8";LimitCollection="All Windows Workstations";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"6.2%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Workstation%`""},
        [pscustomobject]@{Name="All Windows 7";LimitCollection="All Windows Workstations";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"6.1%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Workstation%`""},
        [pscustomobject]@{Name="All Windows Vista";LimitCollection="All Windows Workstations";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"6.0%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Workstation%`""},
        [pscustomobject]@{Name="All Windows XP";LimitCollection="All Windows Workstations";Query="select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"5.1%`" and SMS_R_System.OperatingSystemNameandVersion like `"%Workstation%`""},
        [pscustomobject]@{Name="All Virtual Servers";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System where SMS_R_System.IsVirtualMachine = `"True`""}
        [pscustomobject]@{Name="All Physical Servers";LimitCollection="All Windows Servers";Query="select *  from  SMS_R_System where SMS_R_System.IsVirtualMachine = `"False`""}
        [pscustomobject]@{Name="All Virtual Workstation";LimitCollection="All Windows Workstations";Query="select *  from  SMS_R_System where SMS_R_System.IsVirtualMachine = `"True`""}
        [pscustomobject]@{Name="All Physical Workstations";LimitCollection="All Windows Workstations";Query="select *  from  SMS_R_System where SMS_R_System.IsVirtualMachine = `"False`""}
        [pscustomobject]@{Name="All x64 Operating Systems";LimitCollection="All Systems";Query="select *  from  SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceId = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM.SystemType = `"x64-based PC`""},
        [pscustomobject]@{Name="All x86 Operating Systems";LimitCollection="All Systems";Query="select *  from  SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceId = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM.SystemType = `"x86-based PC`""},
        [pscustomobject]@{Name="All Laptops and Tablets";LimitCollection="All Windows Workstations";Query="select *  from SMS_R_System inner join SMS_G_System_SYSTEM_ENCLOSURE on SMS_G_System_SYSTEM_ENCLOSURE.ResourceID = SMS_R_System.ResourceId where SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes in ( `"8`", `"9`", `"10`", `"14`", `"31`" )"}
    )

    Foreach($SC in $StaticCollections)
        {
        CreateCollection -Name $SC.Name -ScheduleHours 6 -FolderName $FolderName -FolderRootPath $FolderRootPath -LimitCollectionName $SC.LimitCollection -MemberQuery $SC.Query
        }

    WriteLog -Text "Finished processing Base Operating System Collections" -Type success
}

Function DoOperatingSystemVersions{

    $FolderName = "Operating Systems - Versions"
    $FolderRootPath = ".\DeviceCollection"
    

    $Query = "select distinct SMS_G_System_OPERATING_SYSTEM.Version, SMS_G_System_OPERATING_SYSTEM.Caption from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId"
    WriteLog -Text "Running Operating Systems Query"
    $OperatingSystems = RunQuery -Query $Query
    WriteLog -Text "$($OperatingSystems.count) results returned"
    Foreach($OS in $OperatingSystems | Where{$_.Caption -ne ""})
        {
        $CollectionName = ""
        $LimitCollection = "All Systems"
        $QueryVal = "*"
        If($OS.Caption -ilike "*Server*") #Server
            {
            Switch -wildcard ($OS.Version)
                {
                #Catchall versions
                '10.0*'{$CollectionName = $OS.Caption; $LimitCollection = "All Windows Servers";$QueryVal ='10.0%' ;Continue}
                '6.3*'{$CollectionName = $OS.Caption; $LimitCollection = "All Server 2012 R2";$QueryVal ='6.3%' ;Continue}
                '6.2*'{$CollectionName = $OS.Caption; $LimitCollection = "All Server 2012";$QueryVal ='6.2%' ;Continue}
                '6.1*'{$CollectionName = $OS.Caption; $LimitCollection = "All Server 2008 R2";$QueryVal ='6.1%' ;Continue}
                '6.0*'{$CollectionName = $OS.Caption; $LimitCollection = "All Server 2008";$QueryVal ='6.0%' ;Continue}
                '5.2*'{$CollectionName = $OS.Caption; $LimitCollection = "All Server 2003";$QueryVal ='5.2%' ;Continue}
                #All else fails
                default{$CollectionName = $OS.Caption;$QueryVal = $OS.Version}
                }
            }
        Else #Client
            {
            Switch -WildCard ($OS.Version)
                {
                #Specific Versions
                '10.0.10240'{$CollectionName = $OS.Caption + " RTM"; $LimitCollection = "All Windows 10";$QueryVal ='10.0.10240' ;Continue}
                '10.0.10586'{$CollectionName = $OS.Caption + " 1511"; $LimitCollection = "All Windows 10";$QueryVal ='10.0.10586' ;Continue}
                '10.0.14393'{$CollectionName = $OS.Caption + " 1607"; $LimitCollection = "All Windows 10";$QueryVal ='10.0.14393' ;Continue}
                '10.0.15063'{$CollectionName = $OS.Caption + " 1703"; $LimitCollection = "All Windows 10";$QueryVal ='10.0.15063' ;Continue}
                '10.0.16299'{$CollectionName = $OS.Caption + " 1709"; $LimitCollection = "All Windows 10";$QueryVal ='10.0.16299' ;Continue}
                '10.0.17134'{$CollectionName = $OS.Caption + " 1803"; $LimitCollection = "All Windows 10";$QueryVal ='10.0.17134' ;Continue}
		        '10.0.17763'{$CollectionName = $OS.Caption + " 1809"; $LimitCollection = "All Windows 10";$QueryVal ='10.0.17763' ;Continue}
                '10.0.18362'{$CollectionName = $OS.Caption + " 1903"; $LimitCollection = "All Windows 10";$QueryVal ='10.0.18362' ;Continue}
                '10.0.18363'{$CollectionName = $OS.Caption + " 1909"; $LimitCollection = "All Windows 10";$QueryVal ='10.0.18363' ;Continue}
                '10.0.19041'{$CollectionName = $OS.Caption + " 2004"; $LimitCollection = "All Windows 10";$QueryVal ='10.0.19041' ;Continue}
                '6.3.9600'{$CollectionName = $OS.Caption + " Update 1"; $LimitCollection = "All Windows 8.1";$QueryVal ='6.3.9600' ;Continue} #Windows 8.1 Update 1
                '6.3.9200'{$CollectionName = $OS.Caption; $LimitCollection = "All Windows 8.1";$QueryVal ='6.3.9200' ;Continue} #Windows 8.1
                '6.2.9200'{$CollectionName = $OS.Caption; $LimitCollection = "All Windows 8";$QueryVal ='6.2.9200' ;Continue} #Windows 8
                '6.1.7601'{$CollectionName = $OS.Caption + " SP1"; $LimitCollection = "All Windows 7";$QueryVal ='6.1.7601' ;Continue} #Windows 7 SP1
                '6.1.7600'{$CollectionName = $OS.Caption; $LimitCollection = "All Windows 7";$QueryVal ='6.1.7600' ;Continue} #Windows 7
                '6.0.6000'{$CollectionName = $OS.Caption; $LimitCollection = "All Windows Vista";$QueryVal ='6.0.6000' ;Continue} #Vista
                '6.0.6001'{$CollectionName = $OS.Caption + " SP1"; $LimitCollection = "All Windows Vista";$QueryVal ='6.0.6001' ;Continue} #Vista SP1
                '6.0.6002'{$CollectionName = $OS.Caption + " SP2"; $LimitCollection = "All Windows Vista";$QueryVal ='6.0.6002' ;Continue} #Vista SP2
                '5.1.2600'{$CollectionName = $OS.Caption; $LimitCollection = "All Windows XP";$QueryVal ='5.1.2600' ;Continue}
                #Catchall Versions
                '10.0*'{WriteLog -Text "Undefined Ver $($OS.Version)" -Type Error; $CollectionName = $OS.Caption; $LimitCollection = "All Windows 10";$QueryVal ='10.0%' ;Continue}
                '6.3*'{WriteLog -Text "Undefined Ver $($OS.Version)" -Type Error; $CollectionName = $OS.Caption; $LimitCollection = "All Windows 8.1";$QueryVal ='6.3%' ;Continue}
                '6.2*'{WriteLog -Text "Undefined Ver $($OS.Version)" -Type Error; $CollectionName = $OS.Caption; $LimitCollection = "All Windows 8";$QueryVal ='6.2%' ;Continue}
                '6.1*'{WriteLog -Text "Undefined Ver $($OS.Version)" -Type Error; $CollectionName = $OS.Caption; $LimitCollection = "All Windows 7";$QueryVal ='6.1%' ;Continue}
                '6.0*'{WriteLog -Text "Undefined Ver $($OS.Version)" -Type Error; $CollectionName = $OS.Caption; $LimitCollection = "All Windows Vista";$QueryVal ='6.0%' ;Continue}
                '5.1*'{WriteLog -Text "Undefined Ver $($OS.Version)" -Type Error; $CollectionName = $OS.Caption; $LimitCollection = "All Windows XP";$QueryVal ='5.1%' ;Continue}
                #All else fails
                default{WriteLog -Text "Undefined Ver $($OS.Version)" -Type Error; $CollectionName = $OS.Caption;$QueryVal = $OS.Version}
                }
            }

        $CollectionName = $CollectionName.Replace("Microsoft ","")
        
        #Workstations don't need filtering on the caption, servers do. Adjust the query if server.
        $ColQuery = "select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"$QueryVal`""
        If($OS.Caption -ilike "*Server*"){$ColQuery = "select *  from  SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceId = SMS_R_System.ResourceId where SMS_G_System_OPERATING_SYSTEM.Version like `"$QueryVal`" AND SMS_G_System_OPERATING_SYSTEM.Caption = `"$($OS.Caption)`""}
        CreateCollection -Name $CollectionName `
            -ScheduleHours 6 `
            -FolderName $FolderName `
            -FolderRootPath $FolderRootPath `
            -LimitCollectionName $LimitCollection `
            -MemberQuery $ColQuery

        }
    WriteLog -Text "Finished processing Operating System Version Collections" -Type success
}

Function DoManufacturers{

    $FolderName = "Hardware - Manufacturers"
    $FolderRootPath = ".\DeviceCollection"
    
    $Query = "select distinct SMS_G_System_COMPUTER_SYSTEM.Manufacturer from  SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceId order by SMS_G_System_COMPUTER_SYSTEM.Manufacturer"
    WriteLog -Text "Running Manufacturer Query"
    $Manufacturers = RunQuery -Query $Query
    WriteLog -Text "$($Manufacturers.count) results returned"
    Foreach($MF in $Manufacturers | Where{$_.Manufacturer -ne ""})
        {
        CreateCollection -Name $MF.Manufacturer `
            -ScheduleHours 6 `
            -FolderName $FolderName `
            -FolderRootPath $FolderRootPath `
            -LimitCollectionName "All Systems" `
            -MemberQuery "select *  from  SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceId = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM.Manufacturer = `"$($MF.Manufacturer)`""
        }
    WriteLog -Text "Finished processing Manufacturers" -Type success
}

Function DoModels{

    $FolderName = "Hardware - Models"
    $FolderRootPath = ".\DeviceCollection"
    
    $Query = "select distinct SMS_G_System_COMPUTER_SYSTEM.Manufacturer, SMS_G_System_COMPUTER_SYSTEM.Model from  SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceId order by SMS_G_System_COMPUTER_SYSTEM.Manufacturer, SMS_G_System_COMPUTER_SYSTEM.Model"
    WriteLog -Text "Running Models Query"
    $Models = RunQuery -Query $Query
    WriteLog -Text "$($Models.count) results returned"
    Foreach($MD in $Models | Where{$_.Model -ne ""})
        {
        CreateCollection -Name "$($MD.Manufacturer) - $($MD.Model)" `
            -ScheduleHours 6 `
            -FolderName $FolderName `
            -FolderRootPath $FolderRootPath `
            -LimitCollectionName "All Systems" `
            -MemberQuery "select *  from  SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceId = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM.Model = `"$($MD.Model)`""
        }
    WriteLog -Text "Finished processing Models" -Type success
}

Function CreateCollection{
    param([string]$Name,
          [int]$ScheduleHours,
          [string]$MemberQuery,
          [string]$FolderName,
          [string]$FolderRootPath,
          [string]$LimitCollectionName)

    $FolderPath = "$FolderRootPath\$FolderName"

    If(!(Get-Item -Path $FolderPath -ErrorAction SilentlyContinue))
        {WriteLog -Text "Creating base $FolderName folder"; New-Item -Name $FolderName -Path $FolderRootPath}

    If(Get-CMDeviceCollection -Name $Name)
        {WriteLog -Text "Existing collection found for $Name" -Type warning}
    else
        {
        WriteLog -Text "Creating Collection $Name"
        $Schedule = New-CMSchedule -Start (Get-Date).ToString("G") -RecurInterval Hours -RecurCount $ScheduleHours
        $NewCol = New-CMDeviceCollection -Name $Name -RefreshSchedule $Schedule -RefreshType 6 -LimitingCollectionName $LimitCollectionName
        Add-CMDeviceCollectionQueryMembershipRule `
            -CollectionName $Name `
            -RuleName "Query-$Name" `
            -QueryExpression $MemberQuery
        Move-CMObject -FolderPath $FolderPath -InputObject $NewCol
        $Schedule = $Null
        $NewCol = $Null
        }
}

Function RunQuery{

    param([string]$Query,
          [string]$LimitingCollection = $null)


    #Generate a temporary name for the query
    $TmpQueryName = "TMP_$(-join ((65..90) + (97..122) | Get-Random -Count 5 | % {[char]$_}))"
    #Create the query
    If($LimitingCollection)
        {New-CMQuery -Name $TmpQueryName -Expression $Query -LimitToCollectionId $LimitingCOllection | Out-Null}
    Else
        {New-CMQuery -Name $TmpQueryName -Expression $Query | Out-Null}

    #Invoke the temporary query, get the results, then remove it
    $Results = $null
    $Results = Invoke-CMQuery -Name $TmpQueryName
    Remove-CMQuery -Name $TmpQueryName -Force

    #Need to determine what kind of data we got back. It seems if you're only returning data for a particular class, then
    #you retrieve the data differently to when you're returning for multiple classes.
    #When returning for multiple classes, the top level properties will be SMS_<something>.
    If($Null -ne ($Results[0].PropertyNames | ?{$_ -ilike "SMS_*"})){$MultipleClasses = $True}else{$MultipleClasses = $False}

    #Create a primary array to hold all the custom objects we're going to create
    $FinalResults = @()
    #Process the results and add them to a custom object so we can display them correctly 
    Foreach($R in $Results)
	    {
        #Create a custom object to hold the Query Result
        $QR = $Null
        $QR = New-Object -TypeName PSObject
        #Process differently depending on whether we've got multiple classes or not
        If($MultipleClasses) #Multiple classes process
            {
            #Go through each class
            Foreach($PN in $R.PropertyNames)
                {
                #Grab the item
                $Detail = $R.GetSingleItem($PN)
                #Go through each property of the class
                Foreach($PN2 in $Detail.PropertyNames)
                    {
                    #Add it to the custom object. If it's an array, convert it to a string separated by commas
                    If($Detail.Item($PN2).IsArray){$QR | Add-Member -Type NoteProperty -Name $PN2 -Value $($Detail.Item($PN2).ObjectArrayValue -join ",")}
                    else{$QR | Add-Member -Type NoteProperty -Name $PN2 -Value $Detail.Item($PN2).StringValue}
                    }
                }
            }
        Else #Single class process
            {
            #Go through each property
            Foreach($PN in $R.PropertyNames)
                {
                #Grab the item
                $Detail = $R.Item($PN)
                #Add it to the custom object. If it's an array, convert it to a string separated by commas
                If($Detail.IsArray){$QR | Add-Member -Type NoteProperty -Name $PN -Value $($Detail.ObjectArrayValue -join ",")}
                else{$QR | Add-Member -Type NoteProperty -Name $PN -Value $Detail.StringValue}
                }
            }

        #Add it to the primary array
        $FinalResults += $QR
	    }

    #At this point we have $FinalResults, but not in the order of the 'Select' statement
    #Format the select part of the statement into an array (based on .)
    $SelectArray = $($Query.substring(0,$Query.ToLower().IndexOf("from"))).Split(".").Split(",").Trim()
    #This'll now be in the format like "select SMS_R_System, Name, SMS_G_System_COMPUTER_SYSTEM, Model, SMS_G_System_COMPUTER_SYSTEM, Manufacturer, sms_r_system, IPAddresses"
    #start with the second item, then grab every second item (these will be the items we wanted to select from the query)
    $SelectStatementsInOrder = for ($i=1;$i -lt $SelectArray.count;$i+=2) {$SelectArray[$i]}

    #Now we have the right output!
    $ToReturn = $($FinalResults | Select $SelectStatementsInOrder)

    Return $ToReturn
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

