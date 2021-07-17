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

    DoDynamicCollections

	Set-Location $CurrLoc
}

Function DoDynamicCollections{

    WriteLog -Text "Creating Dynamic collections"
    
    $FolderName = "Software Updates"
    $FolderRootPath = ".\DeviceCollection"

    $StaticCollections = @(
        [pscustomobject]@{Name="SU.Workstations.Dynamic.0-3";LimitCollection="All Windows Workstations";Query="select * from SMS_R_System where SMS_R_System.SMSUniqueIdentifier like `"%0`" or SMS_R_System.SMSUniqueIdentifier like `"%1`" or SMS_R_System.SMSUniqueIdentifier like `"%2`" or SMS_R_System.SMSUniqueIdentifier like `"%3`""},
        [pscustomobject]@{Name="SU.Workstations.Dynamic.4-7";LimitCollection="All Windows Workstations";Query="select * from SMS_R_System where SMS_R_System.SMSUniqueIdentifier like `"%4`" or SMS_R_System.SMSUniqueIdentifier like `"%5`" or SMS_R_System.SMSUniqueIdentifier like `"%6`" or SMS_R_System.SMSUniqueIdentifier like `"%7`""},
        [pscustomobject]@{Name="SU.Workstations.Dynamic.8-B";LimitCollection="All Windows Workstations";Query="select * from SMS_R_System where SMS_R_System.SMSUniqueIdentifier like `"%8`" or SMS_R_System.SMSUniqueIdentifier like `"%9`" or SMS_R_System.SMSUniqueIdentifier like `"%A`" or SMS_R_System.SMSUniqueIdentifier like `"%B`""},
        [pscustomobject]@{Name="SU.Workstations.Dynamic.C-F";LimitCollection="All Windows Workstations";Query="select * from SMS_R_System where SMS_R_System.SMSUniqueIdentifier like `"%C`" or SMS_R_System.SMSUniqueIdentifier like `"%D`" or SMS_R_System.SMSUniqueIdentifier like `"%E`" or SMS_R_System.SMSUniqueIdentifier like `"%F`""}
    )

    Foreach($SC in $StaticCollections)
        {
        CreateCollection -Name $SC.Name -ScheduleHours 6 -FolderName $FolderName -FolderRootPath $FolderRootPath -LimitCollectionName $SC.LimitCollection -MemberQuery $SC.Query
        }

    WriteLog -Text "Finished processing Dynamic Collections" -Type success
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

