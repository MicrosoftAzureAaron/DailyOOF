#get current username from local user foldername
function CurrentUserNamefromWindows 
{
	$Global:CurrentUser = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
    Write-Host "CurrentUser is " -NoNewline
	Write-Host "$Global:CurrentUser" -ForegroundColor Blue
}

function get-Alias 
{
	CurrentUserNamefromWindows
	#Write-Host "$Global:CurrentUser " -ForegroundColor Blue -NoNewline

	####back in the day when microsoftsupport.com was a thing
	<#
	Write-Host "please enter the Alias Suffix of the Account to change. Ex. " -NoNewline
	Write-Host "$Global:UserAliasSuffix : " -ForegroundColor Blue -NoNewline

	$Global:UserAliasSuffix = Read-Host
    if($Global:UserAliasSuffix -eq ""){ #if user doesn't input anything use default
		$Global:UserAliasSuffix="@Microsoft.com"
	}
    if($Global:UserAliasSuffix.StartsWith("@")) {#ensure @ at begining
		
	}
	else {
		$Global:UserAliasSuffix = "@" + $Global:UserAliasSuffix
	}
	#>
	
    $Global:UserAlias = "$Global:CurrentUser$Global:UserAliasSuffix"
    Write-Host "UserAlias is " -NoNewline
	Write-Host "$Global:UserAlias" -ForegroundColor Blue
    #Write-Host "UserAliasSuffix is " -NoNewline
	#Write-Host "$Global:UserAliasSuffix" -ForegroundColor Blue
}

function ConnectAlias2EXO 
{
	InstallEXOM #is EXO module installed
	Write-Host "Connecting to your Outlook Account $UserAlias`n" 
	Connect-ExchangeOnline -UserPrincipalName $Global:UserAlias
	Write-Host "Done Connecting"
}

function get-ARC 
{
	#add choice load from file or load from online exchange
	#prefers local store over remote
    $TempPath = $Global:MessageFilePath + "AutoReplyConfig.json"

	if(FileDNE $TempPath) 
	{
        Write-Host "AutoConfig has pre-existing file " $TempPath
        get-ARCFile
		Write-Host "ARC File Loaded from Local File"
		#Write-Host $Global:MailboxARC
	}
    else 
	{
		$Global:MailboxARC = Get-MailboxAutoReplyConfiguration -identity $UserAlias

		$SaveIt = YesNo "Do you want to save a local copy on $TempPath ?"
		if($SaveIt -eq "Yes") 
		{
			Write-Host "AutoConfig is being written to JSON file $TempPath"
			$Global:MailboxARC = Get-MailboxAutoReplyConfiguration -identity $UserAlias
			$Global:MailboxARC | ConvertTo-Json -depth 100 | Set-Content $TempPath
		}
    }
	Write-Host "Current Auto Reply State is : "$Global:MailboxARC.AutoReplyState
}

function get-ARCFile 
{
    $TempPath = $Global:MessageFilePath + "AutoReplyConfig.json"
    $Global:MailboxARC = Get-Content $TempPath | ConvertFrom-Json 
}
function FileDNE($FilePath) 
{
    return (Get-Item -Path $FilePath -ErrorAction Ignore)
}

#set autoreply to scheduled
#this requires start and end times
#will ask for start and end times if they dne
function Set-ARCSTATEScheduled 
{
	if($null -eq $Global:MailboxARC){
		$Global:MailboxARC = get-arc
	}
	if($null -eq $Global:UserAlias){
		$Global:UserAlias = get-Alias
	}

	#is Reply state disabled or enabled by the user manually instead of scheduled
	if($Global:MailboxARC.AutoReplyState -eq "Disabled" -or $Global:MailboxARC.AutoReplyState -eq "Enabled"){
		Write-Host "Auto Reply state is currently set to " $Global:MailboxARC.AutoReplyState
	}

	##gets office hours, if not hardcoded at end of this file, ask user for input
	### need to add days of the week check for the 4x10 works
	$ioh = IsOfficeHours($ioh)
	switch($ioh)
	{
		0 {
			#use todays start and end if ran before shift starts, still need to be reran during or after shift to set for next off period
		}
		1 {
			$Global:StartOfShift = [datetime] $Global:StartOfShift.adddays(1)
		}
		-1 {
			#should be never here
		}
	}

	#$Global:StartOfShift = [datetime] $Global:StartOfShift
	#$Global:EndOfShift = [datetime] $Global:EndOfShift

	#Write-Host ([datetime] $Global:StartOfShift) ([datetime] $Global:EndOfShift)
	#Set-MailboxAutoReplyConfiguration -identity $UserAlias -ExternalMessage $Global:MailboxARC.ExternalMessage -InternalMessage $Global:MailboxARC.InternalMessage -StartTime $Global:EndOfShift -EndTime $Global:StartOfShift -AutoReplyState "Scheduled"
	Set-MailboxAutoReplyConfiguration -identity $UserAlias -AutoReplyState "Scheduled"
	Write-Host "Set Auto Reply state to Scheduled. `nStart time for OOF Message " $Global:EndOfShift "`nOOF Message will End at " $Global:StartOfShift
}

function IsOfficeHours($duringshift) 
{
	$duringshift = -1
	#check if it is during shift return bool based on start and end time
	#get start and end times
	#Write-Host ([datetime] $Global:StartOfShift) 
	#Write-Host ([datetime] $Global:EndOfShift)

	$CurrentTime =  Get-Date #-Format "MM/dd/yyyy HH:mm"
	$CurrentTime =  [datetime] $CurrentTime

	$Global:StartOfShift = GetShiftTime "start" 
	$Global:EndOfShift = GetShiftTime "end" 

	<#
	if(-not $Global:EndOfShift)
	{
		$Global:StartOfShift = GetShiftTime "start" 
	}
	if(-not $Global:EndOfShift)
	{
		$Global:EndOfShift = GetShiftTime "end" 
	}#>

	#Write-Host ($Global:StartOfShift) 
	#Write-Host ($Global:EndOfShift)
	#Write-Host ($CurrentTime)
	#Write-Host ($CurrentTime -le $Global:EndOfShift)
	#Write-Host ($CurrentTime -ge $Global:StartOfShift)

	$WorkDays = Workdays_of_week($WorkDays)
	
	if($CurrentTime.DayOfWeek -in $WorkDays)
	{
		Write-Host "You should be working today," $CurrentTime.DayOfWeek
		if($CurrentTime -lt $StartOfShift){ 
			Write-Host "Currently Before Shift" ### use todays start and end times, rerun during shift to set for overnight oof
			$duringshift = 0
		}
		elseif($CurrentTime -gt $EndOfShift){
			Write-Host "Currently After Shift" ### use tomorrows start time and todays end time
			$duringshift = 1 
		}
		elseif($CurrentTime -le $Global:EndOfShift -And $CurrentTime -ge $Global:StartOfShift){
			Write-Host "Currently During Shift" ### use tomorrows start time and todays end time
			$duringshift = 1
		}
		else {Write-Host "Twilight Zone"}
	}
	else
	{
		Write-Host "You are not working today," $CurrentTime.DayOfWeek
		### What should be the end time for the OOF Message
		### Next Workday? day++ 
		while(!($CurrentTime.DayOfWeek -in $WorkDays))
		{
			Write-Host $CurrentTime.DayOfWeek " is not currently a work day" $WorkDays
			$CurrentTime = $CurrentTime.adddays(1)			
		}
	}
	return $duringshift
}

function Workdays_of_week($WD) ### this is a function to declar a variable, it will become a switch later to prompt the user if they do not define a set of work days statically like the start and times
{   
	### These are the days of the week that you work
	### Common examples can be uncommented
	### Or edit the default

	### 4 Days Sunday - Wednesday 
	#return $WD = 'Monday', 'Tuesday', 'Wednesday', 'Sunday'

	### 4 Days Wednesday - Saturday
	#return $WD = 'Wednesday', 'Thursday', 'Friday', 'Saturday'

	### Twitter Employee Working 7 days wont need this script
    #return $WD = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'

	### no wednesdays testing
    return $WD = @('Monday', 'Tuesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')

	### Standard Monday - Friday
	#return $WD = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'
}

function GetShiftTime($StartEnd) 
{
	#### add check for start and end times in file
	if($StartEnd -eq "start" -and $Global:MailboxARC.StartTime)
	{
		if((YesNo "Do you want to used the saved $Startend time? ") -eq "Yes")
		{
			return [datetime] $Global:MailboxARC.StartTime
		}
	}
	if($StartEnd -eq "end" -and $Global:MailboxARC.EndTime)
	{
		if((YesNo "Do you want to used the saved $Startend time? ") -eq "Yes")
		{
			return [datetime] $Global:MailboxARC.EndTime
		}
	}

	$PT = "Enter when you $StartEnd your work day. Format 9:00am"
	$ShiftTime = Read-Host -Prompt $PT
	#Write-Host $ShiftTime
	return [datetime] $ShiftTime
} 

function DisconnectEXO 
{
	Disconnect-ExchangeOnline -Confirm:$false
}

function YesNo($Prompt) 
{
	$PT = $Prompt + " [Yes] No"
	$YN = Read-Host -Prompt $PT
    if($YN -eq "" -or $YN -eq "Yes"  -or  $YN -eq "YES"  -or  $YN -eq "Y"  -or  $YN -eq "y"){ #if user doesn't input anything use default
		return "Yes"
	}	
	return 
}

function InstallEXOM 
{
	if ((Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
		#Write-Host "ExchangeOnlineManagement exists, not installing`n"
        #no output if it is installed, less chatty
		Update-Module -Name ExchangeOnlineManagement
        return
	} 
	else {
		Write-Host "ExchangeOnlineManagement does not exist, installing`n"
		Install-Module -Name ExchangeOnlineManagement -force
	}
	return
}

$Global:CurrentUser=$null #obatined from user folder name
$Global:UserAlias=get-Alias #$null #combined with suffix

$Global:UserAliasSuffix="@Microsoft.com"
$Global:MailboxARC=$null #auto reply configuration object

$Global:EndOfShift=$null#[datetime]"6:00pm"
$Global:StartOfShift=$null#[datetime]"9:00am"
$Global:AliasPath = $Global:UserAlias.replace("@","_")
$Global:MessageFilePath= Get-Location
$Global:MessageFilePath= $Global:MessageFilePath.tostring() + "\"