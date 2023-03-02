


#get current username from local user foldername
function CurrentUserNamefromWindows 
{
	$CurrentUser = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
    Write-Host "CurrentUser is " -NoNewline
	Write-Host "$CurrentUser" -ForegroundColor Blue
	return $CurrentUser
}

function get-Alias 
{
	$CurrentUser = CurrentUserNamefromWindows
    $UA = (-join($CurrentUser,$Global:UserAliasSuffix))
    Write-Host "UserAlias is " -NoNewline
	Write-Host "$UA" -ForegroundColor Blue
    #Write-Host "UserAliasSuffix is " -NoNewline
	#Write-Host "$Global:UserAliasSuffix" -ForegroundColor Blue
	return $UA
}

function ConnectAlias2EXO 
{
	InstallEXOM #is EXO module installed
	Write-Host "Connecting to your Outlook Account with Alias $Global:UserAlias`n" 
	Connect-ExchangeOnline -UserPrincipalName $Global:UserAlias
	Write-Host "Done Connecting"
}

function get-ARC 
{
	#add choice load from file or load from online exchange
	#prefers local store over remote
   
	if(FileDNE $Global:MessageFilePath) 
	{
        Write-Host "ARC File stored locally" $Global:MessageFilePath
        get-ARCFile
		Write-Host "ARC File Loaded from Local File"
		#Write-Host $Global:MailboxARC
	}
    else 
	{
		$Global:MailboxARC = Get-MailboxAutoReplyConfiguration -UserPrincipalName $UserAlias

		$SaveIt = YesNo "Do you want to save a local copy on $Global:MessageFilePath ?"
		if($SaveIt -eq "Yes") 
		{
			Write-Host "AutoConfig is being written to JSON file $Global:MessageFilePath"
			$Global:MailboxARC = Get-MailboxAutoReplyConfiguration -UserPrincipalName $UserAlias
			$Global:MailboxARC | ConvertTo-Json -depth 100 | Set-Content $Global:MessageFilePath
		}
    }
	Write-Host "Current Auto Reply State is : "$Global:MailboxARC.AutoReplyState
}

function get-ARCFile 
{
	#Write-Host $Global:MessageFilePath
    $Global:MailboxARC = Get-Content $Global:MessageFilePath -raw | ConvertFrom-Json 
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
	
	#is Reply state disabled or enabled by the user manually instead of scheduled
	if($Global:MailboxARC.AutoReplyState -eq "Disabled" -or $Global:MailboxARC.AutoReplyState -eq "Enabled"){
		Write-Host "Auto Reply state is currently set to " $Global:MailboxARC.AutoReplyState
	}

	##gets office hours, if not hardcoded at end of this file, ask user for input
	$daystoadd = IsOfficeHours
	$Global:StartOfShift = [datetime] $Global:StartOfShift.adddays($daystoadd)
	$Global:EndOfShift = [datetime] $Global:EndOfShift.adddays($daystoadd)
	
	#$Global:StartOfShift = [datetime] $Global:StartOfShift
	#$Global:EndOfShift = [datetime] $Global:EndOfShift

	#Write-Host ([datetime] $Global:StartOfShift) ([datetime] $Global:EndOfShift)
	#Set-MailboxAutoReplyConfiguration -Identity $UserAlias -ExternalMessage $Global:MailboxARC.ExternalMessage -InternalMessage $Global:MailboxARC.InternalMessage -StartTime $Global:EndOfShift -EndTime $Global:StartOfShift -AutoReplyState "Scheduled"

	Set-MailboxAutoReplyConfiguration -Identity $Global:UserAlias -AutoReplyState "Scheduled"
	Write-Host "Set Auto Reply state to Scheduled. `nOOF Message will start:" $Global:StartofShift "`nOOF Message will End: " $Global:EndOfShift
}

function IsOfficeHours 
{
	$duringshift = -1
	#check if it is during shift return bool based on start and end time
	#get start and end times
	#Write-Host ([datetime] $Global:StartOfShift) 
	#Write-Host ([datetime] $Global:EndOfShift)

	$CuTime =  Get-Date #-Format "MM/dd/yyyy HH:mm"
	$CuTime =  [datetime] $CuTime

	$Global:StartOfShift = GetShiftTime "start" 
	$Global:EndOfShift = GetShiftTime "end" 

	#Write-Host ($Global:StartOfShift) 
	#Write-Host ($Global:EndOfShift)
	#Write-Host ($CuTime)
	#Write-Host ($CuTime -le $Global:EndOfShift)
	#Write-Host ($CuTime -ge $Global:StartOfShift)

	$WorkDays = Workdays_of_week($WorkDays)
	#check if you work today, then calculate the next work day
	if($CuTime.DayOfWeek -in $WorkDays)
	{
		Write-Host "You should be working today," $CuTime.DayOfWeek
		if($CuTime -lt $Global:StartOfShift){ 
			Write-Host "Currently Before Shift" ### use todays start and end times, rerun during shift to set for overnight oof
			$duringshift = 0
		}
		elseif($CuTime -gt $Global:EndOfShift){
			Write-Host "Currently After Shift" ### use tomorrows start time and todays end time
			$duringshift = 1 
		}
		elseif($CuTime -le $Global:EndOfShift -And $CuTime -ge $Global:StartOfShift){
			Write-Host "Currently During Shift" ### use tomorrows start time and todays end time
			$duringshift = 1
		}
		else {Write-Host "Twilight Zone"}
	}
	else
	{
		Write-Host "You are not working today" $CuTime.DayOfWeek
		### What should be the end time for the OOF Message
		### Next Workday? day++ 
		$i = 0
		while(!($CuTime.DayOfWeek -in $WorkDays))
		{
			$i += 1
			Write-Host $CuTime.DayOfWeek -ForegroundColor Red -NoNewline 
			Write-Host " is not currently a work day [" -NoNewline
			Write-Host  $WorkDays -NoNewline -ForegroundColor Blue
			Write-Host "]"
			$CuTime = $CuTime.adddays(1)		
			
		}
		$duringshift = $i
		#Write-Host $CuTime.DayOfWeek
		#Write-Host $Global:StartOfShift.TimeofDay
		Write-Host (-join("The start of the next workday is ",$CuTime.DayOfWeek," ",$Global:StartOfShift))
		
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

	### no wednesdays or thursdays testing
    return $WD = @('Monday', 'Tuesday', 'Friday', 'Saturday', 'Sunday')

	### Standard Monday - Friday
	#return $WD = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'
}

function GetShiftTime($StartEnd) 
{
	if(FileDNE $Global:MessageFilePath) 
	{
		get-ARCFile
		#### check for start and end times in file
		if($StartEnd -eq "start")
		{
			$PT = (-join("Do you want to used the saved $StartEnd of shift time? This is when the OOF message will end ",$Global:MailboxARC.EndTime.TimeofDay))
			if((YesNo $PT -eq "Yes"))
			{
				#Write-Host $Global:MailboxARC.StartTime
				$Global:StartOfShift = $Global:MailboxARC.EndTime
				return [datetime] $Global:MailboxARC.StartTime
			}
		}

		if($StartEnd -eq "end")
		{
			$PT = (-join("Do you want to used the saved $StartEnd of shift time? This is when the OOF message will start ",$Global:MailboxARC.StartTime.TimeofDay))
			if((YesNo $PT -eq "Yes"))
			{
				#Write-Host $Global:MailboxARC.EndTime
				$Global:EndOfShift = $Global:MailboxARC.StartTime
				return [datetime] $Global:MailboxARC.EndTime
			}
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
	$PT = $Prompt + "[Yes] No"
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

$Global:UserAliasSuffix = "@Microsoft.com"
$Global:UserAlias = get-Alias #based on user folder name combined with suffix
$Global:MessageFilePath = Get-Location
$Global:MessageFilePath = (-join($Global:MessageFilePath.tostring(),'\','AutoReplyConfig.json'))
ConnectAlias2EXO