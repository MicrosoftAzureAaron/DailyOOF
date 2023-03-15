$global:StartOfShift =  [datetime]"9:00am" #$null
$global:EndOfShift = [datetime]"6:00pm" #$null 
$UserAliasSuffix = "@microsoft.com"
$WD = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')

#Get current username from local user foldername
function Get-UsernameFromWindows 
{
	$CurrentUser = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
    #Write-Host "CurrentUser is " -NoNewline
	#Write-Host "$CurrentUser" -ForegroundColor Blue
	return $CurrentUser
}

#Get alias from userfolder, if this fails, exo connection will prompt for creds
function Get-Alias 
{
	if($UserAliasSuffix -eq "" -or $null -eq $UserAliasSuffix)
	{
		$UserAliasSuffix = Get-Suffix
	}
	$CurrentUser = Get-UsernameFromWindows
    return (-join($CurrentUser,$UserAliasSuffix))
	#Write-Host "Current account is " -NoNewline
	#Write-Host "${global:UserAlias}" -ForegroundColor Blue
}
function Get-Suffix 
{
    Write-Host "Current suffix is ${UserAliasSuffix}"
	$PT = "What email suffix would you like to use? Format @microsoft.com"
	$UserAliasSuffix = Read-Host -Prompt $PT
	return $UserAliasSuffix
}

#connect to exchange online
function Get-EXOConnection 
{
	Get-EXOM #is EXO module installed
	#Write-Host "Current account is " -NoNewline
	#Write-Host "${global:UserAlias}" -ForegroundColor Blue
	#Write-Host "Connecting to your Outlook Account with alias $global:UserAlias " 
	Connect-ExchangeOnline -UserPrincipalName $global:UserAlias 
	#Write-Host "Done Connecting"
}

function Get-ARCFilePath 

{
	$AFP = Get-Location #store local copy in same folder as script
	$AFP = (-join($AFP.tostring(),'\','AutoReplyConfig.json'))
	return $AFP
}

#write current config to file, warn about overwrite
function Set-ARCFile
{
	$ARCFilePath = Get-ARCFilePath
	# if(Test-Path $ARCFilePath) 
	# {
    #     ###file exists do you want to overwrite
	# 	$Q = YesNo "Auto Reply config file already exists, over write ${ARCFilePath}?"
	# }
	# else
	# {
	# 	###write file
	# 	$Q = YesNo "No local copy found, do you want to save a local copy on ${ARCFilePath}?"
	# }
	# if($Q -eq "Yes") 
	# {
	# 	Write-Host "Auto Reply config is being written to JSON file from current configuration to ${ARCFilePath}"
	# 	$MailboxARC | ConvertTo-Json -depth 100 | Set-Content $ARCFilePath
	# }

	#This is only called when writing the file, no need to check to overwrite
	$MailboxARC | ConvertTo-Json -depth 100 | Set-Content $ARCFilePath
}

#Get current config from online
#save to local file
function Get-ARC
{
	$MailboxARC = Get-MailboxAutoReplyConfiguration -Identity $global:UserAlias #get arc
	Set-ARCFile #Always write the file to disk
	Return $MailboxARC
}

#read the locally stored file
function Get-ARCFile 
{
	$ARCFilePath = Get-ARCFilePath
	#Write-Host $ARCFilePath
    return Get-Content $ARCFilePath -raw | ConvertFrom-Json 
}


#Set auto reply to scheduled/endabled/disabled
function Set-ARCState($S)
{
	#get current configuration
	$MailboxARC = Get-ARC
	Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState

	if(!$S)
	{
		$S = Read-Host -Prompt "What mode should Auto Reply be set to?`n1. Enabled`n2. Disabled`n3. Scheduled`nChoice "
	}
	switch($S)
	{
		'1'
		{
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -AutoReplyState "Enabled"
		}
		'2'
		{
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -AutoReplyState "Disabled"
		}
		'3'
		{
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -AutoReplyState "Scheduled"
		}
	}	
	#Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState
	#update json
	$MailboxARC = Get-ARC
}

#Set auto reply start and end times
function Set-ARCTimes
{
	##Gets office hours, if not hardcoded at the start of this file, ask user for input
	if($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift){Get-ShiftTime}
	
	#get current configuration
	$MailboxARC = Get-ARC

	$daysToAdd = 0
	#how many days till next day of work
	$daysToAdd = Get-Schedule

	#convert daily time to todays time
	$hours = Get-Date $global:StartOfShift
	# Write-Host $global:StartOfShift
	# Write-Host $hours
	$global:StartOfShift = [datetime] (Get-Date).Date.AddHours($hours.Hour)

	#add the number of days till next shift to the time for when the OOF message should end, aka the START of your next shift
	$global:StartOfShift = $global:StartOfShift.adddays($daysToAdd)

	#convert daily time to todays time
	$hours = Get-Date $global:EndOfShift
	$global:EndOfShift = [datetime] (Get-Date).Date.AddHours($hours.Hour)

	#Write-Host "Current Online start:" $MailboxARC.StartTime "`nCurrent Online will End: " $MailboxARC.EndTime
	#Write-Host "Live Config start:" $global:EndOfShift "`nLive Config will End: " $global:StartOfShift

	#Set start and end time for scheduled auto reply
	Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -StartTime $global:EndOfShift -EndTime $global:StartOfShift
	
	#Write Current Config to file
	$MailboxARC = Get-ARC
}

#Set auto reply message
function Set-ARCMessage($IOE,$message)
{

	switch -Regex ($IOE)
	{
		"Internal"
		{
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -InternalMessage $message 
		}
		"External"
		{
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -ExternalMessage $message  
		}
		"Both"
		{
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -ExternalMessage $message -InternalMessage $message 
		}
	}
}

#save online message to html file
function Set-ARCmessagefile
{
	$ARCMessageFile = Get-Location #store local copy in same folder as script
	$ARCMessageFile = (-join($ARCMessageFile.tostring(),'\','message.html'))
	#Write-Host $ARCMessageFile
	$text = $MailboxARC.ExternalMessage.tostring()
	$text | Out-File -FilePath $ARCMessageFile
	Write-Host "Message file saved as $ARCMessageFile"
}

#returns the number of days till next work day
function Get-Schedule
{
	if($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift){Get-ShiftTime}

	$duringshift = 0
	$CuTime =  Get-Date #-Format "MM/dd/yyyy HH:mm"
	$CuTime =  [datetime] $CuTime

	#what days of the week do you work hard code it if you dont wanna be asked
	$WorkDays = Get-WD

	if(!($CuTime.DayOfWeek -in $WorkDays))
	{
		$i = 0
		Write-Host "You are not working today" $CuTime.DayOfWeek
		while(!($CuTime.DayOfWeek -in $WorkDays))
		{
			$i += 1
			#Write-Host $CuTime.DayOfWeek -ForegroundColor Red -NoNewline 
			#Write-Host " is not currently a work day [" -NoNewline
			#Write-Host  $WorkDays -NoNewline -ForegroundColor Blue
			#Write-Host "]"
			$CuTime = $CuTime.adddays(1)		
		}
		$duringshift = $i
		#Write-Host $CuTime.DayOfWeek
		#Write-Host $global:StartOfShift.TimeOfDay
		Write-Host (-join("The start of the next workday is ",$CuTime.DayOfWeek," ",$global:StartOfShift.TimeOfDay))
	}
	else
	{
		if($CuTime -lt $global:StartOfShift)
		{ 
			#Write-Host "${CuTime} Currently Before Shift" ### use todays start and end times, rerun during shift to Set for overnight oof
			$duringshift = 0
		}
		elseif($CuTime -gt $global:EndOfShift)
		{
			#Write-Host "${CuTime} Currently After Shift"### use tomorrows start time and todays end time
			$duringshift = 1
		}
		elseif($CuTime -le $global:EndOfShift -And $CuTime -ge $global:StartOfShift)
		{
			#Write-Host "${CuTime} Currently During Shift" ### use tomorrows start time and todays end time
			$duringshift = 1
		}
		else {
			Write-Host "Twilight Zone"
		}
	}
	return $duringshift
}

function Get-WD 
{   
	### this is a function to either Set an array of days of the week that you work by uncommenting or configuring your own line below
	### These are the days of the week that you work
	### Common examples can be uncommented
	### Or edit the default

	### 4 Days Sunday - Wednesday 
	#$WD = @('Monday', 'Tuesday', 'Wednesday', 'Sunday')

	### 4 Days Wednesday - Saturday
	#$WD = @('Wednesday', 'Thursday', 'Friday', 'Saturday')

	### Twitter Employee Working 7 days wont need this script
    #$WD = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')

	### no wednesdays or thursdays testing
    #$WD = @('Monday', 'Tuesday', 'Friday', 'Saturday', 'Sunday')

	### Standard Monday - Friday
	#$WD = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')

	if(!$WD)
	{
		$S = Read-Host -Prompt "Which of the following matches your weekly work schedule`n1. 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'`n2. 'Monday', 'Tuesday', 'Wednesday', 'Sunday'`n3. 'Wednesday', 'Thursday', 'Friday', 'Saturday'`nChoice "
		switch($S)
		{
			'1'
			{
				$WD = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
			}
			'2'
			{
				$WD = @('Monday', 'Tuesday', 'Wednesday', 'Sunday')
			}
			'3'
			{
				$WD = @('Wednesday', 'Thursday', 'Friday', 'Saturday')
			}
		}
		Write-Host $WD
	}
	return $WD
}

#what time do you start or end your shift
function Get-ShiftTime
{
	#only ask user if not hardcoded
	$PT = "Enter when you start your work day. Format 9:00am"
	$global:StartOfShift = Read-Host -Prompt $PT
	$global:StartOfShift = [datetime] $global:StartOfShift
	
	$PT = "Enter when you end your work day. Format 6:00pm"
	$global:EndOfShift = Read-Host -Prompt $PT
	$global:EndOfShift = [datetime] $global:EndOfShift

} 

#force disconnect
function Set-EXODisconnect 
{
	Disconnect-ExchangeOnline -Confirm:$false
}

#reusable yesno prompt
function YesNo($Prompt) 
{
	$PT = $Prompt + " [Yes] No"
	$YN = Read-Host -Prompt $PT
    if($YN -eq "" -or $YN -eq "Yes"  -or  $YN -eq "YES"  -or  $YN -eq "Y"  -or  $YN -eq "y"){ #if user doesn't input anything use default
		return "Yes"
	}	
	return 
}

#install the module
function Get-EXOM
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

#set useralias

function Show-Menu 
{
    param (
        [string]$Title = 'Email Out of Office Automation for ${global:UserAlias}'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    
    Write-Host "1: Press '1' Enable Scheduled Auto Reply and Quit"
    Write-Host "2: Press '2' To set your email suffix"
	Write-Host "3: Press '3' To set your office hours"
    Write-Host "4: Press '4' To set your work days"
	Write-Host "5: Press '5' To set the Auto Reply state to Enable:Disable:Scheduled"
	Write-Host "6: Press '6' "
    Write-Host "Q: Press 'Q' to quit."
}

do
{
	
	Show-Menu
	$S = Read-Host "Please make a selection"
	switch ($S)
	{
		'1'
		{
			$global:UserAlias = Get-Alias
			#connect to exchange online
			Get-EXOConnection
			#get the users work days and start/end of shift time
			#if hardcoded at start of file this will be silent
			Get-Schedule
			
			Write-Host "Current account is " -NoNewline
			Write-Host "${global:UserAlias}" -ForegroundColor Blue

			#set to scheduled
			Set-ARCState '3' 

			#set start and end times
			Set-ARCTimes

			#get current configuration, get-arc saves local file
			$MailboxARC = Get-ARC

			Write-Host "Auto Reply state is currently Set to" $MailboxARC.AutoReplyState
			Write-Host "Auto Reply will start at" $MailboxARC.StartTime
			Write-Host "Auto Reply will end at" $MailboxARC.EndTime

			Set-EXODisconnect
			$selection = 'q'
		}
		'2'
		{
			$UserAliasSuffix = Get-Suffix
		}
		'3'
		{
			Get-ShiftTime
			Get-ShiftTime
			Set-ARCTimes
			$MailboxARC = Get-ARC
		}
		'4'
		{
			$WD = ''
			$WD = Get-WD
		}
		'5'
		{
			Get-EXOConnection
			Set-ARCState
			Set-EXODisconnect
		}
		'6'
		{
		}
	}
	pause
}
until ($selection -eq 'q')

#ensure disconnection
Set-EXODisconnect
