$global:StartOfShift =  [datetime]"9:00am" #$null
$global:EndOfShift = [datetime]"6:00pm" #$null 
$global:UserAliasSuffix = "@microsoft.com"
$global:UserAlias = Get-Alias
$WD = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')

#Get current username from local user foldername
function Get-UsernameFromWindows 
{
	$CurrentUser = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
    #Write-Host "CurrentUser is " -NoNewline
	#Write-Host "$CurrentUser" -ForegroundColor Blue
	Return $CurrentUser
}

#Get alias from userfolder, if this fails, exo connection will prompt for creds
function Get-Alias 
{
	if($global:UserAliasSuffix -eq "" -or $null -eq $global:UserAliasSuffix)
	{
		$global:UserAliasSuffix = Get-Suffix
	}
	$CurrentUser = Get-UsernameFromWindows
    Return (-join($CurrentUser,$global:UserAliasSuffix))
	#Write-Host "Current account is " -NoNewline
	#Write-Host "${global:UserAlias}" -ForegroundColor Blue
}
function Get-Suffix 
{
    Write-Host "Current suffix is ${UserAliasSuffix}"
	$PT = "What email suffix would you like to use? Format @microsoft.com"
	$global:UserAliasSuffix = Read-Host -Prompt $PT
	Return $global:UserAliasSuffix
}

#connect to exchange online
function Get-EXOConnection 
{
	$global:UserAlias = Get-Alias
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
	Return $AFP
}

#write current config to file
function Set-ARCFile
{
	$ARCFilePath = Get-ARCFilePath
	#This is only called when writing the file, no need to check to overwrite
	Get-ARC | ConvertTo-Json -depth 100 | Set-Content $ARCFilePath
}

#Get current config from online
#save to local file
function Get-ARC
{
	Return Get-MailboxAutoReplyConfiguration -Identity $global:UserAlias
}

#read the locally stored file
function Get-ARCFile 
{
	$ARCFilePath = Get-ARCFilePath
	#Write-Host $ARCFilePath
    Return Get-Content $ARCFilePath -raw | ConvertFrom-Json 
}


#Set auto reply to scheduled/endabled/disabled
function Set-ARCState($S)
{
	#get current configuration
	$MailboxARC = Get-ARC
	#Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState

	if(!$S)
	{
		Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState
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
			#have to set the times again too
			Set-ARCTimes
		}
	}	
	#Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState
	Set-ARCFile
}

#Set auto reply start and end times
function Set-ARCTimes
{
	##Gets office hours, if not hardcoded at the start of this file, ask user for input
	if($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift){Get-ShiftTime}
	
	$daysToAdd = 0
	#how many days till next day of work
	$daysToAdd = Get-NextWorkDay

	#convert daily time to todays time
	$hours = Get-Date $global:StartOfShift
	# Write-Host $global:StartOfShift
	# Write-Host $hours
	$global:StartOfShift = [datetime] (Get-Date).Date.AddHours($hours.Hour)

	#add the number of days till next shift to the time for when the OOF message should end, aka the START of your next shift
	$EndOfAR = $global:StartOfShift.adddays($daysToAdd)

	#convert daily time to todays time, round to hour
	$hours = Get-Date $global:EndOfShift
	$global:EndOfShift = [datetime] (Get-Date).Date.AddHours($hours.Hour)

	#Write-Host "Current Online start:" $MailboxARC.StartTime "`nCurrent Online will End: " $MailboxARC.EndTime
	#Write-Host "Live Config start:" $global:EndOfShift "`nLive Config will End: " $global:StartOfShift

	#Set start and end time for scheduled auto reply
	Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -StartTime $global:EndOfShift -EndTime $EndOfAR
	
	#Write Current Config to file
	Set-ARCFILE
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
	$TempARC = Get-ARC
	$text = $TempARC.ExternalMessage.tostring()
	$text | Out-File -FilePath $ARCMessageFile
	Write-Host "Message file saved as $ARCMessageFile"
}

#Returns the number of days till next work day
function Get-NextWorkDay
{
	if($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift){Get-ShiftTime}

	$duringshift = 0
	$CTime =  Get-Date #-Format "MM/dd/yyyy HH:mm"
	$CTime =  [datetime] $CTime

	#what days of the week do you work hard code it if you dont wanna be asked
	$WD = Get-WD

	if(!($CTime.DayOfWeek -in $WD))
	{
		$i = 0
		#Write-Host "You are not working today" $CTime.DayOfWeek
		while(!($CTime.DayOfWeek -in $WD))
		{
			$i += 1
			#Write-Host $CTime.DayOfWeek -ForegroundColor Red -NoNewline 
			#Write-Host " is not currently a work day [" -NoNewline
			#Write-Host  $WD -NoNewline -ForegroundColor Blue
			#Write-Host "]"
			$CTime = $CTime.adddays(1)		
		}
		$duringshift = $i
		#Write-Host $CTime.DayOfWeek
		#Write-Host $global:StartOfShift.TimeOfDay
		Write-Host (-join("The start of the next workday is ",$CTime.DayOfWeek," ",$global:StartOfShift.TimeOfDay))
	}
	else
	{
		#Write-Host "You are working today" $CTime.DayOfWeek
		#When is next work day?
		#Do I work Tomorrow?
		$CTime = $CTime.adddays(1)
		$i = 1
		while(!($CTime.DayOfWeek -in $WD))
		{
			$i += 1
			$CTime = $CTime.adddays(1)	
			#Write-Host $CTime.DayOfWeek -ForegroundColor Red -NoNewline 
			#Write-Host " is not currently a work day [" -NoNewline
			#Write-Host  $WD -NoNewline -ForegroundColor Blue
			#Write-Host "]"
		}
		if($i -gt 1)
		{
			#Write-Host $CTime.DayOfWeek
			#Write-Host $global:StartOfShift.TimeOfDay
			Write-Host (-join("The start of the next workday is ",$CTime.DayOfWeek," ",$global:StartOfShift.TimeOfDay))
			$global:StartOfShift = $CTime - $CTime.TimeOfDay + $global:StartOfShift.TimeOfDay
			Return $i
		}

		# if $i is not > 1 then next work day is today, or tomorrow
		$CTime =  Get-Date #reset CTime to current
		$CTime =  [datetime] $CTime
		if($CTime -lt $global:StartOfShift)
		{ 
			#Write-Host "${CuTime} Currently Before Shift" ### use todays start and end times, rerun during shift to Set for overnight oof
			$duringshift = 0
		}
		elseif($CTime -gt $global:EndOfShift)
		{
			#Write-Host "${CuTime} Currently After Shift"### use tomorrows start time and todays end time
			$duringshift = 1
		}
		elseif($CTime -le $global:EndOfShift -And $CTime -ge $global:StartOfShift)
		{
			#Write-Host "${CuTime} Currently During Shift" ### use tomorrows start time and todays end time
			$duringshift = 1
		}
		else 
		{
			Write-Host "Twilight Zone"
		}
	}
	Return $duringshift
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
		$S = Read-Host -Prompt "Which of the following matches your weekly work schedule`n1. 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'`n2. 'Monday', 'Tuesday', 'Wednesday', 'Sunday'`n3. 'Wednesday', 'Saturday','Sunday','Monday'`nChoice "
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
				$WD = @('Wednesday','Saturday','Sunday','Monday')
			}
		}
		Write-Host $WD
	}
	Return $WD
}

#what time do you start and end your shift
function Get-ShiftTime
{
	$TempT = Read-Host -Prompt "Enter when you start your work day. Format 9:00am"
	$global:StartOfShift = [datetime] $TempT
	
	$TempT = Read-Host -Prompt "Enter when you end your work day. Format 6:00pm"
	$global:EndOfShift = [datetime] $TempT
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
		Return "Yes"
	}	
	Return 
}

#install the module
function Get-EXOM
{
	if ((Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
		#Write-Host "ExchangeOnlineManagement exists, not installing`n"
        #no output if it is installed, less chatty
		Update-Module -Name ExchangeOnlineManagement
        Return
	} 
	else {
		Write-Host "ExchangeOnlineManagement does not exist, installing`n"
		Install-Module -Name ExchangeOnlineManagement -force
	}
	Return
}

#menu here
function Show-Menu 
{
    param (
        [string]$Title = "Email Out of Office Automation"
		
    )
	$alias = Get-Alias
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "Current account is " -NoNewline
	Write-Host "$alias" -ForegroundColor Blue
    Write-Host "1: Press '1' Enable Scheduled Auto Reply and Quit"
    Write-Host "2: Press '2' To display the currect Auto Reply Configuration"
	Write-Host "3: Press '3' To set your office hours"
    Write-Host "4: Press '4' To set your work days"
	Write-Host "5: Press '5' To set the Auto Reply state to Enable:Disable:Scheduled"
	Write-Host "6: Press '6' Save Auto Reply Message to Local HTML File"
    Write-Host "Q: Press 'Q' to quit."
}



##################### here is where the magic starts ####################
Get-EXOConnection
#### get connected once, this assumes the suffix is correctly hardcoded, if not everything breaks lol
do
{
	Show-Menu
	$S = Read-Host "Please make a selection"
	switch ($S)
	{
		'1'
		{
			#get the users work days and start/end of shift time
			#if hardcoded at start of file this will be silent
			$waste = Get-NextWorkDay
			
			#Write-Host "Current account is " -NoNewline
			#Write-Host "${global:UserAlias}" -ForegroundColor Blue

			#set to scheduled
			Set-ARCState '3' 

			#set start and end times
			Set-ARCTimes

			#save current config to local file why?
			Set-ARCFILE
			$TempARC = Get-ARC
			Write-Host "Auto Reply state is currently Set to" $TempARC.AutoReplyState
			Write-Host "Auto Reply will start at" $TempARC.StartTime
			Write-Host "Auto Reply will end at" $TempARC.EndTime

			#quitting time
			$S = 'q'
		}
		'2'
		{
			Get-ARC
		}
		'3'
		{
			Get-ShiftTime
			Set-ARCTimes
		}
		'4'
		{
			$WD = ''
			$WD = Get-WD
			$waste = Get-NextWorkDay
			Set-ARCTimes
		}
		'5'
		{
			Set-ARCState
		}
		'6'
		{
			Set-ARCmessagefile
		}
	}
	pause
}
until ($S -eq 'q')

#ensure disconnection
Set-EXODisconnect
