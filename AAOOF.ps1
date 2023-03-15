$StartOfShift = "9:00am"
$EndOfShift = "6:00pm"
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

#Get alias from userfolder, if this fails, it will prompt for creds
function Get-Alias 
{
	if($UserAliasSuffix -eq "" -or $null -eq $UserAliasSuffix)
	{
		Get-Suffix
	}
	$CurrentUser = Get-UsernameFromWindows
    $UA = (-join($CurrentUser,$UserAliasSuffix))
	return $UA
}
function Get-Suffix 
{
    Write-Host "Current suffix is ${UserAliasSuffix}"
	$PT = "What email suffix would you like to use? Format @microsoft.com"
	$UserAliasSuffix = Read-Host -Prompt $PT
	return $UserAliasSuffix
}

#connect to exchange online
function Connect-Alias2EXO 
{
	Get-EXOM #is EXO module installed
	$UserAlias = Get-Alias
	#Write-Host "Current account is " -NoNewline
	#Write-Host "${UserAlias}" -ForegroundColor Blue
	#Write-Host "Connecting to your Outlook Account with alias $UserAlias" 
	Connect-ExchangeOnline -UserPrincipalName $UserAlias
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
function Get-ARC 
{
	# $ARCFilePath = Get-ARCFilePath
	# #add choice load from file or load from online exchange
	# #prefers local store over remote
	# if(Test-Path $ARCFilePath) 
	# {
    #     Write-Host "ARC File stored locally" $ARCFilePath
    #     $MailboxARC = Get-ARCFile
	# 	Write-Host "ARC File Loaded from Local File"
	# 	#Write-Host $MailboxARC
	# }
    # else 
	# {
	# 	$UserAlias = Get-Alias
	# 	$MailboxARC = Get-MailboxAutoReplyConfiguration -Identity $UserAlias

	# 	# $Q = YesNo "Do you want to save current online configuration to a local copy at $ARCFilePath ?"
	# 	# if($Q -eq "Yes") 
	# 	# {
	# 	# 	$MailboxARC = Get-MailboxAutoReplyConfiguration -Identity $UserAlias
	# 	# 	Set-ARCFile
	# 	# 	# SaveIt "Auto Reply config is being written to JSON file from current Exchange Online connection to $ARCFilePath"
	# 	# }

		
    # }

	$UserAlias = Get-Alias #get alias
	$MailboxARC = Get-MailboxAutoReplyConfiguration -Identity $UserAlias #get arc
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
function Set-ARCState 
{
	$UserAlias = Get-Alias	
	$MailboxARC = Get-ARC
	Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState
	#is Reply state disabled or enabled by the user manually instead of scheduled
	#if($MailboxARC.AutoReplyState -eq "Disabled" -or $MailboxARC.AutoReplyState -eq "Enabled"){
	#	Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState
	#}
	$Swit = Read-Host -Prompt "What mode should Auto Reply be set to?`n1. Enabled`n2. Disabled`n3. Scheduled`nChoice "
	switch($Swit)
	{
		'1'
		{
			Set-MailboxAutoReplyConfiguration -Identity $UserAlias -AutoReplyState "Enabled"
		}
		'2'
		{
			Set-MailboxAutoReplyConfiguration -Identity $UserAlias -AutoReplyState "Disabled"
		}
		'3'
		{
			Set-MailboxAutoReplyConfiguration -Identity $UserAlias -AutoReplyState "Scheduled"
		}
	}	
	#Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState
	###update json
	#Set-ARCFile  add option to save from menu to local file
	$MailboxARC = Get-ARC
}

#Set auto reply start and end times
function Set-ARCTimes
{
	if($null -eq $StartOfShift){$StartOfShift = (Get-ShiftTime "start")}
	if($null -eq $EndOfShift){$EndOfShift = (Get-ShiftTime "end")}
	##Gets office hours, if not hardcoded at end of this file, ask user for input
	$daysToAdd = 0
	$daysToAdd = Get-Schedule

	#convert daily time to todays time
	$hours = (Get-Date $StartOfShift)
	# Write-Host $StartOfShift
	# Write-Host $hours
	$StartOfShift = [datetime] (Get-Date).Date.AddHours($hours.Hour)

	#add the number of days till next shift to the time for when the OOF message should end, aka the START of your next shift
	$StartOfShift = $StartOfShift.adddays($daysToAdd)

	#convert daily time to todays time
	$hours = Get-Date "$EndOfShift"
	$EndOfShift = [datetime] (Get-Date).Date.AddHours($hours.Hour)

	Write-Host "From File start:" $MailboxARC.StartTime "`nFrom File will End: " $MailboxARC.EndTime
	Write-Host "`nLive Config start:" $EndOfShift "`nLive Config will End: " $StartOfShift
	Set-MailboxAutoReplyConfiguration -Identity $UserAlias -StartTime $EndOfShift -EndTime $StartOfShift
	#Set-ARCFile  add option to save from menu to local file
}

#Set auto reply message
function Set-ARCMessage($IOE,$message)
{

	switch -Regex ($IOE)
	{
		"Internal"
		{
			Set-MailboxAutoReplyConfiguration -Identity $UserAlias -InternalMessage $message 
		}
		"External"
		{
			Set-MailboxAutoReplyConfiguration -Identity $UserAlias -ExternalMessage $message  
		}
		"Both"
		{
			Set-MailboxAutoReplyConfiguration -Identity $UserAlias -ExternalMessage $message -InternalMessage $message 
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

function Get-Schedule
{
	if($null -eq $StartOfShift){$StartOfShift = (Get-ShiftTime "start")}
	if($null -eq $EndOfShift){$EndOfShift = (Get-ShiftTime "end")}

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
		#Write-Host $StartOfShift.TimeOfDay
		Write-Host (-join("The start of the next workday is ",$CuTime.DayOfWeek," ",$StartOfShift.TimeOfDay))
	}
	else
	{
		if($CuTime -lt $StartOfShift)
		{ 
			#Write-Host "Currently Before Shift" ### use todays start and end times, rerun during shift to Set for overnight oof
			$duringshift = 0
		}
		elseif($CuTime -gt $StartOfShift)
		{
			#Write-Host "Currently After Shift"### use tomorrows start time and todays end time
			$duringshift = 1
		}
		elseif($CuTime -le $EndOfShift -And $CuTime -ge $StartOfShift)
		{
			#Write-Host "Currently During Shift" ### use tomorrows start time and todays end time
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
		$Swit = Read-Host -Prompt "Which of the following matches your weekly work schedule`n1. 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'`n2. 'Monday', 'Tuesday', 'Wednesday', 'Sunday'`n3. 'Wednesday', 'Thursday', 'Friday', 'Saturday'`nChoice "
		switch($Swit)
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
function Get-ShiftTime($StartEnd) 
{
	$MailboxARC = Get-ARC
	$ARCFilePath = Get-ARCFilePath
	if(Test-Path $ARCFilePath) 
	{
		#### check for start and end times in file
		if($StartEnd -eq "start")
		{
			$ST = [datetime] $MailboxARC.EndTime
			$TODST = $ST.TimeOfDay
			# Write-Host $MailboxARC.EndTime
			# Write-Host $ST
			$PT = "Do you want to used the saved ${StartEnd} of shift time? This is when the OOF message will end ${TODST}"
			#$PT = (-join("Do you want to used the saved $StartEnd of shift time? This is when the OOF message will end ",$ST.TimeOfDay," "))
			if(((YesNo $PT) -eq "Yes"))
			{
				#Write-Host $MailboxARC.StartTime
				#$StartOfShift = $ST
				return $ST
			}
		}

		if($StartEnd -eq "end")
		{
			$ET = [datetime] $MailboxARC.StartTime
			$TODET = $ET.TimeOfDay
			# Write-Host $MailboxARC.StartTime
			# Write-Host $ET
			#$ET = $ET.TimeOfDay
			#Write-Host $ET.TimeOfDay			
			$PT = (-join("Do you want to used the saved $StartEnd of shift time? This is when the OOF message will start ",$TODET))
			if((YesNo $PT -eq "Yes"))
			{
				#Write-Host $MailboxARC.EndTime
				#$EndOfShift = $ET
				return $ET
			}
		}
	}

	$PT = "Enter when you $StartEnd your work day. Format 9:00am"
	$ShiftTime = Read-Host -Prompt $PT
	#Write-Host $ShiftTime
	return [datetime] $ShiftTime
} 

#force disconnect
function DisconnectEXO 
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

function Show-Menu 
{
    param (
        [string]$Title = 'Email Out of Office Automation'
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
	$selection = Read-Host "Please make a selection"
	switch ($selection)
	{
		'1'
		{
			Get-Schedule
			$UserAlias = Get-Alias
			Connect-Alias2EXO
			Set-MailboxAutoReplyConfiguration -Identity $UserAlias -AutoReplyState "Scheduled"
			$MailboxARC = Get-ARC
			Write-Host "Auto Reply state is currently Set to" $MailboxARC.AutoReplyState
			Write-Host "Auto Reply will start at" $MailboxARC.StartTime
			Write-Host "Auto Reply will end at" $MailboxARC.EndTime
			DisconnectEXO
			$selection = 'q'
		}
		'2'
		{
			$UserAliasSuffix = Get-Suffix
		}
		'3'
		{
			$StartOfShift = (Get-ShiftTime "start")
			$EndOfShift = (Get-ShiftTime "end")
		}
		'4'
		{
			$WD = ''
			$WD = Get-WD
		}
		'5'
		{
			Connect-Alias2EXO
			Set-ARCState
			DisconnectEXO
		}
		'6'
		{
		}
	}
	pause
}
until ($selection -eq 'q')
DisconnectEXO