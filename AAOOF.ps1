param([string]$InputParm)
$global:StartOfShift = $null
$global:EndOfShift = $null
$global:WorkDays = @()
#I really dont like that the first 4 lines of this script must be in this order, as we store the user's values here after this is run the first time
$global:UserAliasSuffix = "@microsoft.com"
$global:UserAlias = $null

#Get alias from userfolder, if this fails, exo connection will prompt for creds
function Get-UserAlias {
	if ($global:UserAliasSuffix -eq "" -or $null -eq $global:UserAliasSuffix) {
		$global:UserAliasSuffix = Get-Suffix
	}

	$CurrentUser = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
    
	#Write-Host "Current account is " -NoNewline
	#Write-Host "$CurrentUser" -ForegroundColor Blue
	$CurrentUser = ( -join ($CurrentUser, $global:UserAliasSuffix))
	Return $CurrentUser 
}

function Get-Suffix {
	Write-Host "Current suffix is $global:UserAliasSuffix"
	$PT = "What email suffix would you like to use? Deafult  [@microsoft.com]"
	try {
		$global:UserAliasSuffix = Read-Host -Prompt $PT
	}
	catch {
		Write-Host "Bad Input $global:UserAliasSuffix"
	}
	Return $global:UserAliasSuffix
}

function Get-ARCFilePath {
	$AFP = Get-Location #store local copy in same folder as script
	$AFP = ( -join ($AFP.tostring(), '\', 'AutoReplyConfig.json'))
	Return $AFP
}

#write current config to file
function Set-ARCFile {
	$ARCFilePath = Get-ARCFilePath
	#This is only called when writing the file, no need to check to overwrite
	Get-ARC | ConvertTo-Json -depth 100 | Set-Content $ARCFilePath
} 

#Get current config from online
#save to local file
function Get-ARC {
	Return Get-MailboxAutoReplyConfiguration -Identity $global:UserAlias
}

#read the locally stored file
function Get-ARCFile {
	$ARCFilePath = Get-ARCFilePath
	#Write-Host $ARCFilePath
	Return Get-Content $ARCFilePath -raw | ConvertFrom-Json 
}

#Set auto reply to scheduled/endabled/disabled
function Set-ARCState($S) {
	#get current configuration
	$MailboxARC = Get-ARC
	#Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState

	if (!$S) {
		Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState
		$S = Read-Host -Prompt "What mode should Auto Reply be set to?`n1. Enabled`n2. Disabled`n3. Scheduled`nChoice "
	}
	switch ($S) {
		'1' {
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -AutoReplyState "Enabled"
		}
		'2' {
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -AutoReplyState "Disabled"
		}
		'3' {
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -AutoReplyState "Scheduled"
			#have to set the times again too
			Set-ARCTimes
		}
	}	
	#Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState
	Set-ARCFile
}

#Set auto reply start and end times
function Set-ARCTimes {
	##Gets office hours, if not hardcoded at the start of this file, ask user for input
	if ($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift) { Get-ShiftTime }
	if ($null -eq $global:WorkDays) { Get-WorkDaysOfTheWeek }
	
	$daysToAdd = 0
	#how many days till next day of work
	$daysToAdd = Get-NextWorkDay
	

	#convert daily time to todays time
	$hours = Get-Date $global:StartOfShift
	# Write-Host $global:StartOfShift
	# Write-Host $hours
	$global:StartOfShift = [datetime] (Get-Date).Date.AddHours($hours.Hour)

	#add the number of days till next shift to the time for when the OOF message should end, aka the START of your next shift
	$global:StartOfShift = $global:StartOfShift.adddays($daysToAdd)

	#convert daily time to todays time, round to hour
	$hours = Get-Date $global:EndOfShift
	$global:EndOfShift = [datetime] (Get-Date).Date.AddHours($hours.Hour)
	
	#Write-Host "Current Online start:" $MailboxARC.StartTime "`nCurrent Online will End: " $MailboxARC.EndTime
	#Write-Host "Live Config start:" $global:EndOfShift "`nLive Config will End: " $global:StartOfShift

	#Set start and end time for scheduled auto reply
	if ($daysToAdd -eq 0) { $global:EndOfShift = $global:EndOfShift.adddays(-1) }#shift has not started, move start of OOF back a day
	Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -StartTime $global:EndOfShift -EndTime $global:StartOfShift
	
	#Write Current Config to file
	Set-ARCFILE
}

#Set auto reply message
function Set-ARCMessage($message, $IOE) {
	switch -Regex ($IOE) {
		"Internal" {
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -InternalMessage $message 
		}
		"1" {
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -InternalMessage $message 
		}
		"I" {
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -InternalMessage $message 
		}
		"External" {
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -ExternalMessage $message  
		}
		"2" {
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -ExternalMessage $message  
		}
		"E" {
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -ExternalMessage $message  
		}
		default {
			Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -ExternalMessage $message -InternalMessage $message 
		}
	}
}

# get the message file, need to deliniate the separate messages, either within a single file or in multiple files for simpler editing
function Get-ARCmessagefile {
	$ARCMessageFile = Get-Location #store local copy in same folder as script
	$ARCMessageFile = ( -join ($ARCMessageFile.tostring(), '\', 'message.html'))
	return $ARCMessageFile
}
#save online message to html file
function Set-ARCmessagefile {
	$ARCMessageFile = Get-Location #store local copy in same folder as script
	$ARCMessageFile = ( -join ($ARCMessageFile.tostring(), '\', 'message.html'))
	#Write-Host $ARCMessageFile
	$TempARC = Get-ARC
	$text = $TempARC.ExternalMessage.tostring()
	$text | Out-File -FilePath $ARCMessageFile
	Write-Host "Message file saved as $ARCMessageFile"
}

#Returns the number of days till next work day
function Get-NextWorkDay {
	if ($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift) { Get-ShiftTime }
	if (!$global:WorkDays) { Get-WorkDaysOfTheWeek }

	$duringshift = 0
	$CTime = Get-Date #-Format "MM/dd/yyyy HH:mm"
	$CTime = [datetime] $CTime

	if (!($CTime.DayOfWeek -in $global:WorkDays)) {
		$i = 0
		#Write-Host "You are not working today" $CTime.DayOfWeek
		while (!($CTime.DayOfWeek -in $global:WorkDays)) {
			$i += 1
			#Write-Host $CTime.DayOfWeek -ForegroundColor Red -NoNewline 
			#Write-Host " is not currently a work day [" -NoNewline
			#Write-Host  $global:WorkDays -NoNewline -ForegroundColor Blue
			#Write-Host "]"
			$CTime = $CTime.adddays(1)		
		}
		$duringshift = $i
		#Write-Host $CTime.DayOfWeek
		#Write-Host $global:StartOfShift.TimeOfDay
		#Write-Host (-join("The start of the next workday is ",$CTime.DayOfWeek," ",$global:StartOfShift.TimeOfDay))
	}
	else {
		#Write-Host "You are working today" $CTime.DayOfWeek
		#When is next work day?
		#Do I work Tomorrow?
		$CTime = $CTime.adddays(1)
		$i = 1
		while (!($CTime.DayOfWeek -in $global:WorkDays)) {
			$i += 1
			$CTime = $CTime.adddays(1)	
			#Write-Host $CTime.DayOfWeek -ForegroundColor Red -NoNewline 
			#Write-Host " is not currently a work day [" -NoNewline
			#Write-Host  $global:WorkDays -NoNewline -ForegroundColor Blue
			#Write-Host "]"
		}
		if ($i -gt 1) {
			#Write-Host $CTime.DayOfWeek
			#Write-Host $global:StartOfShift.TimeOfDay
			#Write-Host (-join("The start of the next workday is ",$CTime.DayOfWeek," ",$global:StartOfShift.TimeOfDay))
			$global:StartOfShift = $CTime - $CTime.TimeOfDay + $global:StartOfShift.TimeOfDay
			Return $i
		}

		# if $i is not > 1 then next work day is today, or tomorrow
		$CTime = Get-Date #reset CTime to current
		$CTime = [datetime] $CTime
		if ($CTime -lt $global:StartOfShift) { 
			#Write-Host "${CuTime} Currently Before Shift" ### use todays start and end times, rerun during shift to Set for overnight oof
			$duringshift = 0
			##### wait till shift starts +1 minute and run script again with default option?
		}
		elseif ($CTime -gt $global:EndOfShift) {
			#Write-Host "${CuTime} Currently After Shift"### use tomorrows start time and todays end time
			$duringshift = 1
		}
		elseif ($CTime -le $global:EndOfShift -And $CTime -ge $global:StartOfShift) {
			#Write-Host "${CuTime} Currently During Shift" ### use tomorrows start time and todays end time
			$duringshift = 1
		}
		else {
			Write-Host "Twilight Zone"
		}
	}
	Return $duringshift
}

function Get-WorkDaysOfTheWeek {   
	### this is a function to either Set an array of days of the week that you work by uncommenting or configuring your own line below
	### These are the days of the week that you work
	### Common examples can be uncommented
	### Or edit the default

	### 4 Days Sunday - Wednesday 
	#$global:WorkDays = @('Monday', 'Tuesday', 'Wednesday', 'Sunday')

	### 4 Days Wednesday - Saturday
	#$global:WorkDays = @('Wednesday', 'Thursday', 'Friday', 'Saturday')

	### Twitter Employee Working 7 days wont need this script
	#$global:WorkDays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')

	### no wednesdays or thursdays testing
	#$global:WorkDays = @('Monday', 'Tuesday', 'Friday', 'Saturday', 'Sunday')

	### Standard Monday - Friday
	#$global:WorkDays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')

	if (!$global:WorkDays) {
		
		#Clear-Host
		Write-Host "`n================ What days of the Week do you Work? ================"
		Write-Host "Which of the following matches your weekly work schedule"
		Write-Host "1. 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday"
		Write-Host "2. 'Monday', 'Tuesday', 'Saturday', 'Sunday'"
		Write-Host "3. 'Thursday','Friday','Saturday','Sunday'"
		$S = Read-Host -Prompt "Choice [1]"
		switch ($S) {
			'1' {
				$global:WorkDays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
			}
			'2' {
				$global:WorkDays = @('Monday', 'Tuesday', 'Saturday', 'Sunday')
			}
			'3' {
				$global:WorkDays = @('Thursday', 'Friday', 'Saturday', 'Sunday')
			}
			default {
				$global:WorkDays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
			}

		}
		#Write-Host $global:WorkDays
	}
	Return $global:WorkDays
}

function Set-WorkDaysToFile {
	$FP = Get-Location
	$FP = ( -join ($FP.tostring(), '\', 'AAOOF.ps1'))
	$content = Get-Content -Path $FP
	$Temp = "`$global:WorkDays = @("
	foreach ($day in $global:WorkDays) {
		$Temp += "`'" + $day + "`',"
	}
	$Temp = $Temp.Substring(0, $Temp.Length - 1)
	$Temp += ")"
	$content[3] = $Temp #overwrite line 3
	#Write-Host $Temp
	Set-Content $FP $content #write to file
}

#what time do you start and end your shift
function Get-ShiftTime {
	if ($null -eq $global:StartOfShift) {
		do {
			try {
				$TempT = Read-Host -Prompt "Enter when you start your work day. Default [9:00am]"
				if ($TempT -eq "") {
					$global:StartOfShift = [datetime] "9:00 am"
				}
				else {
					$global:StartOfShift = [datetime] $TempT
				}
				$valid = $true
			}
			catch {
				Write-Host "Invalid input: $TempT"
				$valid = $false
			}
		}
		until ($valid)
		$valid = $false
	}
	if ($null -eq $global:EndOfShift) {
		do {
			try {
				$TempT = Read-Host -Prompt "Enter when you end your work day. Default [6:00pm]"
				if ($TempT -eq "") {
					$global:EndOfShift = [datetime] "6:00 pm"
				}
				else {
					$global:EndOfShift = [datetime] $TempT
				}
				$valid = $true
			}
			catch {
				Write-Host "Invalid input: $TempT"
				$valid = $false
			}
		}
		until ($valid)
		$valid = $false
	}
	# $Temp +=  ""
	# $Temp += $global:StartOfShift.TimeOfDay 
	# $Temp += " till "
	# $Temp += $global:EndOfShift.TimeOfDay
	# Write-Host $Temp
}
#currently writes to test file not actual ps1
function Set-WorkTimesToFile {
	$FP = Get-Location
	$FP = ( -join ($FP.tostring(), '\', 'AAOOF.ps1'))
	$content = Get-Content -Path $FP
	$content[1] = "`$global:StartOfShift = [datetime] `"" + $global:StartOfShift.TimeOfDay + "`""
	$content[2] = "`$global:EndOfShift = [datetime] `"" + $global:EndOfShift.TimeOfDay + "`""
	#Write-Host $content[1]
	#Write-Host	$content[2]
	Set-Content $FP $content
}

##create scheduled task function
function Set-DailyScriptTask {
	
	$FP = Get-Location
	$FP = ( -join ($FP.tostring(), '\', 'AAOOF.ps1'))

	#the scheduled task
	$taskname = "AAOOF"
	

	$Task = Get-ScheduledTask -TaskPath '\' -TaskName $taskname
	$Task.Actions[0].Arguments = "${FP} 1"

	Set-ScheduledTask -InputObject $Task | Out-Null

	# if task dne
	# $action = New-ScheduledTaskAction -Execute 'powershell.exe'
	# $date = Get-Date -Date (Get-Date).Date
	# $TriggerTime = $global:StartOfShift.TimeOfDay
	# $TriggerTime = $date.AddMinutes(15) + $TriggerTime
	
	# # Define the trigger for the scheduled task 15 minutes after start of shift
	# $trigger = New-ScheduledTaskTrigger -Daily -At $TriggerTime
	# Register-ScheduledTask -TaskName ${taskname} -Trigger $trigger -Action $action -RunLevel Highest
}

#get date for return to work, this sets autoreply to start at end of shift today and end on start of shift on date entered
function Get-VacationDate ($TempT) {
	if (!$TempT) {
		$TempT = Read-Host -Prompt "Enter the next date of work when you return from vacation. Format YYYY/MM/DD"
		#Write-Host "Time for end of autoreply is "$global:StartOfShift.TimeOfDay
		$TempT = [datetime] $TempT
	}
	else {
		$TempT = [datetime] $TempT
	}
	#Write-Host "Date for end of autoreply is " $Tempt
	$global:StartOfShift = $Tempt + $global:StartOfShift.TimeOfDay
	Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -StartTime $global:EndOfShift -EndTime $global:StartOfShift
}

#connect to exchange online
function Get-EXOConnection {
	if ($null -eq $global:UserAlias) { $global:UserAlias = Get-UserAlias }
	Get-EXOM #is EXO module installed
	#Write-Host "Current account is " -NoNewline
	#Write-Host "${global:UserAlias}" -ForegroundColor Blue
	#Write-Host "Connecting to your Outlook Account with alias $global:UserAlias " 
	
	# Get the current PowerShell connection
	$session = Get-ConnectionInformation
	# Check if there is an active Exchange Online PowerShell v3 connection
	if ($null -ne $session) {
		$exchangeSession = $session | Where-Object { $_.Name -like "ExchangeOnline_*" }
		if ($null -ne $exchangeSession) {
			Write-Host "An active Exchange Online PowerShell v3 connection exists."
		}
		else {
			Write-Host "No active Exchange Online PowerShell v3 connection exists."
			Connect-ExchangeOnline -UserPrincipalName $global:UserAlias
		}
	}
	else {
		#Write-Host "No PowerShell session exists." + $global:UserAlias + "`n"
		Connect-ExchangeOnline -UserPrincipalName $global:UserAlias
	}
	#Write-Host "Done Connecting"
	Clear-Host
}

#force disconnect
function Set-EXODisconnect {
	Disconnect-ExchangeOnline -Confirm:$false
	## close msedge window when done
}

#install the module
function Get-EXOM {
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
}

#menu here
function Show-Menu {
	Clear-Host
	Write-Host "================ Email Out of Office Automation ================"
	Write-Host "Current account is " -NoNewline
	Write-Host "$global:UserAlias" -ForegroundColor Blue
	Write-Host "1: Press '1' Enable Scheduled Auto Reply and Quit"
	Write-Host "2: Press '2' To set an end date for a extended out of office message`n`n"
	Write-Host "================ Configure the Script Defaults ================"
	Write-Host "3: Press '3' To set your office hours and save to script"
	Write-Host "4: Press '4' To set your work days and save to script`n`n"
	Write-Host "================ Configure the Auto Reply Settings ================"
	Write-Host "5: Press '5' To set the Auto Reply state to Enable:Disable:Scheduled"
	Write-Host "6: Press '6' " -NoNewline
	Write-Host "REQUIRES ADMIN" -ForegroundColor Red -NoNewline
	Write-Host "Schedule Task to run the 'AAOOF.ps1 1' 15 minutes after the start of your shift daily`n`n"
	Write-Host "================ Configure the Auto Reply Message ================"
	Write-Host "9: Press '9' Save the current Auto Reply Message to File"
	Write-Host "0: Press '0' Load an Auto Reply Message to File`n`n"
	Write-Host "Q: Press 'Q' to quit."
}

do {
	if ($InputParm -ne 'z') {
		if ($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift) {
			Get-ShiftTime
			Set-WorkTimesToFile
		}
		if (!$global:WorkDays) {
			$global:WorkDays = Get-WorkDaysOfTheWeek
			Set-WorkDaysToFile			
		}
		$S = 'q'
	}
	if ($InputParm -as [datetime]) {
		#connect
		Get-EXOConnection
		Get-VacationDate $InputParm
		$TempARC = Get-ARC
		Write-Host "Auto Reply state is currently Set to" $TempARC.AutoReplyState
		Write-Host "Auto Reply will start at" $TempARC.StartTime
		Write-Host "Auto Reply will end at" $TempARC.EndTime
		$S = 'q'
	}
	if (!$InputParm) {
		Show-Menu
		$S = Read-Host "Please make a selection"
	}
	else {
		$S = $InputParm
	}

	switch ($S) {
		'1' {
			# run daily option after start of shift
			#Write-Host "Current account is " -NoNewline
			#Write-Host "${global:UserAlias}" -ForegroundColor Blue

			
			#connect
			Get-EXOConnection

			#set to scheduled, scheduled also calls set-arctimes
			Set-ARCState '3' 

			$TempARC = Get-ARC
			Write-Host "Auto Reply state is currently Set to" $TempARC.AutoReplyState
			Write-Host "Auto Reply will start at" $TempARC.StartTime
			Write-Host "Auto Reply will end at" $TempARC.EndTime

			#quitting time
			$S = 'q'
		}
		'2' {
			#connect
			Get-EXOConnection

			#set the return date for vacation oof, use vacation oof message if present otherwise use default		
			Get-VacationDate
			$TempARC = Get-ARC
			Write-Host "Auto Reply state is currently Set to" $TempARC.AutoReplyState
			Write-Host "Auto Reply will start at" $TempARC.StartTime
			Write-Host "Auto Reply will end at" $TempARC.EndTime
			$InputParm = $null
		}
		'3' {
			#change the start and end of shift times and save to script
			$global:StartOfShift = $null
			$global:EndOfShift = $null
			Get-ShiftTime
			Set-ARCTimes
			Set-WorkTimesToFile
			$InputParm = $null
		}
		'4' {
			#change what days of the week you work and save to script
			$global:WorkDays = @()
			$global:WorkDays = Get-WorkDaysOfTheWeek
			Set-WorkDaysToFile
			$InputParm = $null
		}
		'5' {
			#connect
			Get-EXOConnection

			#set endable disabled or scheduled
			Set-ARCState
			$InputParm = $null
		}
		'6' {
			#create schedule windows task
			Set-DailyScriptTask
			$InputParm = $null
		}
		'9' {
			#connect
			Get-EXOConnection

			#save the current message to the default message.html file
			Set-ARCmessagefile
			#Set-WorkTimesToFile
			#Set-WorkDaysToFile	
			#Set-ARCState '3'
			$InputParm = $null
		}
		'Z' {
			### hidden reset option, allows me to commit the script with the default null values and not have to manually update the main branch 
			#set null defaults to script
			$FP = Get-Location
			$FP = ( -join ($FP.tostring(), '\', 'AAOOF.ps1'))
			$content = Get-Content -Path $FP
			$content[1] = "`$global:StartOfShift = `$null"
			$content[2] = "`$global:EndOfShift = `$null"
			$content[3] = "`$global:WorkDays = @()"
			Set-Content $FP $content
			$InputParm = $null
			$S = 'q'
		}
		'Q' {
			try {
				#ensure disconnection
				Set-EXODisconnect	
			}
			catch {
				Write-Host "Did not disconnect"
			}
		}
	}
}
until ($S -eq 'q')

try {
	#ensure disconnection
	Set-EXODisconnect	
}
catch {
	Write-Host "Did not disconnect"
}  
