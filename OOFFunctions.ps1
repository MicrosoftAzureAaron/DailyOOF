#Get current username from local user foldername
function CurrentUserNamefromWindows 
{
	$CurrentUser = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
    Write-Host "CurrentUser is " -NoNewline
	Write-Host "$CurrentUser" -ForegroundColor Blue
	return $CurrentUser
}

#Get alias from userfolder, if this fails, it will prompt for creds
function Get-Alias 
{
	$CurrentUser = CurrentUserNamefromWindows
    $UA = (-join($CurrentUser,$UserAliasSuffix))
    Write-Host "UserAlias is " -NoNewline
	Write-Host "$UA" -ForegroundColor Blue
    #Write-Host "UserAliasSuffix is " -NoNewline
	#Write-Host "$UserAliasSuffix" -ForegroundColor Blue
	return $UA
}

#connect to exchange online
function ConnectAlias2EXO 
{
	InstallEXOM #is EXO module installed
	Write-Host "Connecting to your Outlook Account with alias $UserAlias" 
	Connect-ExchangeOnline -UserPrincipalName $UserAlias
	Write-Host "Done Connecting"
}

#write current config to file, warn about overwrite
function Set-ARCFile
{
	if($MessageFilePath) 
	{
        ###file exists do you want to overwrite
		$Q = YesNo "Auto Reply config file already exists, over write ${MessageFilePath}?"
	}
	else
	{
		###write file
		$Q = YesNo "No local copy found, do you want to save a local copy on ${MessageFilePath}?"
	}
	if($Q -eq "Yes") 
	{
		SaveIt "Auto Reply config is being written to JSON file from current configuration to {$MessageFilePath}"
	}
}

#write the file file from 'memory'
function SaveIt($PT)
{	
	Write-Host "$PT"
	$MailboxARC | ConvertTo-Json -depth 100 | Set-Content $MessageFilePath
}

#Get current config, local file first, otherwise whats online
function Get-ARC 
{
	#add choice load from file or load from online exchange
	#prefers local store over remote
	if(Test-Path $MessageFilePath) 
	{
        Write-Host "ARC File stored locally" $MessageFilePath
        $MailboxARC = Get-ARCFile
		Write-Host "ARC File Loaded from Local File"
		#Write-Host $MailboxARC
	}
    else 
	{
		$MailboxARC = Get-MailboxAutoReplyConfiguration -UserPrincipalName $UserAlias

		$Q = YesNo "Do you want to save current online configuration to a local copy at $MessageFilePath ?"
		if($Q -eq "Yes") 
		{
			$MailboxARC = Get-MailboxAutoReplyConfiguration -UserPrincipalName $UserAlias
			SaveIt "Auto Reply config is being written to JSON file from current Exchange Online connection to $MessageFilePath"
		}
    }
	Return $MailboxARC
}

#read the locally stored file
function Get-ARCFile 
{
	#Write-Host $MessageFilePath
    return Get-Content $MessageFilePath -raw | ConvertFrom-Json 
}

# check to see if file is there
# function FileDNE($FilePath) 
# {
#     return (Get-Item -Path $FilePath -ErrorAction Ignore)
# }

#Set auto reply to scheduled
function Set-ARCState 
{
	#is Reply state disabled or enabled by the user manually instead of scheduled
	if($MailboxARC.AutoReplyState -eq "Disabled" -or $MailboxARC.AutoReplyState -eq "Enabled"){
		Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState
	}
	Set-MailboxAutoReplyConfiguration -Identity $UserAlias -AutoReplyState "Scheduled"
	#Write-Host "Auto Reply state is currently Set to"$MailboxARC.AutoReplyState
	###update json
	#Set-ARCFile  add option to save from menu to local file
}

#Set auto reply start and end times
function Set-ARCTimes
{
	if($null -eq $StartOfShift){$StartOfShift = (GetShiftTime "start")}
	if($null -eq $EndOfShift){$EndOfShift = (GetShiftTime "end")}
	##Gets office hours, if not hardcoded at end of this file, ask user for input
	$daystoadd = 0
	$daystoadd = IsOfficeHours $daystoadd

	#convert daily time to todays time
	$hours = (Get-Date $StartOfShift)
	$StartOfShift = [datetime] (Get-Date).Date.AddHours($hours.Hour)

	#add the number of days till next shift to the time for when the OOF message should end, aka the START of your next shift
	$StartOfShift = $StartOfShift.adddays($daystoadd)

	#convert daily time to todays time
	$hours = Get-Date "$EndOfShift"
	$EndOfShift = [datetime] (Get-Date).Date.AddHours($hours.Hour)

	#Write-Host "From File start:" $MailboxARC.StartTime "`nFrom File will End: " $MailboxARC.EndTime
	#Write-Host "`nLive Config start:" $EndOfShift "`nLive Config will End: " $StartOfShift
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
	$MessageFilePath = Get-Location #store local copy in same folder as script
	$MessageFilePath = (-join($MessageFilePath.tostring(),'\','message.html'))
	#Write-Host $MessageFilePath
	$text = $MailboxARC.ExternalMessage.tostring()
	$text | Out-File -FilePath $MessageFilePath
	Write-Host "Message file saved as $MessageFilePath"
}

function IsOfficeHours($duringshift) 
{
	if($null -eq $StartOfShift){$StartOfShift = (GetShiftTime "start")}
	if($null -eq $EndOfShift){$EndOfShift = (GetShiftTime "end")}

	$duringshift = 0
	$CuTime =  Get-Date #-Format "MM/dd/yyyy HH:mm"
	$CuTime =  [datetime] $CuTime

	#what days of the week do you work hard code it if you dont wanna be asked
	$WorkDays = Workdays_of_week

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

function Workdays_of_week 
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
	$WD = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')

	if(!$WD)
	{
		$Swit = Read-Host -Prompt "Which of the following matches your weekly work schedule`n1. 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'`n2. 'Monday', 'Tuesday', 'Wednesday', 'Sunday'`n3. 'Wednesday', 'Thursday', 'Friday', 'Saturday'`n Choice "
		switch($Swit)
		{
			1{$WD = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')}
			2{$WD = @('Monday', 'Tuesday', 'Wednesday', 'Sunday')}
			3{$WD = @('Wednesday', 'Thursday', 'Friday', 'Saturday')}
		}
	}
	return $WD
}

#what time do you start or end your shift
function GetShiftTime($StartEnd) 
{
	if(Test-Path $MessageFilePath) 
	{
		### ask to use file or online
		$MailboxARC = Get-ARC
		#### check for start and end times in file
		if($StartEnd -eq "start")
		{
			#Write-Host $MailboxARC.EndTime
			$ST = [datetime] $MailboxARC.EndTime
			$TODST = $ST.TimeOfDay
			#Write-Host $ST
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


$UserAliasSuffix = "@microsoft.com"
$UserAlias = Get-Alias #based on user folder name combined with suffix, or hard code it
$MessageFilePath = Get-Location #store local copy in same folder as script
$MessageFilePath = (-join($MessageFilePath.tostring(),'\','AutoReplyConfig.json'))
ConnectAlias2EXO
$MailboxARC = Get-ARC
$StartOfShift = $null #9:00am #GetShiftTime "start" #hard code a time here if you dont want to be asked
$EndOfShift = $null #"6:00pm" #GetShiftTime "end" #hard code a time here if you dont want to be asked