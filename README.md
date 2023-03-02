These functions are used to set your OOF message based on your office hours
Currently you can hardcode the email/domain name suffix and your office hours

TO ADD:

DONE - days of the week selector
days of the week + custom hours per day

How To Use this:

run the setOOFMessage.ps1



Function list:

get-Alias #get username and suffix default is username from local machine plus @microsoft.com

ConnectAlias2EXO # connect to Exchange online with local username + suffix, if using AAD, you should be already connected

get-ARC	#check for local config file in same directory as script, if no local file, get auto reply config, use current message, save to file 

Set-ARCSTATEScheduled #set to start and end times, set auto reply to scheduled setting

set-ARCTimes  #sets start and end time for oof message

set-ARCMessage($IOE,$message) # sets your messsage 