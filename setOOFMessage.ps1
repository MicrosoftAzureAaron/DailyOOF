. "./OOFFunctions.ps1" #include all the fancy functions
#get-Alias executed when variables are inited #get username and suffix default is username from local machine plus @microsoft.com
#ConnectAlias2EXO # connect to Exchange online
#get-ARC	#check for local config, if none get auto reply config, use current message

Set-ARCSTATEScheduled #set to state to scheduled
set-ARCTimes    #set start and end times

set-ARCMessage 'Boy Howdy' Both





DisconnectEXO