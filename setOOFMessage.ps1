. "./OOFFunctions.ps1" #include all the fancy functions
get-Alias #get username and suffix default is username from local machine plus @microsoftsupport.com
ConnectAlias2EXO # connect to Exchange online
get-ARC	#check for local config, if none get auto reply config, use current message

Set-ARCSTATEScheduled #set to start and end times





DisconnectEXO