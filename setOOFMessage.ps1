. "./OOFFunctions.ps1" #include all the fancy functions
#Get-Alias executed when variables are inited #Get username and suffix default is username from local machine plus @microsoft.com
#ConnectAlias2EXO # connect to Exchange online
#Get-ARC	#check for local config, if none Get auto reply config, use current message

Set-ARCState #Set to state to scheduled
Set-ARCTimes    #Set start and end times
Set-ARCmessagefile

#Set-ARCMessage Both 'message here'
Set-ARCFile #save auto reply config to json file
DisconnectEXO