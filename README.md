The first time you run the script it will ask for your input and save the values to your script file

$global:StartOfShift =  $null<br>
$global:EndOfShift = #$null<br>
$global:UserAliasSuffix = "@microsoft.com"<br>
$global:UserAlias = Get-Alias<br>
$WD = $null<br>

After configuring Work Days of the Week and Start and End of your shift, you can run with CLI commands to automate
Run once a day after start of shift

run with 1 for default options based on hard coded values, run once a day to set your oof message times automatically

.\aaoof.ps1 1

run with a Date for vacation mode, run once and leave for vacation OOF will turn off when you get back

.\aaoof.ps1 '2023/04/01'

run with nothing for menu

.\aaoof.ps1

Example Menu Structure<br>
================ Email Out of Office Automation ================<br>
Current account is aarosanders@microsoft.com<br>
1: Press '1' Enable Scheduled Auto Reply and Quit<br>
2: Press '2' To set an end date for a extended out of office message<br>
<br>
<br>
================ Configure the Script Defaults ================<br>
3: Press '3' To set your office hours and save to script<br>
4: Press '4' To set your work days and save to script<br>
<br>
<br>
================ Configure the Auto Reply Message and Settings ================<br>
5: Press '5' To set the Auto Reply state to Enable:Disable:Scheduled<br>
6: Press '6' Save Auto Reply Message to Local HTML File<br>
Q: Press 'Q' to quit.<br>
<br><br>

TO ADD:<br><br>
remove unused functions

set scheduled task to run ps1 once a day<br>

save message to html file for local loading<br><br>
load message from file function<br><br>
Pre saved messages - load from html file<br>
normal oof<br>
vacation oof<br>
sick oof<br>
Holiday oof<br>
