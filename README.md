Hardcode your office hours and days of the week that you work at the top of the ps1 file otherwise the set them to null and the script will ask for your input

$global:StartOfShift =  [datetime]"9:00am" #$null
$global:EndOfShift = [datetime]"6:00pm" #$null 
$global:UserAliasSuffix = "@microsoft.com"
$global:UserAlias = Get-Alias
$WD = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')

run with 1 for default options based on hard coded values, run once a day to set your oof message times automatically

./aaoof.ps1 1

run with a Date for vacation mode, run once and leave for vacation OOF will turn off when you get back

./aaoof.ps1 '2023/04/01'

run with nothing for menu

./aaoof.ps1


================ Email Out of Office Automation ================
Current account is aarosanders@microsoft.com
1: Press '1' Enable Scheduled Auto Reply and Quit
2: Press '2' To set an end date for a extended out of office message
3: Press '3' To set your office hours
4: Press '4' To set your work days
5: Press '5' To set the Auto Reply state to Enable:Disable:Scheduled
6: Press '6' Save Auto Reply Message to Local HTML File
Q: Press 'Q' to quit.
Please make a selection: