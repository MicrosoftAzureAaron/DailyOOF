Hardcode your office hours and days of the week that you work at the top of the ps1 file otherwise the set them to null and the script will ask for your input

$global:StartOfShift =  [datetime]"9:00am" #$null<br>
$global:EndOfShift = [datetime]"6:00pm" #$null<br>
$global:UserAliasSuffix = "@microsoft.com"<br>
$global:UserAlias = Get-Alias<br>
$WD = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')<br>

run with 1 for default options based on hard coded values, run once a day to set your oof message times automatically

.\aaoof.ps1 1

run with a Date for vacation mode, run once and leave for vacation OOF will turn off when you get back

.\aaoof.ps1 '2023/04/01'

run with nothing for menu

.\aaoof.ps1


================ Email Out of Office Automation ================<br>
Current account is aarosanders@microsoft.com<br>
1: Press '1' Enable Scheduled Auto Reply and Quit<br>
2: Press '2' To set an end date for a extended out of office message<br>
3: Press '3' To set your office hours<br>
4: Press '4' To set your work days<br>
5: Press '5' To set the Auto Reply state to Enable:Disable:Scheduled<br>
6: Press '6' Save Auto Reply Message to Local HTML File<br>
Q: Press 'Q' to quit.<br>
Please make a selection:<br>