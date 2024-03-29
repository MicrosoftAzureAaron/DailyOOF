﻿# First Run
Git Clone this repo to a folder and run the Script
- Example c:\tools\DailyOOF  
  
` mkdir c:\tools\`  
` cd c:\tools\`  
` git clone https://github.com/MicrosoftAzureAaron/DailyOOF`  
` cd DailyOOF`  
` .\aaoof.ps1`  

The first time you run the script it will ask for your input and save the values to your script file

- $global:StartOfShift = $null
- $global:EndOfShift = $null
- $global:WorkDays = $null

After configuring the above during the first run, you can run the script with CLI commands to automate the oof configuration.  

I suggest you run once the script once a day after the start of your shift. Or use option 6 from an admin terminal to create a scheduled task to run the script with option 1, 15 minutes after the start of your shift. You need to have one the script once with no options and set the start time, end time, and days of the week you work, which are saved to the script file in the above global variables<br><br>
` .\aaoof.ps1 6`<br>

# CLI examples

Run with 1 for default options based on stored values, run once a day to set your oof message times automatically

`.\aaoof.ps1 1`

Run with a Date for vacation mode, run once and leave for vacation OOF will turn off when you get back

`.\aaoof.ps1 '4044/04/04'`

Run with nothing for menu

`.\aaoof.ps1`
<br>
# Example Menu
================ Email Out of Office Automation ================  
Current account is  
1: Press '1' Enable Scheduled Auto Reply and Quit  
2: Press '2' To set an end date for a extended out of office message  
<br>
================ Configure the Script Defaults ================  
3: Press '3' To set your office hours and save to script   
4: Press '4' To set your work days and save to script  
<br>
================ Configure the Auto Reply Settings ================  
5: Press '5' To set the Auto Reply state to Enable:Disable:Scheduled  
6: Press '6' To set a Schedule Task to run the 'AAOOF.ps1 1' 15 minutes after the start of your shift daily  
<br>
================ Configure the Auto Reply Message ================  
9: Press '9' Save the current Auto Reply Message to File  
0: Press '0' Load an Auto Reply Message to File  
<br>
Q: Press 'Q' to quit.<br>
<br><br><br>
# TO ADD:

- remove unused functions
- save message to html file for local loading
- load message from file function
- Pre saved messages - load from html file
  - normal oof
  - vacation oof
  - sick oof
  - holiday oof
