<#
.SYNOPSIS
    Daily Out of Office (OOF) Automation Tool with WPF GUI.

.DESCRIPTION
    A PowerShell-based GUI application for managing Exchange Online Out of Office
    auto-reply messages. Supports daily scheduled OOF, vacation/extended OOF,
    custom HTML templates with placeholders, and auto-generated email signatures.
    Can also run headless via CLI parameter for scheduled task automation.

.PARAMETER InputParameter
    Optional CLI parameter for headless operation:
    - '1'          : Run daily scheduled OOF update (for use with Windows Task Scheduler).
    - A date string: Set vacation OOF until the specified return date.

.NOTES
    Requires : ExchangeOnlineManagement module (prompted to install if missing).
    Config   : config/config.json (auto-created, .gitignored).
    Templates: config/*.html (auto-downloaded from GitHub if missing).
    XAML     : config/AAOOF-GUI.xaml (UI layout, auto-downloaded if missing).
#>
param([string]$InputParameter)

# ===================== WPF GUI Setup =====================
# Load .NET assemblies required for WPF windows, controls, and file dialogs.
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Resolve paths for the script directory, config folder, config file, and XAML layout.
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigDir = Join-Path $ScriptDir "config"
$ConfigFile = Join-Path $ConfigDir "config.json"
$XamlFile = Join-Path $ConfigDir "AAOOF-GUI.xaml"

# Ensure config directory exists
if (!(Test-Path $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir | Out-Null }

# ===================== Auto-Download Missing Config Files =====================
# On first run, download XAML layout and HTML templates from the GitHub repository
# so that the tool works out of the box without manual file setup.
$RepoBaseUrl = "https://raw.githubusercontent.com/MicrosoftAzureAaron/DailyOOF/main/config"
$DefaultConfigFiles = @(
    "AAOOF-GUI.xaml",
    "normal_oof.html",
    "vacation_oof.html",
    "sick_oof.html",
    "holiday_oof.html"
)

foreach ($fileName in $DefaultConfigFiles) {
    $localPath = Join-Path $ConfigDir $fileName
    if (!(Test-Path $localPath)) {
        $url = "$RepoBaseUrl/$fileName"
        try {
            Invoke-WebRequest -Uri $url -OutFile $localPath -UseBasicParsing
            Write-Host "Downloaded missing file: $fileName" -ForegroundColor Green
        }
        catch {
            Write-Host "Warning: Could not download $fileName from $url" -ForegroundColor Yellow
        }
    }
}

# ===================== Configuration (loaded from config.json) =====================
# Global variables hold the user's settings. Defaults are set here and then
# overwritten by Import-AppConfiguration if a config.json file exists.
$script:StartOfShift   = $null                       # Shift start time (datetime)
$script:EndOfShift     = $null                       # Shift end time (datetime)
$script:WorkDays       = $null                       # Array of day names, e.g. @('Monday','Tuesday',...)
$script:UserAlias      = ""                           # Email address used as Exchange identity
$script:UserAliasSuffix = ""                           # Domain suffix appended to the Windows username
$script:FullName       = ""                           # Display name for auto-generated signature
$script:Role           = ""                           # Job title inserted into templates via [ROLE]
$script:OverrideAccount = $false                      # True if user manually overrides the account email
$script:SelectedHolidayName = ""                      # Name of the selected holiday for [HOLIDAY NAME] placeholder

# Script-level tracking for EXO sync state
$script:IsConnectedToEXO = $false
$script:EXOMessageSynced = $true

# Import-AppConfiguration: Read config.json and populate global variables.
function Import-AppConfiguration {
    if (Test-Path $ConfigFile) {
        $cfg = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        if ($cfg.StartOfShift)    { $script:StartOfShift = [datetime]$cfg.StartOfShift }
        if ($cfg.EndOfShift)      { $script:EndOfShift = [datetime]$cfg.EndOfShift }
        if ($cfg.WorkDays)        { $script:WorkDays = @($cfg.WorkDays) }
        if ($cfg.UserAlias)       { $script:UserAlias = $cfg.UserAlias }
        if ($cfg.UserAliasSuffix) { $script:UserAliasSuffix = $cfg.UserAliasSuffix }
        if ($cfg.FullName)        { $script:FullName = $cfg.FullName }
        if ($cfg.Role)            { $script:Role = $cfg.Role }
        if ($null -ne $cfg.OverrideAccount) { $script:OverrideAccount = [bool]$cfg.OverrideAccount }
    }
}

# Load config immediately on script start
Import-AppConfiguration

# ===================== Core Functions =====================

# Resolve-UserAlias: Build the user's email alias from the Windows login name + suffix.
function Resolve-UserAlias {
    if ([string]::IsNullOrEmpty($script:UserAliasSuffix)) {
        # Try to derive suffix from the machine's DNS domain
        if ($env:USERDNSDOMAIN) {
            $script:UserAliasSuffix = "@$($env:USERDNSDOMAIN.ToLower())"
        }
    }
    $ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    if ($ComputerSystem.Username) {
        $CurrentUser = $ComputerSystem.Username.Split('\')[-1]
    } else {
        $CurrentUser = $env:USERNAME
    }
    $script:UserAlias = "$CurrentUser$script:UserAliasSuffix"
}

# Get-AutoReplyConfigPath: Return the file path for the local auto-reply config cache.
function Get-AutoReplyConfigPath {
    return Join-Path $ConfigDir "AutoReplyConfig.json"
}

# Save-AutoReplyConfigToFile: Fetch current auto-reply config from Exchange and save to disk.
function Save-AutoReplyConfigToFile {
    $AutoReplyConfigPath = Get-AutoReplyConfigPath
    Get-AutoReplyConfiguration | ConvertTo-Json -Depth 100 | Set-Content $AutoReplyConfigPath
}

# Get-AutoReplyConfiguration: Retrieve the mailbox auto-reply configuration from Exchange Online.
function Get-AutoReplyConfiguration {
    return Get-MailboxAutoReplyConfiguration -Identity $script:UserAlias
}

# Import-AutoReplyConfigFromFile: Load a previously saved auto-reply config from disk.
function Import-AutoReplyConfigFromFile {
    $AutoReplyConfigPath = Get-AutoReplyConfigPath
    return Get-Content $AutoReplyConfigPath -Raw | ConvertFrom-Json
}

# Set-AutoReplyState: Change the auto-reply state (Enabled|Disabled|Scheduled) on Exchange.
function Set-AutoReplyState($State) {
    switch ($State) {
        'Enabled'   { Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState "Enabled" }
        'Disabled'  { Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState "Disabled" }
        'Scheduled' {
            Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState "Scheduled"
            Set-AutoReplyScheduleTimes
        }
    }
    Save-AutoReplyConfigToFile
}

# Set-AutoReplyScheduleTimes: Calculate OOF start/end times based on shift and work days,
# then apply them to the Exchange mailbox configuration.
function Set-AutoReplyScheduleTimes {
    if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift) { return }
    if ($null -eq $script:WorkDays) { return }

    $DaysToAdd = Get-NextWorkDayOffset

    $StartTime = (Get-Date).Date.Add($script:StartOfShift.TimeOfDay).AddDays($DaysToAdd)

    $EndTime = (Get-Date).Date.Add($script:EndOfShift.TimeOfDay)

    if ($DaysToAdd -eq 0) { $EndTime = $EndTime.AddDays(-1) }

    Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -StartTime $EndTime -EndTime $StartTime
    Save-AutoReplyConfigToFile
}

# Set-AutoReplyMessage: Apply an HTML message body as the auto-reply for Internal, External, or Both.
function Set-AutoReplyMessage($Message, $MessageScope) {
    switch ($MessageScope) {
        'Internal' { Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -InternalMessage $Message }
        'External' { Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -ExternalMessage $Message }
        default    { Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -ExternalMessage $Message -InternalMessage $Message }
    }
}

# Get-NextWorkDayOffset: Calculate the number of days until the next working day.
# Returns 0 if today is a work day and before shift start, 1 for tomorrow, or more
# if the next work day is further out (e.g., over a weekend).
function Get-NextWorkDayOffset {
    if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift) { return 1 }
    if (!$script:WorkDays) { return 1 }

    $CurrentTime = [datetime](Get-Date)

    if (!($CurrentTime.DayOfWeek -in $script:WorkDays)) {
        $DaysAhead = 0
        while (!($CurrentTime.DayOfWeek -in $script:WorkDays)) {
            $DaysAhead += 1
            $CurrentTime = $CurrentTime.AddDays(1)
        }
        return $DaysAhead
    }
    else {
        $NextDay = $CurrentTime.AddDays(1)
        $DaysAhead = 1
        while (!($NextDay.DayOfWeek -in $script:WorkDays)) {
            $DaysAhead += 1
            $NextDay = $NextDay.AddDays(1)
        }
        if ($DaysAhead -gt 1) {
            return $DaysAhead
        }

        $CurrentTime = [datetime](Get-Date)
        if ($CurrentTime -lt $script:StartOfShift) {
            return 0
        }
        else {
            return 1
        }
    }
}

# Connect-ExchangeOnlineSession: Ensure a live Exchange Online connection exists.
# Prompts to install the EXO module if missing, reuses existing sessions if found.
function Connect-ExchangeOnlineSession {
    if ([string]::IsNullOrEmpty($script:UserAlias)) { Resolve-UserAlias }

    # Check if ExchangeOnlineManagement module is available
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        $result = [System.Windows.MessageBox]::Show(
            "The ExchangeOnlineManagement module is required but not installed.`n`nWould you like to install it now?",
            "Module Not Found",
            'YesNo',
            'Warning'
        )
        if ($result -eq 'Yes') {
            Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
        } else {
            throw "ExchangeOnlineManagement module is required to connect."
        }
    }

    $session = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -ne $session) {
        $exchangeSession = $session | Where-Object { $_.Name -like "ExchangeOnline_*" }
        if ($null -ne $exchangeSession) {
            return $true
        }
    }
    Connect-ExchangeOnline -UserPrincipalName $script:UserAlias
    return $true
}

# Disconnect-ExchangeOnlineSession: Safely tear down the Exchange Online connection.
function Disconnect-ExchangeOnlineSession {
    try { Disconnect-ExchangeOnline -Confirm:$false } catch { }
}

# Get-USFederalHolidays: Returns US federal holidays with observed dates and return dates.
function Get-USFederalHolidays {
    param([int]$Year = (Get-Date).Year)
    $holidays = @()
    # Fixed-date holidays (adjusted to observed weekday if on weekend)
    $fixedDates = @(
        @{ Name = "New Year's Day"; Month = 1; Day = 1 },
        @{ Name = "Juneteenth"; Month = 6; Day = 19 },
        @{ Name = "Independence Day"; Month = 7; Day = 4 },
        @{ Name = "Veterans Day"; Month = 11; Day = 11 },
        @{ Name = "Christmas Day"; Month = 12; Day = 25 }
    )
    foreach ($fd in $fixedDates) {
        $date = [datetime]::new($Year, $fd.Month, $fd.Day)
        if ($date.DayOfWeek -eq 'Sunday') { $date = $date.AddDays(1) }
        elseif ($date.DayOfWeek -eq 'Saturday') { $date = $date.AddDays(-1) }
        $holidays += [PSCustomObject]@{ Name = $fd.Name; Date = $date }
    }
    # MLK Day: 3rd Monday of January
    $d = [datetime]::new($Year, 1, 1); while ($d.DayOfWeek -ne 'Monday') { $d = $d.AddDays(1) }
    $holidays += [PSCustomObject]@{ Name = "Martin Luther King Jr. Day"; Date = $d.AddDays(14) }
    # Presidents' Day: 3rd Monday of February
    $d = [datetime]::new($Year, 2, 1); while ($d.DayOfWeek -ne 'Monday') { $d = $d.AddDays(1) }
    $holidays += [PSCustomObject]@{ Name = "Presidents' Day"; Date = $d.AddDays(14) }
    # Memorial Day: Last Monday of May
    $d = [datetime]::new($Year, 5, 31); while ($d.DayOfWeek -ne 'Monday') { $d = $d.AddDays(-1) }
    $holidays += [PSCustomObject]@{ Name = "Memorial Day"; Date = $d }
    # Labor Day: 1st Monday of September
    $d = [datetime]::new($Year, 9, 1); while ($d.DayOfWeek -ne 'Monday') { $d = $d.AddDays(1) }
    $holidays += [PSCustomObject]@{ Name = "Labor Day"; Date = $d }
    # Columbus Day: 2nd Monday of October
    $d = [datetime]::new($Year, 10, 1); while ($d.DayOfWeek -ne 'Monday') { $d = $d.AddDays(1) }
    $holidays += [PSCustomObject]@{ Name = "Columbus Day"; Date = $d.AddDays(7) }
    # Thanksgiving: 4th Thursday of November
    $d = [datetime]::new($Year, 11, 1); while ($d.DayOfWeek -ne 'Thursday') { $d = $d.AddDays(1) }
    $holidays += [PSCustomObject]@{ Name = "Thanksgiving"; Date = $d.AddDays(21) }
    # Sort by date and compute return date (next business day after holiday)
    $holidays = $holidays | Sort-Object Date
    foreach ($h in $holidays) {
        $nextDay = $h.Date.AddDays(1)
        while ($nextDay.DayOfWeek -eq 'Saturday' -or $nextDay.DayOfWeek -eq 'Sunday') { $nextDay = $nextDay.AddDays(1) }
        $h | Add-Member -NotePropertyName ReturnDate -NotePropertyValue $nextDay
    }
    return $holidays
}

# Get-TemplateWarnings: Check the current editor message for unresolved placeholders.
function Get-TemplateWarnings {
    $warnings = @()
    $msg = $txtMessage.Text
    if ([string]::IsNullOrWhiteSpace($msg)) { return $warnings }
    if ($msg -match '\[RETURN DATE\]') {
        $warnings += "No return date selected \u2014 '[RETURN DATE]' will appear as literal text in the email."
    }
    if ($msg -match '\[HOLIDAY NAME\]') {
        $warnings += "No holiday selected \u2014 '[HOLIDAY NAME]' will appear as literal text in the email."
    }
    if ($msg -match '\[ROLE\]') {
        $warnings += "Role not configured \u2014 '[ROLE]' will appear as literal text in the email."
    }
    if ($msg -match '\[SIGNATURE\]') {
        $warnings += "Signature was not resolved \u2014 '[SIGNATURE]' will appear as literal text in the email."
    }
    return $warnings
}

# Export-MessageToFile: Write an HTML message body to disk.
function Export-MessageToFile($FilePath, $Content) {
    $Content | Out-File -FilePath $FilePath -Encoding utf8
}

# Set-VacationAutoReply: Configure an extended/vacation OOF that runs from now
# until the given return date at shift-start time.
function Set-VacationAutoReply($ReturnDate) {
    if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift) { return }
    $ParsedDate = [datetime]$ReturnDate
    $EndTime = $ParsedDate + $script:StartOfShift.TimeOfDay
    Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState "Scheduled" -StartTime $script:EndOfShift -EndTime $EndTime
    Save-AutoReplyConfigToFile
}

# Disable-VacationAutoReply: Turn off the vacation/extended OOF by setting auto-reply to Disabled.
function Disable-VacationAutoReply {
    Set-AutoReplyState 'Disabled'
}

# Register-DailyScheduledTask: Create a Windows Scheduled Task named 'AAOOF' that
# runs this script daily with CLI parameter '1'. Requires the script to be running
# as Administrator; shows a friendly error if not elevated.
function Register-DailyScheduledTask {
    # Check for admin privileges before attempting
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        throw "This action requires Administrator privileges.`n`nPlease close the app and re-run the script as Administrator, then try again."
    }

    if (!(Get-ScheduledTask -TaskName "AAOOF" -ErrorAction SilentlyContinue)) {
        $scriptPath = Join-Path $ScriptDir "AAOOF-GUI.ps1"
        $taskname = "AAOOF"
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "`"$scriptPath`" 1"
        $date = Get-Date -Date (Get-Date).Date
        $TriggerTime = $script:StartOfShift.TimeOfDay
        $TriggerTime = $date.AddMinutes(15) + $TriggerTime
        $trigger = New-ScheduledTaskTrigger -Daily -At $TriggerTime

        Register-ScheduledTask -TaskName $taskname -Trigger $trigger -Action $action -RunLevel Highest -ErrorAction Stop
        return $true
    }
    return $false
}

# Export-AppConfiguration: Persist all global settings to config.json.
function Export-AppConfiguration {
    $cfg = @{
        StartOfShift    = if ($null -ne $script:StartOfShift) { $script:StartOfShift.ToString("o") } else { $null }
        EndOfShift      = if ($null -ne $script:EndOfShift) { $script:EndOfShift.ToString("o") } else { $null }
        WorkDays        = $script:WorkDays
        UserAlias       = $script:UserAlias
        UserAliasSuffix = $script:UserAliasSuffix
        FullName        = $script:FullName
        Role            = $script:Role
        OverrideAccount = $script:OverrideAccount
    }
    $cfg | ConvertTo-Json -Depth 5 | Set-Content $ConfigFile -Encoding utf8
}

# ===================== CLI Mode (for scheduled task / automation) =====================
# When invoked with a parameter, skip the GUI and run headless.
#   '1'   — Daily scheduled OOF update. Checks for active vacation before overwriting.
#   <date> — Set vacation/extended OOF until that return date.
if ($InputParameter) {
    if ($InputParameter -eq '1') {
        if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift -or $null -eq $script:WorkDays) {
            Write-Host "Configuration not set. Please run the GUI first to configure." -ForegroundColor Red
            exit
        }
        Connect-ExchangeOnlineSession

        # Skip if a vacation/extended OOF is active (end time is more than 1 day out)
        $arc = Get-AutoReplyConfiguration
        if ($arc.AutoReplyState -eq 'Scheduled') {
            $endTime = [datetime]$arc.EndTime
            if ($endTime -gt (Get-Date).AddDays(1)) {
                Write-Host "Vacation/extended OOF is active until $endTime — skipping daily update." -ForegroundColor Yellow
                Disconnect-ExchangeOnlineSession
                exit
            }
        }

        Set-AutoReplyState 'Scheduled'
        $arc = Get-AutoReplyConfiguration
        Write-Host "Auto Reply: $($arc.AutoReplyState) | Start: $($arc.StartTime) | End: $($arc.EndTime)"
        Disconnect-ExchangeOnlineSession
        exit
    }
    if ($InputParameter -as [datetime]) {
        if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift) {
            Write-Host "Configuration not set. Please run the GUI first to configure." -ForegroundColor Red
            exit
        }
        Connect-ExchangeOnlineSession
        Set-VacationAutoReply $InputParameter
        $arc = Get-AutoReplyConfiguration
        Write-Host "Auto Reply: $($arc.AutoReplyState) | Start: $($arc.StartTime) | End: $($arc.EndTime)"
        Disconnect-ExchangeOnlineSession
        exit
    }
}

# ===================== Load XAML GUI from File =====================
# Parse the external XAML layout file and build the WPF window.
if (!(Test-Path $XamlFile)) {
    Write-Host "FATAL: XAML file not found at $XamlFile" -ForegroundColor Red
    exit 1
}
[xml]$XAML = Get-Content $XamlFile -Raw

# ===================== Build the Window =====================
# Instantiate the WPF window from XAML and bind all named controls to variables.
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# --- Quick Actions tab controls ---
$txtAccount = $Window.FindName("txtAccount")
$chkOverrideAccount = $Window.FindName("chkOverrideAccount")
$txtConnectionStatus = $Window.FindName("txtConnectionStatus")
$btnConnect = $Window.FindName("btnConnect")
$btnDisconnect = $Window.FindName("btnDisconnect")
$btnEnableScheduled = $Window.FindName("btnEnableScheduled")
$dpReturnDate = $Window.FindName("dpReturnDate")
$cmbHoliday = $Window.FindName("cmbHoliday")
$btnSetVacation = $Window.FindName("btnSetVacation")
$btnCancelVacation = $Window.FindName("btnCancelVacation")
$txtARCState = $Window.FindName("txtARCState")
$txtARCStart = $Window.FindName("txtARCStart")
$txtARCEnd = $Window.FindName("txtARCEnd")
$btnRefreshStatus = $Window.FindName("btnRefreshStatus")
$btnViewCurrentMsg = $Window.FindName("btnViewCurrentMsg")
$tcMain = $Window.FindName("tcMain")
$wbCurrentOOF = $Window.FindName("wbCurrentOOF")
$btnRefreshCurrentOOF = $Window.FindName("btnRefreshCurrentOOF")
$txtCurrentOOFStatus = $Window.FindName("txtCurrentOOFStatus")

# --- Configuration tab controls ---
$txtFullName = $Window.FindName("txtFullName")
$txtRole = $Window.FindName("txtRole")
$txtSuffix = $Window.FindName("txtSuffix")
$cmbStartHour = $Window.FindName("cmbStartHour")
$cmbStartMin = $Window.FindName("cmbStartMin")
$cmbStartAmPm = $Window.FindName("cmbStartAmPm")
$cmbEndHour = $Window.FindName("cmbEndHour")
$cmbEndMin = $Window.FindName("cmbEndMin")
$cmbEndAmPm = $Window.FindName("cmbEndAmPm")
$chkMon = $Window.FindName("chkMon")
$chkTue = $Window.FindName("chkTue")
$chkWed = $Window.FindName("chkWed")
$chkThu = $Window.FindName("chkThu")
$chkFri = $Window.FindName("chkFri")
$chkSat = $Window.FindName("chkSat")
$chkSun = $Window.FindName("chkSun")
$btnPresetMF = $Window.FindName("btnPresetMF")
$btnPresetSunWed = $Window.FindName("btnPresetSunWed")
$btnPresetWedSat = $Window.FindName("btnPresetWedSat")
$btnStateEnabled = $Window.FindName("btnStateEnabled")
$btnStateDisabled = $Window.FindName("btnStateDisabled")
$btnStateScheduled = $Window.FindName("btnStateScheduled")
$btnCreateTask = $Window.FindName("btnCreateTask")

# --- Message Templates tab controls ---
$cmbTemplate = $Window.FindName("cmbTemplate")
$btnLoadTemplate = $Window.FindName("btnLoadTemplate")
$btnBrowseFile = $Window.FindName("btnBrowseFile")
$txtMessage = $Window.FindName("txtMessage")
$btnApplyInternal = $Window.FindName("btnApplyInternal")
$btnApplyExternal = $Window.FindName("btnApplyExternal")
$btnApplyBoth = $Window.FindName("btnApplyBoth")
$btnSaveTemplate = $Window.FindName("btnSaveTemplate")
$btnSaveOnlineMsg = $Window.FindName("btnSaveOnlineMsg")

# --- Template Options panel controls ---
$chkIncludeSignature = $Window.FindName("chkIncludeSignature")
$chkIncludeOfficeHours = $Window.FindName("chkIncludeOfficeHours")
$chkIncludeWorkDays = $Window.FindName("chkIncludeWorkDays")
$chkIncludeTimezone = $Window.FindName("chkIncludeTimezone")
$tcMessageView = $Window.FindName("tcMessageView")
$wbPreview = $Window.FindName("wbPreview")

# --- HTML Formatting toolbar controls ---
$btnFmtBold = $Window.FindName("btnFmtBold")
$btnFmtItalic = $Window.FindName("btnFmtItalic")
$btnFmtUnderline = $Window.FindName("btnFmtUnderline")
$btnFmtH3 = $Window.FindName("btnFmtH3")
$btnFmtP = $Window.FindName("btnFmtP")
$btnFmtBr = $Window.FindName("btnFmtBr")
$btnFmtLink = $Window.FindName("btnFmtLink")
$btnFmtColor = $Window.FindName("btnFmtColor")
$btnFmtList = $Window.FindName("btnFmtList")
$btnFmtRef = $Window.FindName("btnFmtRef")

# --- Status bar ---
$txtStatusBar = $Window.FindName("txtStatusBar")
$borderStatusBar = $Window.FindName("borderStatusBar")

# ===================== Helper: UI Dialog Functions =====================
# Update-StatusBar: Set the bottom status bar text and force a UI render.
function Update-StatusBar($Message) {
    $txtStatusBar.Text = $Message
    # Update status bar color based on EXO sync state
    if ($script:IsConnectedToEXO -and -not $script:EXOMessageSynced) {
        $borderStatusBar.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xD8, 0x3B, 0x01))
        $txtStatusBar.Text = [char]0x26A0 + " Message not yet applied to Exchange | $Message"
    } else {
        $borderStatusBar.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
    }
    $Window.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
}

# Show-InfoDialog: Display an informational popup.
function Show-InfoDialog($Title, $Message) {
    [System.Windows.MessageBox]::Show($Message, $Title, 'OK', 'Information')
}

# Show-ErrorDialog: Display an error popup.
function Show-ErrorDialog($Title, $Message) {
    [System.Windows.MessageBox]::Show($Message, $Title, 'OK', 'Error')
}

# Ensure-ExchangeConnection: Check for an active Exchange Online session and auto-connect
# if none exists. Updates the connection status UI. Returns $true if connected, $false on failure.
function Ensure-ExchangeConnection {
    $session = Get-ConnectionInformation -ErrorAction SilentlyContinue
    $connected = $null -ne ($session | Where-Object { $_.Name -like "ExchangeOnline_*" })
    if ($connected) { return $true }

    Update-StatusBar "Not connected — attempting to connect..."
    if (-not $chkOverrideAccount.IsChecked) {
        $script:UserAliasSuffix = $txtSuffix.Text
        Resolve-UserAlias
        $txtAccount.Text = $script:UserAlias
    } else {
        $script:UserAlias = $txtAccount.Text
    }
    Connect-ExchangeOnlineSession
    $txtConnectionStatus.Text = "Connected"
    $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Green
    $script:IsConnectedToEXO = $true
    return $true
}

# ===================== Populate Combos =====================
# Fill the hour, minute, and AM/PM dropdown lists for shift time selection.
1..12 | ForEach-Object { $cmbStartHour.Items.Add($_.ToString()) | Out-Null; $cmbEndHour.Items.Add($_.ToString()) | Out-Null }
@("00","15","30","45") | ForEach-Object { $cmbStartMin.Items.Add($_) | Out-Null; $cmbEndMin.Items.Add($_) | Out-Null }
@("AM","PM") | ForEach-Object { $cmbStartAmPm.Items.Add($_) | Out-Null; $cmbEndAmPm.Items.Add($_) | Out-Null }

# ===================== Load Saved Config into UI =====================
# Initialize-UIFromConfig: Populate all UI controls from global config values,
# setting sensible defaults where config values are missing.
function Initialize-UIFromConfig {
    # Full Name
    if (![string]::IsNullOrEmpty($script:FullName)) {
        $txtFullName.Text = $script:FullName
    }
    # Role
    if (![string]::IsNullOrEmpty($script:Role)) {
        $txtRole.Text = $script:Role
    }
    # Suffix
    if (![string]::IsNullOrEmpty($script:UserAliasSuffix)) {
        $txtSuffix.Text = $script:UserAliasSuffix
    }
    # Account override checkbox
    $chkOverrideAccount.IsChecked = $script:OverrideAccount
    $txtAccount.IsEnabled = $script:OverrideAccount
    # Account
    if (![string]::IsNullOrEmpty($script:UserAlias)) {
        $txtAccount.Text = $script:UserAlias
    } else {
        Resolve-UserAlias
        $txtAccount.Text = $script:UserAlias
    }

    # Shift times
    if ($null -ne $script:StartOfShift) {
        $h = (Get-Date $script:StartOfShift).Hour
        $m = (Get-Date $script:StartOfShift).Minute
        $ampm = if ($h -ge 12) { "PM" } else { "AM" }
        $displayH = if ($h -gt 12) { $h - 12 } elseif ($h -eq 0) { 12 } else { $h }
        $cmbStartHour.SelectedItem = $displayH.ToString()
        $nearestMin = @("00","15","30","45") | Sort-Object { [Math]::Abs([int]$_ - $m) } | Select-Object -First 1
        $cmbStartMin.SelectedItem = $nearestMin
        $cmbStartAmPm.SelectedItem = $ampm
    } else {
        $cmbStartHour.SelectedIndex = 8; $cmbStartMin.SelectedIndex = 0; $cmbStartAmPm.SelectedIndex = 0 # 9 AM
    }

    if ($null -ne $script:EndOfShift) {
        $h = (Get-Date $script:EndOfShift).Hour
        $m = (Get-Date $script:EndOfShift).Minute
        $ampm = if ($h -ge 12) { "PM" } else { "AM" }
        $displayH = if ($h -gt 12) { $h - 12 } elseif ($h -eq 0) { 12 } else { $h }
        $cmbEndHour.SelectedItem = $displayH.ToString()
        $nearestMin = @("00","15","30","45") | Sort-Object { [Math]::Abs([int]$_ - $m) } | Select-Object -First 1
        $cmbEndMin.SelectedItem = $nearestMin
        $cmbEndAmPm.SelectedItem = $ampm
    } else {
        $cmbEndHour.SelectedIndex = 5; $cmbEndMin.SelectedIndex = 0; $cmbEndAmPm.SelectedIndex = 1 # 6 PM
    }

    # Work days
    if ($script:WorkDays) {
        $chkMon.IsChecked = ('Monday' -in $script:WorkDays)
        $chkTue.IsChecked = ('Tuesday' -in $script:WorkDays)
        $chkWed.IsChecked = ('Wednesday' -in $script:WorkDays)
        $chkThu.IsChecked = ('Thursday' -in $script:WorkDays)
        $chkFri.IsChecked = ('Friday' -in $script:WorkDays)
        $chkSat.IsChecked = ('Saturday' -in $script:WorkDays)
        $chkSun.IsChecked = ('Sunday' -in $script:WorkDays)
    } else {
        # Default Mon-Fri
        $chkMon.IsChecked = $true; $chkTue.IsChecked = $true; $chkWed.IsChecked = $true
        $chkThu.IsChecked = $true; $chkFri.IsChecked = $true
    }
}

# Resolve-TemplateFilePath: Map a template display name to its file path in the config directory.
function Resolve-TemplateFilePath($TemplateName) {
    switch ($TemplateName) {
        "Normal OOF"   { return Join-Path $ConfigDir "normal_oof.html" }
        "Vacation OOF" { return Join-Path $ConfigDir "vacation_oof.html" }
        "Sick OOF"     { return Join-Path $ConfigDir "sick_oof.html" }
        "Holiday OOF"  { return Join-Path $ConfigDir "holiday_oof.html" }
        default        { return $null }
    }
}

# Resolve-TemplatePlaceholders: Process an HTML template string, replacing:
#   [RETURN DATE] — with the selected return date from the date picker
#   [ROLE]        — with the user's configured role (or default)
#   [SIGNATURE]   — with an auto-generated signature block (or removed if unchecked)
# The signature block includes: greeting, display name, office details, and email link.
function Resolve-TemplatePlaceholders($text) {
    # Replace [RETURN DATE] if present
    if ($null -ne $dpReturnDate -and $null -ne $dpReturnDate.SelectedDate) {
        $text = $text -replace '\[RETURN DATE\]', $dpReturnDate.SelectedDate.ToString('MMMM d, yyyy')
    }

    # Replace [HOLIDAY NAME] with selected holiday or generic fallback
    $holidayName = if (![string]::IsNullOrWhiteSpace($script:SelectedHolidayName)) { $script:SelectedHolidayName } else { 'a company holiday' }
    $text = $text -replace '\[HOLIDAY NAME\]', $holidayName

    # Replace [ROLE] with the role from the text box, or generic fallback
    $role = if (![string]::IsNullOrWhiteSpace($txtRole.Text)) { $txtRole.Text } else { 'member of my team' }
    $text = $text -replace '\[ROLE\]', $role

    # Auto-generate signature block: the greeting/name is conditional on the
    # "Include Signature" checkbox, but office details are always included per their own toggles.
    $sigLines = @()

    if ($chkIncludeSignature.IsChecked) {
        # Use the Full Name text box, fall back to alias-derived name
        if (![string]::IsNullOrWhiteSpace($txtFullName.Text)) {
            $displayName = $txtFullName.Text
        } else {
            $aliasLocal = ($script:UserAlias -split '@')[0]
            if ($aliasLocal) {
                if ($aliasLocal -match '\.' ) {
                    $nameParts = $aliasLocal -split '\.'
                } else {
                    $nameParts = [regex]::Split($aliasLocal, '(?<=[a-z])(?=[A-Z])')
                }
                $displayName = ($nameParts | ForEach-Object { (Get-Culture).TextInfo.ToTitleCase($_.ToLower()) }) -join ' '
            } else {
                $displayName = $env:USERNAME
            }
        }
        $sigLines += "<p><b>Best Regards,</b><br/>"
        $sigLines += "$displayName</p>"
    }

    # Office details line (independent of signature toggle)
    $detailParts = @()
    if ($chkIncludeOfficeHours.IsChecked -and $null -ne $script:StartOfShift -and $null -ne $script:EndOfShift) {
        $detailParts += "$($script:StartOfShift.ToString('h:mm tt')) - $($script:EndOfShift.ToString('h:mm tt'))"
    }
    if ($chkIncludeTimezone.IsChecked) {
        $detailParts += [System.TimeZoneInfo]::Local.DisplayName
    }
    if ($chkIncludeWorkDays.IsChecked -and $script:WorkDays) {
        $weekOrder = @('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday')
        $sorted = $script:WorkDays | Sort-Object { $weekOrder.IndexOf($_) }
        $detailParts += ($sorted -join ', ')
    }
    if ($detailParts.Count -gt 0) {
        $sigLines += "<p style='color: #555; font-size: 10pt;'>$($detailParts -join ' | ')</p>"
    }

    # Email link (independent of signature toggle)
    if (![string]::IsNullOrWhiteSpace($script:UserAlias)) {
        $sigLines += "<p><a href='mailto:$($script:UserAlias)'>$($script:UserAlias)</a></p>"
    }

    if ($sigLines.Count -gt 0) {
        $signatureHtml = $sigLines -join "`n"
        $text = $text -replace '\[SIGNATURE\]', $signatureHtml
    } else {
        $text = $text -replace '(?m)^\s*\[SIGNATURE\]\s*\r?\n?', ''
    }
    return $text
}

# Read-WorkDaysFromUI: Collect the checked work-day checkboxes into an array of day names.
function Read-WorkDaysFromUI {
    $days = @()
    if ($chkSun.IsChecked) { $days += 'Sunday' }
    if ($chkMon.IsChecked) { $days += 'Monday' }
    if ($chkTue.IsChecked) { $days += 'Tuesday' }
    if ($chkWed.IsChecked) { $days += 'Wednesday' }
    if ($chkThu.IsChecked) { $days += 'Thursday' }
    if ($chkFri.IsChecked) { $days += 'Friday' }
    if ($chkSat.IsChecked) { $days += 'Saturday' }
    return $days
}

# Read-ShiftTimesFromUI: Parse the start/end time combo boxes into datetime globals,
# converting from 12-hour format (with AM/PM) to 24-hour.
function Read-ShiftTimesFromUI {
    $StartHour = [int]$cmbStartHour.SelectedItem
    $StartMinute = [int]$cmbStartMin.SelectedItem
    $StartAmPm = $cmbStartAmPm.SelectedItem
    if ($StartAmPm -eq "PM" -and $StartHour -ne 12) { $StartHour += 12 }
    if ($StartAmPm -eq "AM" -and $StartHour -eq 12) { $StartHour = 0 }
    $script:StartOfShift = [datetime](Get-Date).Date.AddHours($StartHour).AddMinutes($StartMinute)

    $EndHour = [int]$cmbEndHour.SelectedItem
    $EndMinute = [int]$cmbEndMin.SelectedItem
    $EndAmPm = $cmbEndAmPm.SelectedItem
    if ($EndAmPm -eq "PM" -and $EndHour -ne 12) { $EndHour += 12 }
    if ($EndAmPm -eq "AM" -and $EndHour -eq 12) { $EndHour = 0 }
    $script:EndOfShift = [datetime](Get-Date).Date.AddHours($EndHour).AddMinutes($EndMinute)
}

# ===================== Event Handlers =====================
# Wire up button clicks, checkbox changes, and other UI events to their logic.

# Connect: Resolve the user alias and establish an Exchange Online session.
$btnConnect.Add_Click({
    try {
        Update-StatusBar "Connecting to Exchange Online..."
        if (-not $chkOverrideAccount.IsChecked) {
            $script:UserAliasSuffix = $txtSuffix.Text
            Resolve-UserAlias
            $txtAccount.Text = $script:UserAlias
        } else {
            $script:UserAlias = $txtAccount.Text
        }
        Connect-ExchangeOnlineSession
        $txtConnectionStatus.Text = "Connected"
        $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Green
        $script:IsConnectedToEXO = $true

        # On first connect, pull current OOF config and message and save locally
        try {
            $arc = Get-AutoReplyConfiguration
            $txtARCState.Text = $arc.AutoReplyState
            $txtARCStart.Text = $arc.StartTime.ToString()
            $txtARCEnd.Text = $arc.EndTime.ToString()

            # Compare EXO message with editor to detect mismatch
            $currentMsg = ($txtMessage.Text -replace '\s+', ' ').Trim()
            $exoMsg = if ($arc.ExternalMessage) { ($arc.ExternalMessage -replace '\s+', ' ').Trim() } else { '' }
            $script:EXOMessageSynced = ($currentMsg -eq $exoMsg) -or [string]::IsNullOrWhiteSpace($txtMessage.Text)

            # Save the current online messages to template files if we don't already have a saved message
            $savedMsgFile = Join-Path $ConfigDir "message.html"
            if (!(Test-Path $savedMsgFile) -and ![string]::IsNullOrWhiteSpace($arc.ExternalMessage)) {
                Export-MessageToFile $savedMsgFile $arc.ExternalMessage
            }
        } catch { }

        Export-AppConfiguration
        Update-StatusBar "Connected as $($script:UserAlias)"
    }
    catch {
        $txtConnectionStatus.Text = "Connection Failed"
        $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Red
        Show-ErrorDialog "Connection Error" $_.Exception.Message
        Update-StatusBar "Connection failed"
    }
})

# Disconnect: Tear down the Exchange Online session and update the UI.
$btnDisconnect.Add_Click({
    try {
        Disconnect-ExchangeOnlineSession
        $txtConnectionStatus.Text = "Disconnected"
        $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Red
        $script:IsConnectedToEXO = $false
        $script:EXOMessageSynced = $true
        Update-StatusBar "Disconnected from Exchange Online"
    }
    catch {
        Show-ErrorDialog "Disconnect Error" $_.Exception.Message
    }
})

# Enable Scheduled Auto Reply: Read shift/work-day settings and apply Scheduled mode.
$btnEnableScheduled.Add_Click({
    try {
        Ensure-ExchangeConnection
        Update-StatusBar "Setting scheduled auto reply..."
        Read-ShiftTimesFromUI
        $script:WorkDays = Read-WorkDaysFromUI
        Set-AutoReplyState 'Scheduled'
        $arc = Get-AutoReplyConfiguration
        $txtARCState.Text = $arc.AutoReplyState
        $txtARCStart.Text = $arc.StartTime.ToString()
        $txtARCEnd.Text = $arc.EndTime.ToString()
        Update-StatusBar "Scheduled auto reply enabled"
        Show-InfoDialog "Success" "Scheduled Auto Reply enabled.`nStart: $($arc.StartTime)`nEnd: $($arc.EndTime)"
    }
    catch {
        Show-ErrorDialog "Error" $_.Exception.Message
        Update-StatusBar "Failed to set scheduled auto reply"
    }
})

# Set Vacation OOF: Configure an extended OOF until the selected return date.
$btnSetVacation.Add_Click({
    try {
        if ($null -eq $dpReturnDate.SelectedDate) {
            Show-ErrorDialog "Missing Date" "Please select a return date."
            return
        }
        Ensure-ExchangeConnection
        Update-StatusBar "Setting vacation OOF..."
        Read-ShiftTimesFromUI
        $returnDate = $dpReturnDate.SelectedDate.ToString("yyyy/MM/dd")
        Set-VacationAutoReply $returnDate

        # If there's a vacation template loaded, apply it
        $vacPath = Resolve-TemplateFilePath "Vacation OOF"
        if ((Test-Path $vacPath) -and [string]::IsNullOrWhiteSpace($txtMessage.Text) -eq $false) {
            # User may want to apply the loaded message - handled separately via Apply buttons
        }

        $arc = Get-AutoReplyConfiguration
        $txtARCState.Text = $arc.AutoReplyState
        $txtARCStart.Text = $arc.StartTime.ToString()
        $txtARCEnd.Text = $arc.EndTime.ToString()
        Update-StatusBar "Vacation OOF set until $returnDate"
        Show-InfoDialog "Success" "Vacation OOF enabled until $returnDate`nStart: $($arc.StartTime)`nEnd: $($arc.EndTime)"
    }
    catch {
        Show-ErrorDialog "Error" $_.Exception.Message
        Update-StatusBar "Failed to set vacation OOF"
    }
})

# Cancel Vacation OOF: Disable the vacation/extended OOF.
$btnCancelVacation.Add_Click({
    try {
        Ensure-ExchangeConnection
        Update-StatusBar "Cancelling vacation OOF..."
        Disable-VacationAutoReply
        $arc = Get-AutoReplyConfiguration
        $txtARCState.Text = $arc.AutoReplyState
        $txtARCStart.Text = $arc.StartTime.ToString()
        $txtARCEnd.Text = $arc.EndTime.ToString()
        Update-StatusBar "Vacation OOF cancelled"
        Show-InfoDialog "Success" "Vacation OOF has been disabled."
    }
    catch {
        Show-ErrorDialog "Error" $_.Exception.Message
        Update-StatusBar "Failed to cancel vacation OOF"
    }
})

# Refresh Status: Pull the current auto-reply state and schedule from Exchange.
$btnRefreshStatus.Add_Click({
    try {
        Ensure-ExchangeConnection
        Update-StatusBar "Refreshing status..."
        $arc = Get-AutoReplyConfiguration
        $txtARCState.Text = $arc.AutoReplyState
        $txtARCStart.Text = $arc.StartTime.ToString()
        $txtARCEnd.Text = $arc.EndTime.ToString()
        Update-StatusBar "Status refreshed"
    }
    catch {
        Show-ErrorDialog "Error" "Could not refresh. Are you connected?`n$($_.Exception.Message)"
        Update-StatusBar "Refresh failed"
    }
})

# View Current OOF Message: Auto-connect if needed, fetch the live auto-reply,
# render it on the Current OOF tab, and switch to that tab.
$btnViewCurrentMsg.Add_Click({
    try {
        Update-StatusBar "Fetching current OOF message..."
        $txtCurrentOOFStatus.Text = "Loading..."

        # Check connection — attempt to connect if disconnected
        $session = Get-ConnectionInformation -ErrorAction SilentlyContinue
        $connected = $null -ne ($session | Where-Object { $_.Name -like "ExchangeOnline_*" })
        if (-not $connected) {
            $txtCurrentOOFStatus.Text = "Disconnected — reconnecting..."
            Update-StatusBar "Not connected — attempting to connect..."
            if (-not $chkOverrideAccount.IsChecked) {
                $script:UserAliasSuffix = $txtSuffix.Text
                Resolve-UserAlias
                $txtAccount.Text = $script:UserAlias
            } else {
                $script:UserAlias = $txtAccount.Text
            }
            Connect-ExchangeOnlineSession
            $txtConnectionStatus.Text = "Connected"
            $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Green
            $script:IsConnectedToEXO = $true
        }

        $arc = Get-AutoReplyConfiguration
        $msg = if (![string]::IsNullOrWhiteSpace($arc.ExternalMessage)) { $arc.ExternalMessage }
               elseif (![string]::IsNullOrWhiteSpace($arc.InternalMessage)) { $arc.InternalMessage }
               else { $null }
        if ($null -eq $msg) {
            $wbCurrentOOF.NavigateToString("<html><body style='font-family:Segoe UI;padding:20px;color:#888;'><h3>No OOF message is currently set.</h3></body></html>")
            $txtCurrentOOFStatus.Text = "No message set"
            Update-StatusBar "No current OOF message"
        } else {
            $wbCurrentOOF.NavigateToString($msg)
            $txtCurrentOOFStatus.Text = "State: $($arc.AutoReplyState) | Loaded $(Get-Date -Format 'h:mm tt')"
            Update-StatusBar "Current OOF message loaded"
        }

        $txtARCState.Text = $arc.AutoReplyState
        $txtARCStart.Text = $arc.StartTime.ToString()
        $txtARCEnd.Text = $arc.EndTime.ToString()

        # Switch to the Current OOF tab
        $tcMain.SelectedIndex = 3
    }
    catch {
        $wbCurrentOOF.NavigateToString("<html><body style='font-family:Segoe UI;padding:20px;color:red;'><h3>Error</h3><p>Could not fetch message.</p><p style='color:#888;font-size:10pt;'>$([System.Web.HttpUtility]::HtmlEncode($_.Exception.Message))</p></body></html>")
        $txtCurrentOOFStatus.Text = "Error loading"
        Update-StatusBar "Failed to fetch current OOF message"
    }
})

# Refresh Current OOF tab: Check connection (auto-reconnect if needed),
# then fetch and render the live auto-reply message in the WebBrowser.
$btnRefreshCurrentOOF.Add_Click({
    try {
        Update-StatusBar "Fetching current OOF message..."
        $txtCurrentOOFStatus.Text = "Loading..."

        # Check connection — attempt to connect if disconnected
        $session = Get-ConnectionInformation -ErrorAction SilentlyContinue
        $connected = $null -ne ($session | Where-Object { $_.Name -like "ExchangeOnline_*" })
        if (-not $connected) {
            $txtCurrentOOFStatus.Text = "Disconnected — reconnecting..."
            Update-StatusBar "Not connected — attempting to connect..."
            try {
                if (-not $chkOverrideAccount.IsChecked) {
                    $script:UserAliasSuffix = $txtSuffix.Text
                    Resolve-UserAlias
                    $txtAccount.Text = $script:UserAlias
                } else {
                    $script:UserAlias = $txtAccount.Text
                }
                Connect-ExchangeOnlineSession
                $txtConnectionStatus.Text = "Connected"
                $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Green
                $script:IsConnectedToEXO = $true
            }
            catch {
                $txtConnectionStatus.Text = "Connection Failed"
                $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Red
                $wbCurrentOOF.NavigateToString("<html><body style='font-family:Segoe UI;padding:20px;color:red;'><h3>Connection Failed</h3><p>Could not connect to Exchange Online. Please check your account settings and try again.</p><p style='color:#888;font-size:10pt;'>$([System.Web.HttpUtility]::HtmlEncode($_.Exception.Message))</p></body></html>")
                $txtCurrentOOFStatus.Text = "Connection failed"
                Update-StatusBar "Connection failed"
                return
            }
        }

        $arc = Get-AutoReplyConfiguration
        $msg = if (![string]::IsNullOrWhiteSpace($arc.ExternalMessage)) { $arc.ExternalMessage }
               elseif (![string]::IsNullOrWhiteSpace($arc.InternalMessage)) { $arc.InternalMessage }
               else { $null }
        if ($null -eq $msg) {
            $wbCurrentOOF.NavigateToString("<html><body style='font-family:Segoe UI;padding:20px;color:#888;'><h3>No OOF message is currently set.</h3></body></html>")
            $txtCurrentOOFStatus.Text = "No message set"
            Update-StatusBar "No current OOF message"
        } else {
            $wbCurrentOOF.NavigateToString($msg)
            $txtCurrentOOFStatus.Text = "State: $($arc.AutoReplyState) | Loaded $(Get-Date -Format 'h:mm tt')"
            Update-StatusBar "Current OOF message loaded"
        }
    }
    catch {
        $wbCurrentOOF.NavigateToString("<html><body style='font-family:Segoe UI;padding:20px;color:red;'><h3>Error</h3><p>Could not fetch message.</p><p style='color:#888;font-size:10pt;'>$([System.Web.HttpUtility]::HtmlEncode($_.Exception.Message))</p></body></html>")
        $txtCurrentOOFStatus.Text = "Error loading"
        Update-StatusBar "Failed to fetch current OOF message"
    }
})

# Auto-load Current OOF tab on first visit: connect if needed and fetch the message.
# When switching to the Message Templates tab, sync config globals from UI
# so templates reflect the latest unsaved changes (work days, shift times, etc.).
$script:CurrentOOFLoaded = $false
$tcMain.Add_SelectionChanged({
    if ($tcMain.SelectedIndex -eq 2) {
        # Sync config globals from UI before template rendering
        Read-ShiftTimesFromUI
        $script:WorkDays = Read-WorkDaysFromUI
        $script:FullName = $txtFullName.Text
        $script:Role = $txtRole.Text
        $script:UserAliasSuffix = $txtSuffix.Text
        & $optionReloadHandler
        return
    }
    if ($tcMain.SelectedIndex -ne 3) { return }
    if ($script:CurrentOOFLoaded) { return }
    $script:CurrentOOFLoaded = $true
    try {
        Update-StatusBar "Loading current OOF message..."
        $txtCurrentOOFStatus.Text = "Loading..."

        $session = Get-ConnectionInformation -ErrorAction SilentlyContinue
        $connected = $null -ne ($session | Where-Object { $_.Name -like "ExchangeOnline_*" })
        if (-not $connected) {
            $txtCurrentOOFStatus.Text = "Disconnected — connecting..."
            Update-StatusBar "Not connected — attempting to connect..."
            if (-not $chkOverrideAccount.IsChecked) {
                $script:UserAliasSuffix = $txtSuffix.Text
                Resolve-UserAlias
                $txtAccount.Text = $script:UserAlias
            } else {
                $script:UserAlias = $txtAccount.Text
            }
            Connect-ExchangeOnlineSession
            $txtConnectionStatus.Text = "Connected"
            $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Green
            $script:IsConnectedToEXO = $true
        }

        $arc = Get-AutoReplyConfiguration
        $msg = if (![string]::IsNullOrWhiteSpace($arc.ExternalMessage)) { $arc.ExternalMessage }
               elseif (![string]::IsNullOrWhiteSpace($arc.InternalMessage)) { $arc.InternalMessage }
               else { $null }
        if ($null -eq $msg) {
            $wbCurrentOOF.NavigateToString("<html><body style='font-family:Segoe UI;padding:20px;color:#888;'><h3>No OOF message is currently set.</h3></body></html>")
            $txtCurrentOOFStatus.Text = "No message set"
            Update-StatusBar "No current OOF message"
        } else {
            $wbCurrentOOF.NavigateToString($msg)
            $txtCurrentOOFStatus.Text = "State: $($arc.AutoReplyState) | Loaded $(Get-Date -Format 'h:mm tt')"
            Update-StatusBar "Current OOF message loaded"
        }

        $txtARCState.Text = $arc.AutoReplyState
        $txtARCStart.Text = $arc.StartTime.ToString()
        $txtARCEnd.Text = $arc.EndTime.ToString()
    }
    catch {
        $script:CurrentOOFLoaded = $false
        $wbCurrentOOF.NavigateToString("<html><body style='font-family:Segoe UI;padding:20px;color:red;'><h3>Error</h3><p>Could not load message.</p><p style='color:#888;font-size:10pt;'>$([System.Web.HttpUtility]::HtmlEncode($_.Exception.Message))</p></body></html>")
        $txtCurrentOOFStatus.Text = "Error loading"
        Update-StatusBar "Failed to load current OOF message"
    }
})

# Save All Settings: Read every config field from the UI and persist to config.json.
$btnSaveAllConfig = $Window.FindName("btnSaveAllConfig")
$btnSaveAllConfig.Add_Click({
    # Profile
    $script:FullName = $txtFullName.Text
    $script:Role = $txtRole.Text
    # Suffix & account
    $script:UserAliasSuffix = $txtSuffix.Text
    if (-not $chkOverrideAccount.IsChecked) {
        Resolve-UserAlias
        $txtAccount.Text = $script:UserAlias
    } else {
        $script:UserAlias = $txtAccount.Text
    }
    # Shift times
    Read-ShiftTimesFromUI
    # Work days
    $script:WorkDays = Read-WorkDaysFromUI
    Export-AppConfiguration
    Update-StatusBar "All settings saved"
    Show-InfoDialog "Saved" "All configuration settings have been saved."
})

# Work day presets: Quick-fill checkbox groups for common schedules.
# Preset Mon-Fri (standard 5x8)
$btnPresetMF.Add_Click({
    $chkMon.IsChecked = $true; $chkTue.IsChecked = $true; $chkWed.IsChecked = $true
    $chkThu.IsChecked = $true; $chkFri.IsChecked = $true
    $chkSat.IsChecked = $false; $chkSun.IsChecked = $false
})

# Preset Sun-Wed (4x10 schedule)
$btnPresetSunWed.Add_Click({
    $chkSun.IsChecked = $true; $chkMon.IsChecked = $true; $chkTue.IsChecked = $true
    $chkWed.IsChecked = $true; $chkThu.IsChecked = $false
    $chkFri.IsChecked = $false; $chkSat.IsChecked = $false
})

# Preset Wed-Sat (4x10 schedule)
$btnPresetWedSat.Add_Click({
    $chkWed.IsChecked = $true; $chkThu.IsChecked = $true
    $chkFri.IsChecked = $true; $chkSat.IsChecked = $true
    $chkSun.IsChecked = $false; $chkMon.IsChecked = $false; $chkTue.IsChecked = $false
})

# Auto Reply State buttons: Directly set the auto-reply mode on Exchange.
$btnStateEnabled.Add_Click({
    try {
        Ensure-ExchangeConnection
        Update-StatusBar "Setting auto reply to Enabled..."
        Set-AutoReplyState 'Enabled'
        Update-StatusBar "Auto reply set to Enabled"
        Show-InfoDialog "Done" "Auto Reply State set to Enabled"
    }
    catch { Show-ErrorDialog "Error" $_.Exception.Message }
})

$btnStateDisabled.Add_Click({
    try {
        Ensure-ExchangeConnection
        Update-StatusBar "Setting auto reply to Disabled..."
        Set-AutoReplyState 'Disabled'
        Update-StatusBar "Auto reply set to Disabled"
        Show-InfoDialog "Done" "Auto Reply State set to Disabled"
    }
    catch { Show-ErrorDialog "Error" $_.Exception.Message }
})

$btnStateScheduled.Add_Click({
    try {
        Ensure-ExchangeConnection
        Update-StatusBar "Setting auto reply to Scheduled..."
        Read-ShiftTimesFromUI
        $script:WorkDays = Read-WorkDaysFromUI
        Set-AutoReplyState 'Scheduled'
        Update-StatusBar "Auto reply set to Scheduled"
        Show-InfoDialog "Done" "Auto Reply State set to Scheduled"
    }
    catch { Show-ErrorDialog "Error" $_.Exception.Message }
})

# Create Scheduled Task: Register a Windows Task Scheduler job to run this script daily.
$btnCreateTask.Add_Click({
    try {
        Read-ShiftTimesFromUI
        $result = Register-DailyScheduledTask
        if ($result) {
            Show-InfoDialog "Success" "Scheduled task 'AAOOF' created successfully."
            Update-StatusBar "Scheduled task created"
        } else {
            Show-InfoDialog "Info" "Scheduled task 'AAOOF' already exists."
            Update-StatusBar "Task already exists"
        }
    }
    catch {
        Show-ErrorDialog "Error" "Failed to create task.`n`n$($_.Exception.Message)"
        Update-StatusBar "Task creation failed"
    }
})

# Populate dynamic items: When the template dropdown opens, scan the config directory
# for backup files, saved messages, and custom_*.html files and insert them dynamically.
$cmbTemplate.Add_DropDownOpened({
    # Remove previously added dynamic items (identified by Tag)
    $dynamicItems = @($cmbTemplate.Items | Where-Object { $_.Tag -ne $null })
    foreach ($item in $dynamicItems) { $cmbTemplate.Items.Remove($item) }

    # Insert dynamic items before "Custom..."
    $customIdx = -1
    for ($i = 0; $i -lt $cmbTemplate.Items.Count; $i++) {
        if ($cmbTemplate.Items[$i].Content -eq "Custom...") { $customIdx = $i; break }
    }
    if ($customIdx -lt 0) { $customIdx = $cmbTemplate.Items.Count }

    # Check for message backup
    $bakFile = Join-Path $ConfigDir "message.html.bak"
    if (Test-Path $bakFile) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = "Last Message Backup"
        $item.Tag = $bakFile
        $cmbTemplate.Items.Insert($customIdx, $item)
        $customIdx++
    }

    # Check for saved online message
    $msgFile = Join-Path $ConfigDir "message.html"
    if (Test-Path $msgFile) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = "Saved Online Message"
        $item.Tag = $msgFile
        $cmbTemplate.Items.Insert($customIdx, $item)
        $customIdx++
    }

    # Check for any custom_*.html files
    $customFiles = Get-ChildItem -Path $ConfigDir -Filter "custom_*.html" -ErrorAction SilentlyContinue
    foreach ($f in $customFiles) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = "Custom: $($f.BaseName)"
        $item.Tag = $f.FullName
        $cmbTemplate.Items.Insert($customIdx, $item)
        $customIdx++
    }
})

# Auto-load template when dropdown selection changes
$cmbTemplate.Add_SelectionChanged({
    if ($null -eq $cmbTemplate.SelectedItem) { return }
    $selected = $cmbTemplate.SelectedItem.Content
    if ($selected -eq "Custom...") {
        Update-StatusBar "Use 'Browse File...' to load a custom template"
        return
    }
    $path = if ($cmbTemplate.SelectedItem.Tag) { $cmbTemplate.SelectedItem.Tag } else { Resolve-TemplateFilePath $selected }
    if ($path -and (Test-Path $path)) {
        # Save current message to backup before overwriting
        if (![string]::IsNullOrWhiteSpace($txtMessage.Text)) {
            $backupFile = Join-Path $ConfigDir "message.html.bak"
            Export-MessageToFile $backupFile $txtMessage.Text
        }
        $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $path -Raw)
        # Refresh preview if on Preview tab
        if ($tcMessageView.SelectedIndex -eq 1) {
            $wbPreview.NavigateToString($txtMessage.Text)
        }
        if ($script:IsConnectedToEXO) { $script:EXOMessageSynced = $false }
        Update-StatusBar "Template loaded: $selected"
    } else {
        Show-ErrorDialog "Not Found" "Template file not found: $path"
        Update-StatusBar "Template file not found"
    }
})

# Load Template button (also loads selected template)
$btnLoadTemplate.Add_Click({
    if ($null -eq $cmbTemplate.SelectedItem) { return }
    $selected = $cmbTemplate.SelectedItem.Content
    if ($selected -eq "Custom...") {
        Update-StatusBar "Use 'Browse File...' to load a custom template"
        return
    }
    $path = if ($cmbTemplate.SelectedItem.Tag) { $cmbTemplate.SelectedItem.Tag } else { Resolve-TemplateFilePath $selected }
    if ($path -and (Test-Path $path)) {
        if (![string]::IsNullOrWhiteSpace($txtMessage.Text)) {
            $backupFile = Join-Path $ConfigDir "message.html.bak"
            Export-MessageToFile $backupFile $txtMessage.Text
        }
        $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $path -Raw)
        if ($tcMessageView.SelectedIndex -eq 1) {
            $wbPreview.NavigateToString($txtMessage.Text)
        }
        if ($script:IsConnectedToEXO) { $script:EXOMessageSynced = $false }
        Update-StatusBar "Template loaded: $selected"
    } else {
        Show-ErrorDialog "Not Found" "Template file not found: $path"
        Update-StatusBar "Template file not found"
    }
})

# Re-resolve template when any option checkbox changes
$optionReloadHandler = {
    if ($null -eq $cmbTemplate.SelectedItem) { return }
    $selected = $cmbTemplate.SelectedItem.Content
    if ($selected -eq "Custom...") { return }
    $path = if ($cmbTemplate.SelectedItem.Tag) { $cmbTemplate.SelectedItem.Tag } else { Resolve-TemplateFilePath $selected }
    if ($path -and (Test-Path $path)) {
        $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $path -Raw)
        $wbPreview.NavigateToString($txtMessage.Text)
    }
}
# Wire option checkboxes: When any template option (signature, hours, work days,
# timezone) is toggled, re-resolve the template to reflect the change immediately.
$chkIncludeSignature.Add_Checked($optionReloadHandler)
$chkIncludeSignature.Add_Unchecked($optionReloadHandler)
$chkIncludeOfficeHours.Add_Checked($optionReloadHandler)
$chkIncludeOfficeHours.Add_Unchecked($optionReloadHandler)
$chkIncludeWorkDays.Add_Checked($optionReloadHandler)
$chkIncludeWorkDays.Add_Unchecked($optionReloadHandler)
$chkIncludeTimezone.Add_Checked($optionReloadHandler)
$chkIncludeTimezone.Add_Unchecked($optionReloadHandler)

# Re-resolve template when return date changes (updates [RETURN DATE] placeholder)
$dpReturnDate.Add_SelectedDateChanged($optionReloadHandler)

# Save and reload on Full Name / Role changes
$txtFullName.Add_TextChanged({
    $script:FullName = $txtFullName.Text
    Export-AppConfiguration
    & $optionReloadHandler
})
$txtRole.Add_TextChanged({
    $script:Role = $txtRole.Text
    Export-AppConfiguration
    & $optionReloadHandler
})

# Override Account checkbox: enable/disable account text box
$chkOverrideAccount.Add_Checked({
    $script:OverrideAccount = $true
    $txtAccount.IsEnabled = $true
    Export-AppConfiguration
})
$chkOverrideAccount.Add_Unchecked({
    $script:OverrideAccount = $false
    $txtAccount.IsEnabled = $false
    # Revert to auto-detected alias
    Resolve-UserAlias
    $txtAccount.Text = $script:UserAlias
    Export-AppConfiguration
    & $optionReloadHandler
})
# Save edited account when user tabs out
$txtAccount.Add_LostFocus({
    if ($chkOverrideAccount.IsChecked) {
        $script:UserAlias = $txtAccount.Text
        Export-AppConfiguration
        & $optionReloadHandler
    }
})

# ===================== HTML Formatting Toolbar Handlers =====================
# Helper: Wrap the selected text in the editor with an HTML tag, or insert at cursor.
function Insert-HtmlTag($openTag, $closeTag) {
    $selStart = $txtMessage.SelectionStart
    $selLen = $txtMessage.SelectionLength
    if ($selLen -gt 0) {
        $selected = $txtMessage.Text.Substring($selStart, $selLen)
        $replacement = "$openTag$selected$closeTag"
        $txtMessage.Text = $txtMessage.Text.Remove($selStart, $selLen).Insert($selStart, $replacement)
        $txtMessage.SelectionStart = $selStart
        $txtMessage.SelectionLength = $replacement.Length
    } else {
        $insert = "$openTag$closeTag"
        $txtMessage.Text = $txtMessage.Text.Insert($selStart, $insert)
        $txtMessage.SelectionStart = $selStart + $openTag.Length
    }
    $txtMessage.Focus()
}

# Helper: Insert a snippet at the cursor position.
function Insert-HtmlSnippet($snippet) {
    $selStart = $txtMessage.SelectionStart
    $txtMessage.Text = $txtMessage.Text.Insert($selStart, $snippet)
    $txtMessage.SelectionStart = $selStart + $snippet.Length
    $txtMessage.Focus()
}

$btnFmtBold.Add_Click({ Insert-HtmlTag '<b>' '</b>' })
$btnFmtItalic.Add_Click({ Insert-HtmlTag '<i>' '</i>' })
$btnFmtUnderline.Add_Click({ Insert-HtmlTag '<u>' '</u>' })
$btnFmtH3.Add_Click({ Insert-HtmlTag '<h3>' '</h3>' })
$btnFmtP.Add_Click({ Insert-HtmlTag '<p>' '</p>' })
$btnFmtBr.Add_Click({ Insert-HtmlSnippet '<br/>' })

$btnFmtLink.Add_Click({
    $selStart = $txtMessage.SelectionStart
    $selLen = $txtMessage.SelectionLength
    $linkText = if ($selLen -gt 0) { $txtMessage.Text.Substring($selStart, $selLen) } else { 'link text' }
    $snippet = "<a href=`"https://`">$linkText</a>"
    if ($selLen -gt 0) {
        $txtMessage.Text = $txtMessage.Text.Remove($selStart, $selLen).Insert($selStart, $snippet)
    } else {
        $txtMessage.Text = $txtMessage.Text.Insert($selStart, $snippet)
    }
    # Position cursor inside the href quotes
    $txtMessage.SelectionStart = $selStart + 9
    $txtMessage.SelectionLength = 8
    $txtMessage.Focus()
})

$btnFmtColor.Add_Click({
    $selStart = $txtMessage.SelectionStart
    $selLen = $txtMessage.SelectionLength
    $colorText = if ($selLen -gt 0) { $txtMessage.Text.Substring($selStart, $selLen) } else { 'text' }
    $snippet = "<span style='color: #0078D4;'>$colorText</span>"
    if ($selLen -gt 0) {
        $txtMessage.Text = $txtMessage.Text.Remove($selStart, $selLen).Insert($selStart, $snippet)
    } else {
        $txtMessage.Text = $txtMessage.Text.Insert($selStart, $snippet)
    }
    # Select the hex color so user can change it
    $txtMessage.SelectionStart = $selStart + 22
    $txtMessage.SelectionLength = 7
    $txtMessage.Focus()
})

$btnFmtList.Add_Click({
    $selStart = $txtMessage.SelectionStart
    $selLen = $txtMessage.SelectionLength
    if ($selLen -gt 0) {
        $selected = $txtMessage.Text.Substring($selStart, $selLen)
        $lines = $selected -split "`r?`n" | Where-Object { $_.Trim() -ne '' }
        $listItems = ($lines | ForEach-Object { "    <li>$($_.Trim())</li>" }) -join "`n"
        $snippet = "<ul>`n$listItems`n</ul>"
        $txtMessage.Text = $txtMessage.Text.Remove($selStart, $selLen).Insert($selStart, $snippet)
    } else {
        $snippet = "<ul>`n    <li></li>`n</ul>"
        $txtMessage.Text = $txtMessage.Text.Insert($selStart, $snippet)
        $txtMessage.SelectionStart = $selStart + 14
    }
    $txtMessage.Focus()
})

$btnFmtRef.Add_Click({
    $ref = @"
HTML Quick Reference:

TEXT FORMATTING
  <b>bold</b>             <i>italic</i>
  <u>underline</u>        <s>strikethrough</s>

STRUCTURE
  <p>paragraph</p>        <br/>  line break
  <h3>heading</h3>        <hr/>  horizontal rule

LISTS
  <ul>                     <ol>
    <li>bullet item</li>     <li>numbered item</li>
  </ul>                    </ol>

LINKS & IMAGES
  <a href="https://url">link text</a>

COLORS & STYLES (inline)
  <span style='color: #D83B01;'>colored text</span>
  <span style='font-size: 14pt;'>sized text</span>
  <p style='color: #555; font-size: 10pt;'>styled paragraph</p>

COMMON COLORS
  #0078D4 (blue)    #2E7D32 (green)    #D83B01 (red/orange)
  #FF8F00 (amber)   #555555 (gray)     #333333 (dark gray)
"@
    Show-InfoDialog "HTML Quick Reference" $ref
})

# Browse File: Open a file picker dialog to load a custom HTML template from disk.
$btnBrowseFile.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*"
    $dialog.InitialDirectory = $ConfigDir
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtMessage.Text = Get-Content $dialog.FileName -Raw
        if ($script:IsConnectedToEXO) { $script:EXOMessageSynced = $false }
        Update-StatusBar "Loaded message from $($dialog.FileName)"
    }
})

# Apply Internal Message: Push the editor text to Exchange as the Internal auto-reply.
$btnApplyInternal.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($txtMessage.Text)) {
            Show-ErrorDialog "Empty Message" "Please enter or load a message first."
            return
        }
        $warnings = Get-TemplateWarnings
        if ($warnings.Count -gt 0) {
            $result = [System.Windows.MessageBox]::Show("The following issues were found:`n`n$($warnings -join "`n")`n`nApply anyway?", "Template Warnings", 'YesNo', 'Warning')
            if ($result -ne 'Yes') { return }
        }
        Ensure-ExchangeConnection
        Update-StatusBar "Applying internal message..."
        Set-AutoReplyMessage $txtMessage.Text 'Internal'
        $script:EXOMessageSynced = $true
        Update-StatusBar "Internal message applied"
        Show-InfoDialog "Done" "Internal auto-reply message updated."
    }
    catch { Show-ErrorDialog "Error" $_.Exception.Message }
})

# Apply External Message: Push the editor text to Exchange as the External auto-reply.
$btnApplyExternal.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($txtMessage.Text)) {
            Show-ErrorDialog "Empty Message" "Please enter or load a message first."
            return
        }
        $warnings = Get-TemplateWarnings
        if ($warnings.Count -gt 0) {
            $result = [System.Windows.MessageBox]::Show("The following issues were found:`n`n$($warnings -join "`n")`n`nApply anyway?", "Template Warnings", 'YesNo', 'Warning')
            if ($result -ne 'Yes') { return }
        }
        Ensure-ExchangeConnection
        Update-StatusBar "Applying external message..."
        Set-AutoReplyMessage $txtMessage.Text 'External'
        $script:EXOMessageSynced = $true
        Update-StatusBar "External message applied"
        Show-InfoDialog "Done" "External auto-reply message updated."
    }
    catch { Show-ErrorDialog "Error" $_.Exception.Message }
})

# Apply Both Messages: Push the editor text as both Internal and External auto-reply.
$btnApplyBoth.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($txtMessage.Text)) {
            Show-ErrorDialog "Empty Message" "Please enter or load a message first."
            return
        }
        $warnings = Get-TemplateWarnings
        if ($warnings.Count -gt 0) {
            $result = [System.Windows.MessageBox]::Show("The following issues were found:`n`n$($warnings -join "`n")`n`nApply anyway?", "Template Warnings", 'YesNo', 'Warning')
            if ($result -ne 'Yes') { return }
        }
        Ensure-ExchangeConnection
        Update-StatusBar "Applying message to both internal and external..."
        Set-AutoReplyMessage $txtMessage.Text 'Both'
        $script:EXOMessageSynced = $true
        Update-StatusBar "Both messages applied"
        Show-InfoDialog "Done" "Internal and External auto-reply messages updated."
    }
    catch { Show-ErrorDialog "Error" $_.Exception.Message }
})

# Save Template: Write the current editor content to the selected template file on disk.
# If "Custom..." is selected, opens a Save dialog; otherwise overwrites the named template.
$btnSaveTemplate.Add_Click({
    $selected = $cmbTemplate.SelectedItem.Content
    if ($selected -eq "Custom...") {
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "HTML Files (*.html)|*.html"
        $dialog.InitialDirectory = $ConfigDir
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-MessageToFile $dialog.FileName $txtMessage.Text
            Update-StatusBar "Message saved to $($dialog.FileName)"
            Show-InfoDialog "Saved" "Message saved to $($dialog.FileName)"
        }
    } else {
        $path = Resolve-TemplateFilePath $selected
        if ($path) {
            Export-MessageToFile $path $txtMessage.Text
            Update-StatusBar "Template saved: $selected"
            Show-InfoDialog "Saved" "Template '$selected' updated."
        }
    }
})

# Save Online Message: Fetch the live auto-reply from Exchange, save it locally
# as message.html, and load it into the editor for review.
$btnSaveOnlineMsg.Add_Click({
    try {
        Update-StatusBar "Fetching current online message..."
        $arc = Get-AutoReplyConfiguration
        $msgFile = Join-Path $ConfigDir "message.html"
        Export-MessageToFile $msgFile $arc.ExternalMessage
        $txtMessage.Text = $arc.ExternalMessage
        Update-StatusBar "Online message saved to message.html"
        Show-InfoDialog "Saved" "Current online auto-reply message saved to:`n$msgFile"
    }
    catch {
        Show-ErrorDialog "Error" "Could not fetch message. Are you connected?`n$($_.Exception.Message)"
        Update-StatusBar "Failed to save online message"
    }
})

# ===================== HTML Preview Tab Handler =====================
# When the user switches to the Preview tab, render the current editor HTML
# in the embedded WebBrowser control for a live WYSIWYG preview.
$tcMessageView.Add_SelectionChanged({
    if ($tcMessageView.SelectedIndex -eq 1) {
        # Preview tab selected - render HTML
        $html = $txtMessage.Text
        if ([string]::IsNullOrWhiteSpace($html)) {
            $html = "<html><body><p style='color:#888;font-family:Segoe UI;'>No message to preview.</p></body></html>"
        }
        $wbPreview.NavigateToString($html)
    }
})

# ===================== Initialize UI =====================
# Apply saved configuration values to all controls before showing the window.
Initialize-UIFromConfig

# Populate holiday picker with upcoming US federal holidays
$today = (Get-Date).Date
$allHolidays = @(Get-USFederalHolidays -Year $today.Year) + @(Get-USFederalHolidays -Year ($today.Year + 1))
$upcomingHolidays = $allHolidays | Where-Object { $_.Date -ge $today } | Select-Object -First 12
$noneItem = New-Object System.Windows.Controls.ComboBoxItem
$noneItem.Content = "(Select a holiday...)"
$cmbHoliday.Items.Add($noneItem) | Out-Null
foreach ($h in $upcomingHolidays) {
    $item = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = "$($h.Name) — $($h.Date.ToString('MMMM d, yyyy'))"
    $item.Tag = $h
    $cmbHoliday.Items.Add($item) | Out-Null
}
$cmbHoliday.SelectedIndex = 0

# Holiday selection: set return date and holiday name when a holiday is chosen
$cmbHoliday.Add_SelectionChanged({
    if ($cmbHoliday.SelectedIndex -le 0) {
        $script:SelectedHolidayName = ""
        return
    }
    $holiday = $cmbHoliday.SelectedItem.Tag
    if ($null -ne $holiday) {
        $dpReturnDate.SelectedDate = $holiday.ReturnDate
        $script:SelectedHolidayName = $holiday.Name
        Update-StatusBar "Holiday: $($holiday.Name) — Return date set to $($holiday.ReturnDate.ToString('MMMM d, yyyy'))"
    }
})

# Load default template into message box
$defaultTemplate = Resolve-TemplateFilePath "Normal OOF"
if (Test-Path $defaultTemplate) {
    $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $defaultTemplate -Raw)
}

# ===================== Screenshot Capture (F12) =====================
# Capture screenshots of every tab for README documentation.
# Uses WPF RenderTargetBitmap to render each tab at screen DPI and save as PNG.
$ScreenshotsDir = Join-Path $ScriptDir "screenshots"

function Save-WindowScreenshot($filePath) {
    $Window.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    Start-Sleep -Milliseconds 200

    $source = [System.Windows.PresentationSource]::FromVisual($Window)
    $dpiX = $source.CompositionTarget.TransformToDevice.M11
    $dpiY = $source.CompositionTarget.TransformToDevice.M22

    $width = [int]($Window.ActualWidth * $dpiX)
    $height = [int]($Window.ActualHeight * $dpiY)

    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($width, $height, 96 * $dpiX, 96 * $dpiY, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($Window)

    $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $stream = [System.IO.File]::Create($filePath)
    $encoder.Save($stream)
    $stream.Close()
}

$Window.Add_KeyDown({
    if ($_.Key -eq 'F12') {
        try {
            if (!(Test-Path $ScreenshotsDir)) { New-Item -ItemType Directory -Path $ScreenshotsDir | Out-Null }
            Update-StatusBar "Capturing screenshots..."

            # Remember current tab positions to restore after
            $originalTab = $tcMain.SelectedIndex
            $originalSubTab = $tcMessageView.SelectedIndex

            # Tab 0: Quick Actions
            $tcMain.SelectedIndex = 0
            Save-WindowScreenshot (Join-Path $ScreenshotsDir "quick-actions.png")

            # Tab 1: Configuration
            $tcMain.SelectedIndex = 1
            Save-WindowScreenshot (Join-Path $ScreenshotsDir "configuration.png")

            # Tab 2: Message Templates — Edit
            $tcMain.SelectedIndex = 2
            $tcMessageView.SelectedIndex = 0
            Save-WindowScreenshot (Join-Path $ScreenshotsDir "message-templates-edit.png")

            # Tab 2: Message Templates — Preview
            $tcMessageView.SelectedIndex = 1
            Save-WindowScreenshot (Join-Path $ScreenshotsDir "message-templates-preview.png")

            # Tab 3: Current OOF
            $tcMain.SelectedIndex = 3
            Save-WindowScreenshot (Join-Path $ScreenshotsDir "current-oof.png")

            # Restore original tab positions
            $tcMessageView.SelectedIndex = $originalSubTab
            $tcMain.SelectedIndex = $originalTab

            Update-StatusBar "Screenshots saved to screenshots/ folder"
            Show-InfoDialog "Screenshots Captured" "5 screenshots saved to:`n$ScreenshotsDir`n`n- quick-actions.png`n- configuration.png`n- message-templates-edit.png`n- message-templates-preview.png`n- current-oof.png"
        }
        catch {
            Show-ErrorDialog "Screenshot Error" "Failed to capture screenshots:`n$($_.Exception.Message)"
            Update-StatusBar "Screenshot capture failed"
        }
    }
})

# ===================== Show the Window =====================
# Display the WPF window (blocks execution until closed).
# On close, disconnect Exchange Online to release the session.
# Warn if there are unapplied message changes while connected.
$Window.Add_Closing({
    if ($script:IsConnectedToEXO -and -not $script:EXOMessageSynced) {
        $result = [System.Windows.MessageBox]::Show(
            "Your template message has not been applied to Exchange Online.`n`nAre you sure you want to exit without applying?",
            "Unapplied Changes",
            'YesNo',
            'Warning'
        )
        if ($result -ne 'Yes') {
            $_.Cancel = $true
        }
    }
})
$Window.Add_Closed({
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
})
$Window.ShowDialog() | Out-Null

# Final cleanup: ensure Exchange session is released even if the Closed event didn't fire.
try { Disconnect-ExchangeOnlineSession } catch { }
