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

.PARAMETER VersionInfo
    Show version information and exit:
    - Local script version.
    - Current GitHub version from the update URL.

    Aliases: -v, -version

.NOTES
    Requires : ExchangeOnlineManagement module (prompted to install if missing).
    Config   : config/config.json (auto-created, .gitignored).
    Templates: config/*.html (auto-downloaded from GitHub if missing).
    XAML     : config/AAOOF-GUI.xaml (UI layout, auto-downloaded if missing).
#>
param(
    [string]$InputParameter,
    [Alias('v', 'version')]
    [switch]$VersionInfo,
    [switch]$DisableTemplateAutoDownload,
    [switch]$DisableAutoUpdate,
    [switch]$DisableAutoUpdateRestart,
    [switch]$UseRootConfig
)

# ===================== WPF GUI Setup =====================
# Load .NET assemblies required for WPF windows, controls, and file dialogs.
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Resolve paths for the script directory, config folder, config file, and XAML layout.
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigDir = Join-Path $ScriptDir "config"
if (!(Test-Path $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir | Out-Null }

if ($UseRootConfig) {
    $ConfigFile = Join-Path $ScriptDir "config.json"
}
else {
    $ConfigFile = Join-Path $ConfigDir "config.json"
}
$XamlFile = Join-Path $ConfigDir "AAOOF-GUI.xaml"

# Ensure config directory exists (idempotent safety check).
if (!(Test-Path $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir | Out-Null }

# ===================== Auto-Download Missing Config Files =====================
# On first run, download XAML layout and HTML templates from the GitHub repository
# so that the tool works out of the box without manual file setup.
$RepoBaseUrl = "https://raw.githubusercontent.com/MicrosoftAzureAaron/DailyOOF/main/config"
$ScriptUpdateUrl = "https://raw.githubusercontent.com/MicrosoftAzureAaron/DailyOOF/main/AAOOF-GUI.ps1"
$script:ScriptVersion = "1.9.17" # Increment this with each release to trigger update checks
$DefaultConfigFiles = @(
    "AAOOF-GUI.xaml",
    "normal_oof.html",
    "vacation_oof.html",
    "sick_oof.html",
    "holiday_oof.html",
    "placeholder_examples.html"
)

# Default to enabled before config is loaded — Import-AppConfiguration will override
# this later if the user has explicitly disabled it in config.json. This ensures
# first-run (no config file) always attempts to download missing files.
$script:EnableTemplateAutoDownload = $true

$downloadFailures = @()
$downloadSkipped = @()

foreach ($fileName in $DefaultConfigFiles) {
    $localPath = Join-Path $ConfigDir $fileName
    if (!(Test-Path $localPath)) {
        if (-not $script:EnableTemplateAutoDownload) {
            $downloadSkipped += $fileName
            continue
        }
        $url = "$RepoBaseUrl/$fileName"
        try {
            Invoke-WebRequest -Uri $url -OutFile $localPath -UseBasicParsing -Headers @{ 'Cache-Control' = 'no-cache' }
            Write-Host "Downloaded missing file: $fileName" -ForegroundColor Green
        }
        catch {
            Write-Host "Warning: Could not download $fileName from $url" -ForegroundColor Yellow
            $downloadFailures += $fileName
        }
    }
}

if ($downloadSkipped.Count -gt 0) {
    $skippedList = $downloadSkipped -join "`n  - "
    $msg = "Template auto-download is disabled. The following files were not downloaded:`n  - $skippedList`n`n" +
    "If you want templates or XAML restored, enable auto-download in config.json or run with -DisableTemplateAutoDownload removed."
    Write-Host $msg -ForegroundColor Yellow
}

if ($downloadFailures.Count -gt 0) {
    $failedList = $downloadFailures -join "`n  - "
    $msg = "The following config files could not be downloaded:`n  - $failedList`n`n" +
    "You can clone the full repository instead:`n`n" +
    "git clone https://github.com/MicrosoftAzureAaron/DailyOOF.git`n`n" +
    "Then run the script from the cloned folder."
    Write-Host $msg -ForegroundColor Yellow
    [System.Windows.MessageBox]::Show(
        $msg,
        "AAOOF - Download Failed",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Warning
    ) | Out-Null
}

# ===================== Self-Update Check =====================
# Get-RemoteScriptVersion: Download the remote script to a temp file and extract its
# $script:ScriptVersion value. Returns the version string or 'unknown'.
function Get-RemoteScriptVersion {
    try {
        $tempFile = [System.IO.Path]::GetTempFileName()
        Invoke-WebRequest -Uri $ScriptUpdateUrl -OutFile $tempFile -UseBasicParsing -TimeoutSec 10 -Headers @{ 'Cache-Control' = 'no-cache' }
        $line = Select-String -Path $tempFile -Pattern '^\$script:ScriptVersion\s*=\s*"(.+)"' | Select-Object -First 1
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        if ($line) { return $line.Matches[0].Groups[1].Value }
        return 'unknown'
    }
    catch { return 'unknown' }
}

# Test-IsRemoteVersionNewer: Return $true only when remote version is strictly greater than local.
function Test-IsRemoteVersionNewer {
    param(
        [string]$RemoteVersion,
        [string]$LocalVersion
    )

    if ([string]::IsNullOrWhiteSpace($RemoteVersion) -or $RemoteVersion -eq 'unknown') {
        return $false
    }

    try {
        $remoteParsed = [version]$RemoteVersion
        $localParsed = [version]$LocalVersion
        return ($remoteParsed -gt $localParsed)
    }
    catch {
        return $false
    }
}

# Get-UpdateVersionState: Classify update comparison for clearer user-facing messaging.
function Get-UpdateVersionState {
    param(
        [string]$RemoteVersion,
        [string]$LocalVersion
    )

    if ([string]::IsNullOrWhiteSpace($RemoteVersion) -or $RemoteVersion -eq 'unknown') {
        return 'Unknown'
    }
    if (Test-IsRemoteVersionNewer -RemoteVersion $RemoteVersion -LocalVersion $LocalVersion) {
        return 'RemoteNewer'
    }
    if (Test-IsRemoteVersionNewer -RemoteVersion $LocalVersion -LocalVersion $RemoteVersion) {
        return 'LocalNewer'
    }
    return 'UpToDate'
}

# Invoke-ScriptSelfUpdateExternal: Start a separate PowerShell process to download the
# latest script and XAML for this running copy.
function Invoke-ScriptSelfUpdateExternal([string]$InputParam = "") {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $psExe = Join-Path $PSHOME 'pwsh.exe'
    }
    else {
        $psExe = Join-Path $PSHOME 'powershell.exe'
    }
    if (-not (Test-Path $psExe)) {
        throw "PowerShell executable not found at '$psExe'. Cannot perform external update."
    }

    $inputArg = if ([string]::IsNullOrEmpty($InputParam)) { '' } else { $InputParam }
    $scriptPath = $PSCommandPath
    $xamlUrl = "$RepoBaseUrl/AAOOF-GUI.xaml"
    $xamlPath = $XamlFile

    $childScript = @"
`$targetScript = "$scriptPath"
`$scriptUrl = "$ScriptUpdateUrl"
`$xamlUrl = "$xamlUrl"
`$xamlPath = "$xamlPath"
`$inputArg = "$inputArg"
`$tempFile = [System.IO.Path]::GetTempFileName()
Invoke-WebRequest -Uri `$scriptUrl -OutFile `$tempFile -UseBasicParsing -TimeoutSec 15 -Headers @{ 'Cache-Control' = 'no-cache' }
while (`$true) {
    try {
        Copy-Item -Path `$tempFile -Destination `$targetScript -Force
        break
    } catch {
        Start-Sleep -Seconds 1
    }
}
Remove-Item -Path `$tempFile -Force -ErrorAction SilentlyContinue
try {
    Invoke-WebRequest -Uri `$xamlUrl -OutFile `$xamlPath -UseBasicParsing -TimeoutSec 10 -Headers @{ 'Cache-Control' = 'no-cache' }
} catch {
}
`$argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", `$targetScript)
if (-not [string]::IsNullOrEmpty(`$inputArg)) { `$argList += `$inputArg }
# The external updater only replaces the local script and XAML.
# The running script exits after showing the update success message.
"@

    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($childScript))
    Start-Process -FilePath $psExe -ArgumentList @("-NoProfile", "-EncodedCommand", $encoded) -WindowStyle Hidden
    return $true
}

# Get-MissingHeadlessFiles: List required config/template files missing for CLI mode.
function Get-MissingHeadlessFiles {
    $missing = @()
    if (-not (Test-Path $ConfigFile)) {
        $missing += "Configuration file missing: $ConfigFile"
    }
    foreach ($fileName in $DefaultConfigFiles) {
        $localPath = Join-Path $ConfigDir $fileName
        if (-not (Test-Path $localPath)) {
            $missing += "Template file missing: $fileName"
        }
    }
    return $missing
}

# Write-MissingHeadlessFiles: Print missing-file warnings for headless execution.
function Write-MissingHeadlessFiles {
    $missing = Get-MissingHeadlessFiles
    if ($missing.Count -gt 0) {
        Write-Host "Required files are missing for headless/scheduled mode:" -ForegroundColor Yellow
        foreach ($item in $missing) {
            Write-Host "  - $item" -ForegroundColor Yellow
        }
        Write-Host "Run the GUI once to restore missing files and save your configuration." -ForegroundColor Yellow
    }
}

# Confirm-HeadlessConfigAvailable: Validate minimum config exists for CLI mode.
function Confirm-HeadlessConfigAvailable {
    if (-not (Test-Path $ConfigFile)) {
        Write-Host "Configuration file not found: $ConfigFile" -ForegroundColor Red
        Write-Host "Please run the GUI once and save your settings before using scheduled or headless mode." -ForegroundColor Red
        return $false
    }
    return $true
}

# ===================== Configuration (loaded from config.json) =====================
# Global variables hold the user's settings. Defaults are set here and then
# overwritten by Import-AppConfiguration if a config.json file exists.
$script:StartOfShift = $null                       # Shift start time (datetime)
$script:EndOfShift = $null                       # Shift end time (datetime)
$script:WorkDays = $null                       # Array of day names, e.g. @('Monday','Tuesday',...)
$script:UserAlias = ""                           # Email address used as Exchange identity
$script:UserAliasSuffix = ""                           # Domain suffix appended to the Windows username
$script:FullName = ""                           # Display name for auto-generated signature
$script:Role = ""                           # Job title inserted into templates via [ROLE]
$script:BackupContact = ""                           # Contact person or mailbox used in [BACKUP CONTACT]
$script:BackupEngineerEmail = ""                     # Email address of backup engineer for [BACKUP ENGINEER EMAIL]
$script:BackupEmail = ""                            # Email address for template backups
$script:TeamAlias = ""                           # Team name or alias used in [TEAM ALIAS]
$script:SupportLink = ""                           # URL used in [SUPPORT LINK]
$script:OverrideAccount = $false                      # True if user manually overrides the account email
$script:SelectedHolidayName = ""                      # Name of the selected holiday for [HOLIDAY NAME] placeholder
$script:EnableTemplateAutoDownload = $true             # Automatically download missing templates/XAML on startup
$script:EnableAutoUpdateCheck = $true                   # Background check for script updates when GUI starts
$script:EnableAutoUpdateRestart = $false                # Restart the GUI automatically after applying an update
$script:UseRootConfig = $false                          # Store config.json alongside the script instead of in config/ folder
$script:TaskStartOffsetMinutes = 15                     # Minutes after shift start to run the daily scheduled task

# Track EXO sync state for status/UI updates.
$script:IsConnectedToEXO = $false
$script:OOFReplyEnabled = $true

# ConvertTo-UserAliasSuffix: Normalize a configured email suffix.
function ConvertTo-UserAliasSuffix($suffix) {
    if ([string]::IsNullOrEmpty($suffix)) { return $suffix }
    $normalized = $suffix.Trim()
    if ($normalized.StartsWith('@')) { $normalized = $normalized.Substring(1) }
    $normalized = $normalized.ToLower()
    if ($normalized -match '\.?microsoft\.com$') { return '@microsoft.com' }
    return "@$normalized"
}

# Import-AppConfiguration: Read config.json and populate global variables.
function Import-AppConfiguration {
    if (Test-Path $ConfigFile) {
        $cfg = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        if ($cfg.StartOfShift) { $script:StartOfShift = [datetime]$cfg.StartOfShift }
        if ($cfg.EndOfShift) { $script:EndOfShift = [datetime]$cfg.EndOfShift }
        if ($cfg.WorkDays) { $script:WorkDays = @($cfg.WorkDays) }
        if ($cfg.UserAlias) { $script:UserAlias = $cfg.UserAlias }
        if ($cfg.UserAliasSuffix) { $script:UserAliasSuffix = ConvertTo-UserAliasSuffix($cfg.UserAliasSuffix) }
        if ($cfg.FullName) { $script:FullName = $cfg.FullName }
        if ($cfg.Role) { $script:Role = $cfg.Role }
        if ($cfg.BackupContact) { $script:BackupContact = $cfg.BackupContact }
        if ($cfg.BackupEngineerEmail) { $script:BackupEngineerEmail = $cfg.BackupEngineerEmail }
        if ($cfg.BackupEmail) { $script:BackupEmail = $cfg.BackupEmail }
        if ($cfg.TeamAlias) { $script:TeamAlias = $cfg.TeamAlias }
        if ($cfg.SupportLink) { $script:SupportLink = $cfg.SupportLink }
        if ($null -ne $cfg.OverrideAccount) { $script:OverrideAccount = [bool]$cfg.OverrideAccount }
        if ($null -ne $cfg.EnableTemplateAutoDownload) { $script:EnableTemplateAutoDownload = [bool]$cfg.EnableTemplateAutoDownload }
        if ($null -ne $cfg.EnableAutoUpdateCheck) { $script:EnableAutoUpdateCheck = [bool]$cfg.EnableAutoUpdateCheck }
        if ($null -ne $cfg.EnableAutoUpdateRestart) { $script:EnableAutoUpdateRestart = [bool]$cfg.EnableAutoUpdateRestart }
        if ($null -ne $cfg.UseRootConfig) { $script:UseRootConfig = [bool]$cfg.UseRootConfig }
        if ($null -ne $cfg.TaskStartOffsetMinutes) { $script:TaskStartOffsetMinutes = [int]$cfg.TaskStartOffsetMinutes }
    }
}

# Load config immediately on script start.
# Detect first run BEFORE loading config — if no config.json exists this is a fresh install.
$script:IsFirstRun = !(Test-Path $ConfigFile)
Import-AppConfiguration

# Apply startup switches and persist them if requested.
$startupSwitchesApplied = $false
if ($DisableTemplateAutoDownload) {
    $script:EnableTemplateAutoDownload = $false
    $startupSwitchesApplied = $true
    Write-Host "Template auto-download disabled by startup parameter and will be saved to config file." -ForegroundColor Yellow
}
if ($DisableAutoUpdate) {
    $script:EnableAutoUpdateCheck = $false
    $startupSwitchesApplied = $true
    Write-Host "Auto-update checks disabled by startup parameter and will be saved to config file." -ForegroundColor Yellow
}
if ($DisableAutoUpdateRestart) {
    $script:EnableAutoUpdateRestart = $false
    $startupSwitchesApplied = $true
    Write-Host "Auto-update restart disabled by startup parameter and will be saved to config file." -ForegroundColor Yellow
}
if ($UseRootConfig) {
    $script:UseRootConfig = $true
    $startupSwitchesApplied = $true
    Write-Host "Using root config file location due to startup parameter." -ForegroundColor Yellow
}
if ($startupSwitchesApplied) {
    Export-AppConfiguration
}

# CLI version output mode: print local and GitHub versions, then exit without launching GUI.
if ($VersionInfo) {
    $remoteVer = Get-RemoteScriptVersion
    Write-Host "Local version : v$($script:ScriptVersion)"
    if ($remoteVer -eq 'unknown') {
        Write-Host "GitHub version: unknown (could not reach update URL)" -ForegroundColor Yellow
    }
    else {
        Write-Host "GitHub version: v$remoteVer"
    }
    exit 0
}

# Test-IsAdmin: Return $true when the current process is elevated.
function Test-IsAdmin {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

# Get-InstallCommand: Return the appropriate Install-Module command for the current context.
function Get-InstallCommand {
    if (Test-IsAdmin) {
        return "Install-Module -Name ExchangeOnlineManagement -Force"
    }
    else {
        return "Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser"
    }
}

# Test-ExchangeOnlineModule: Ensure ExchangeOnlineManagement is installed and available.
function Test-ExchangeOnlineModule {
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        $installCmd = Get-InstallCommand
        try {
            Update-StatusBar "Installing ExchangeOnlineManagement module..."
        }
        catch { }

        try {
            if (Test-IsAdmin) {
                Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -ErrorAction Stop
            }
            else {
                Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            }
        }
        catch {
            $msg = "The ExchangeOnlineManagement module is required but could not be installed automatically.`n`n" +
            "Install it manually by running:`n`n" +
            "$installCmd`n`n" +
            "Then restart the application.`n`n" +
            "Install error: $($_.Exception.Message)"
            throw $msg
        }

        if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            $msg = "The ExchangeOnlineManagement module is required but still not available after installation.`n`n" +
            "Try installing manually:`n`n" +
            "$installCmd`
Then restart the application."
            throw $msg
        }
    }
}

# For GUI mode, check module on launch
if (!$InputParameter) {
    try {
        Test-ExchangeOnlineModule
    }
    catch {
        [System.Windows.MessageBox]::Show(
            $_.Exception.Message,
            "Module Not Found",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        exit
    }
}

# Test-ValidEmailAddress: Return $true when a mailbox address is syntactically valid.
function Test-ValidEmailAddress($EmailAddress) {
    if ([string]::IsNullOrWhiteSpace($EmailAddress)) { return $false }
    try {
        [void][System.Net.Mail.MailAddress]::new($EmailAddress)
        return $true
    }
    catch {
        return $false
    }
}

# Get-ConfigurationValidationErrors: Collect user-facing validation errors for key actions.
function Get-ConfigurationValidationErrors {
    param(
        [switch]$RequireShiftTimes,
        [switch]$RequireWorkDays,
        [switch]$RequireOverrideEmail,
        [switch]$RequireFutureReturnDate,
        [switch]$RequireTaskOffset
    )

    $errors = @()

    if ($RequireOverrideEmail -and $chkOverrideAccount.IsChecked) {
        if (-not (Test-ValidEmailAddress $txtAccount.Text)) {
            $errors += "Override Account must be a valid email address."
        }
    }

    if ($RequireShiftTimes) {
        Read-ShiftTimesFromUI
        if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift) {
            $errors += "Shift start and end times are required."
        }
        elseif ($script:StartOfShift -ge $script:EndOfShift) {
            $errors += "Shift start must be earlier than shift end."
        }
    }

    if ($RequireWorkDays) {
        $script:WorkDays = Read-WorkDaysFromUI
        if (-not $script:WorkDays -or $script:WorkDays.Count -eq 0) {
            $errors += "Select at least one work day."
        }
    }

    if ($RequireFutureReturnDate) {
        if ($null -eq $dpReturnDate.SelectedDate) {
            $errors += "Please select a return date."
        }
        elseif ($dpReturnDate.SelectedDate.Date -lt (Get-Date).Date) {
            $errors += "Return date cannot be in the past."
        }
    }

    if ($RequireTaskOffset) {
        $parsedOffset = 0
        if (-not [int]::TryParse($txtTaskOffsetMinutes.Text, [ref]$parsedOffset)) {
            $errors += "Task start offset must be a whole number of minutes."
        }
        elseif ($parsedOffset -lt 0 -or $parsedOffset -gt 180) {
            $errors += "Task start offset must be between 0 and 180 minutes."
        }
        else {
            $script:TaskStartOffsetMinutes = $parsedOffset
        }
    }

    return $errors
}

# Assert-ConfigurationValid: Throw a single validation error when prerequisites are missing.
function Assert-ConfigurationValid {
    param(
        [switch]$RequireShiftTimes,
        [switch]$RequireWorkDays,
        [switch]$RequireOverrideEmail,
        [switch]$RequireFutureReturnDate,
        [switch]$RequireTaskOffset
    )

    $errors = Get-ConfigurationValidationErrors `
        -RequireShiftTimes:$RequireShiftTimes `
        -RequireWorkDays:$RequireWorkDays `
        -RequireOverrideEmail:$RequireOverrideEmail `
        -RequireFutureReturnDate:$RequireFutureReturnDate `
        -RequireTaskOffset:$RequireTaskOffset

    if ($errors.Count -gt 0) {
        throw ($errors -join "`n")
    }
}

# ===================== Core Functions =====================

# Resolve-UserAlias: Build the user's email alias from the Windows login name + suffix.
function Resolve-UserAlias {
    if ([string]::IsNullOrEmpty($script:UserAliasSuffix)) {
        # Try to derive suffix from the machine's DNS domain.
        # Normalize Microsoft corporate domains to @microsoft.com for internal users.
        if ($env:USERDNSDOMAIN) {
            $dnsDomain = $env:USERDNSDOMAIN.ToLower()
            if ($dnsDomain -match '\.?microsoft\.com$') {
                $script:UserAliasSuffix = '@microsoft.com'
            }
            else {
                $script:UserAliasSuffix = "@$dnsDomain"
            }
        }
    }
    $ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    if ($ComputerSystem.Username) {
        $CurrentUser = $ComputerSystem.Username.Split('\')[-1]
    }
    else {
        $CurrentUser = $env:USERNAME
    }
    $script:UserAlias = "$CurrentUser$script:UserAliasSuffix"
}

# Resolve-ProfileFromEXO: Query EXO for profile fields and populate any that are currently blank.
# Pulls DisplayName (-> FullName) from Get-EXOMailbox and Title (-> Role) from Get-EXORecipient.
# Only fills fields that are blank — never overwrites user-entered values.
# Returns $true if FullName was resolved, $false if it still needs manual input.
function Resolve-ProfileFromEXO {
    $nameResolved = $false
    $anyChange = $false

    try {
        # --- Full Name from mailbox DisplayName ---
        if ([string]::IsNullOrWhiteSpace($script:FullName)) {
            $mbx = Get-EXOMailbox -Identity $script:UserAlias -ErrorAction Stop
            if (![string]::IsNullOrWhiteSpace($mbx.DisplayName)) {
                $script:FullName = $mbx.DisplayName
                $txtFullName.Text = $mbx.DisplayName
                $anyChange = $true
                $nameResolved = $true
                Write-Host "Full name resolved from EXO: $($mbx.DisplayName)" -ForegroundColor Green
            }
        }
        else {
            $nameResolved = $true  # already set by user
        }
    }
    catch {
        Write-Host "Could not resolve display name from EXO: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    try {
        # --- Role from recipient Title ---
        if ([string]::IsNullOrWhiteSpace($script:Role)) {
            $recip = Get-EXORecipient -Identity $script:UserAlias -Properties Title -ErrorAction Stop
            if (![string]::IsNullOrWhiteSpace($recip.Title)) {
                $script:Role = $recip.Title
                $txtRole.Text = $recip.Title
                $anyChange = $true
                Write-Host "Role resolved from EXO: $($recip.Title)" -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Host "Could not resolve role/title from EXO: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    if ($anyChange) { Export-AppConfiguration }

    return $nameResolved
}

# Show-NameInputDialog: Prompt the user to enter their display name manually.
# Saves the result to config and updates the UI field. Skips silently if the user cancels.
function Show-NameInputDialog {
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
    $nameInput = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Your display name could not be retrieved from Exchange Online.`n`nEnter your full name for use in OOF message signatures:`n(You can also set this later in Configuration > Profile > Full Name)",
        "Enter Your Name",
        ""
    )
    if (![string]::IsNullOrWhiteSpace($nameInput)) {
        $script:FullName = $nameInput.Trim()
        $txtFullName.Text = $script:FullName
        Export-AppConfiguration
        Update-StatusBar "Name saved: $($script:FullName)"
    }
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

# Set-AutoReplyState: Change the auto-reply state (Enabled|Disabled|Scheduled) on Exchange.
function Set-AutoReplyState($State) {
    switch ($State) {
        'Enabled' { Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState "Enabled" }
        'Disabled' { Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState "Disabled" }
        'Scheduled' {
            # Calculate times first, then set state + times in a single atomic call
            # to avoid a race where Exchange activates scheduling with stale times.
            $schedTimes = Get-AutoReplyScheduleTimes
            if ($null -eq $schedTimes) {
                throw "Cannot enable scheduled auto reply: shift times or work days are not configured."
            }
            if ($schedTimes.StartTime -ge $schedTimes.EndTime) {
                throw "Cannot enable scheduled auto reply: calculated OOF start ($($schedTimes.StartTime)) is not before end ($($schedTimes.EndTime)). Check your shift times and work days."
            }
            Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias `
                -AutoReplyState "Scheduled" `
                -StartTime $schedTimes.StartTime `
                -EndTime $schedTimes.EndTime
        }
    }
    Save-AutoReplyConfigToFile
}

# Get-AutoReplyScheduleTimes: Calculate OOF start/end times based on shift and work days.
# Returns a hashtable with StartTime (OOF begins = end of shift) and EndTime (OOF ends = next shift start),
# or $null if configuration is incomplete.
function Get-AutoReplyScheduleTimes {
    if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift) { return $null }
    if ($null -eq $script:WorkDays) { return $null }

    $DaysToAdd = Get-NextWorkDayOffset

    # EndTime of OOF = start of next work day shift
    $OofEndTime = (Get-Date).Date.Add($script:StartOfShift.TimeOfDay).AddDays($DaysToAdd)

    # StartTime of OOF = end of shift (today, or yesterday if before shift start)
    $OofStartTime = (Get-Date).Date.Add($script:EndOfShift.TimeOfDay)
    if ($DaysToAdd -eq 0) { $OofStartTime = $OofStartTime.AddDays(-1) }

    return @{ StartTime = $OofStartTime; EndTime = $OofEndTime }
}

# Set-AutoReplyMessage: Apply an HTML message body as the auto-reply for Internal, External, or Both.
function Set-AutoReplyMessage($Message, $MessageScope) {
    switch ($MessageScope) {
        'Internal' { Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -InternalMessage $Message }
        'External' { Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -ExternalMessage $Message }
        default { Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -ExternalMessage $Message -InternalMessage $Message }
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
        if ($CurrentTime.TimeOfDay -lt $script:StartOfShift.TimeOfDay) {
            return 0
        }
        else {
            return 1
        }
    }
}

# Connect-ExchangeOnlineSession: Ensure a live Exchange Online connection exists.
# Throws error if EXO module is missing, reuses existing sessions if found.
function Connect-ExchangeOnlineSession {
    if ([string]::IsNullOrEmpty($script:UserAlias)) { Resolve-UserAlias }

    # Check if ExchangeOnlineManagement module is available
    Test-ExchangeOnlineModule

    # Ensure the module is loaded before calling any of its commands
    if (!(Get-Module -Name ExchangeOnlineManagement)) {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
    }

    $session = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -ne $session) {
        $exchangeSession = $session | Where-Object { $_.Name -like "ExchangeOnline_*" }
        if ($null -ne $exchangeSession) {
            return $true
        }
    }

    # Flush UI before blocking on auth so status messages are visible
    if ($null -ne $Window) {
        $Window.Dispatcher.Invoke([action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
    }
    Connect-ExchangeOnline -UserPrincipalName $script:UserAlias -ShowBanner:$false -CommandName Get-MailboxAutoReplyConfiguration, Set-MailboxAutoReplyConfiguration, Get-EXOMailbox, Get-EXORecipient
    return $true
}

# Disconnect-ExchangeOnlineSession: Safely tear down the Exchange Online connection.
function Disconnect-ExchangeOnlineSession {
    try { Disconnect-ExchangeOnline -Confirm:$false } catch { }
}

# Show-ConnectingWindow: Display a small "Connecting..." progress window with a live elapsed
# timer on a dedicated STA runspace so it keeps updating while the main UI thread is blocked
# on Connect-ExchangeOnline. Returns a context hashtable to pass to Close-ConnectingWindow.
# The window includes a Cancel button; check $ctx.SyncHash.Cancelled after Close-ConnectingWindow
# returns to detect user-requested cancellation.
function Show-ConnectingWindow {
    $syncHash = [System.Collections.Hashtable]::Synchronized(@{
        Done      = $false
        Cancelled = $false
        Window    = $null
    })

    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
    $runspace.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable('syncHash', $syncHash)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $runspace
    [void]$ps.AddScript({
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Connecting" Width="320" Height="150"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        WindowStyle="ToolWindow" Topmost="True">
    <StackPanel Margin="20" VerticalAlignment="Center">
        <TextBlock Text="Connecting to Exchange Online..." FontSize="13"
                   HorizontalAlignment="Center" TextWrapping="Wrap"/>
        <TextBlock Name="lblElapsed" Text="0s elapsed" FontSize="11" Foreground="Gray"
                   HorizontalAlignment="Center" Margin="0,8,0,12"/>
        <Button Name="btnCancel" Content="Cancel" Width="90" HorizontalAlignment="Center"
                Padding="8,4" FontSize="12"/>
    </StackPanel>
</Window>
'@
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
        $win    = [System.Windows.Markup.XamlReader]::Load($reader)
        $lbl    = $win.FindName('lblElapsed')
        $btn    = $win.FindName('btnCancel')
        $syncHash.Window = $win

        $btn.Add_Click({
            $syncHash.Cancelled = $true
            $syncHash.Done      = $true
            $win.Close()
        })

        $startTime = [datetime]::Now
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [System.TimeSpan]::FromMilliseconds(500)
        $timer.Add_Tick(( {
            if ($syncHash.Done) {
                $timer.Stop()
                $win.Close()
                return
            }
            $elapsed = [int]([datetime]::Now - $startTime).TotalSeconds
            $lbl.Text = "${elapsed}s elapsed"
        } ).GetNewClosure())
        $timer.Start()
        $win.ShowDialog() | Out-Null
    })

    $handle = $ps.BeginInvoke()

    # Wait until the window object is set (up to 3s) before returning.
    $waited = 0
    while ($null -eq $syncHash.Window -and $waited -lt 3000) {
        Start-Sleep -Milliseconds 50
        $waited += 50
    }

    return @{ SyncHash = $syncHash; PS = $ps; Runspace = $runspace; Handle = $handle }
}

# Close-ConnectingWindow: Signal the connecting window to close and clean up the runspace.
function Close-ConnectingWindow {
    param($ctx)
    if ($null -eq $ctx) { return }
    $ctx.SyncHash.Done = $true
    try { $null = $ctx.PS.EndInvoke($ctx.Handle) } catch {}
    try { $ctx.PS.Dispose() }      catch {}
    try { $ctx.Runspace.Close(); $ctx.Runspace.Dispose() } catch {}
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

# Get-TemplateWarnings: Check for unresolved placeholders and missing profile config.
# Returns an array of user-facing warning strings with fix guidance.
#
# Profile advisory warnings (Role, Backup Contact, etc.) only fire when the config field is
# blank AND the message still contains the silent fallback string that Resolve-TemplatePlaceholders
# inserted. This prevents false positives when the user's template never used that placeholder,
# or when the user has deliberately written a message that doesn't rely on those fields.
function Get-TemplateWarnings {
    $warnings = @()
    $msg = $txtMessage.Text
    if ([string]::IsNullOrWhiteSpace($msg)) { return $warnings }

    # --- Residual token scan: catch any [TOKEN] that Resolve-TemplatePlaceholders left literal ---
    # (These only remain when the prerequisite value/selection is genuinely missing.)
    if ($msg -match '\[RETURN DATE\]') {
        $warnings += "[RETURN DATE] was not replaced — Fix: select a return date in Quick Actions > Vacation OOF."
    }
    if ($msg -match '\[HOLIDAY NAME\]') {
        $warnings += "[HOLIDAY NAME] was not replaced — Fix: select a holiday in Quick Actions > Holiday OOF."
    }
    if ($msg -match '\[SIGNATURE\]') {
        $warnings += "[SIGNATURE] was not resolved — Fix: ensure Include Signature is checked on the Message Templates tab."
    }

    # Catch-all: any remaining [ALL CAPS TOKEN] not already flagged above.
    $knownResiduals = @('RETURN DATE', 'HOLIDAY NAME', 'SIGNATURE')
    $remaining = [regex]::Matches($msg, '\[[A-Z][A-Z ]+\]') |
        ForEach-Object { $_.Value -replace '[\[\]]', '' } | Select-Object -Unique
    foreach ($token in $remaining) {
        if ($token -notin $knownResiduals) {
            $warnings += "[$token] was not replaced — check Configuration or remove this placeholder from the message."
        }
    }

    # --- Profile advisory checks ---
    # Each check only fires when BOTH conditions are true:
    #   1. The config field is blank (so a silent fallback was used during resolution)
    #   2. The fallback value is present in the message (confirming the placeholder was actually used)
    # This avoids warning when the user wrote a message that never needed that placeholder.

    # Full Name: warn only when name was derived from the alias (field was blank)
    if ([string]::IsNullOrWhiteSpace($txtFullName.Text)) {
        $warnings += "Full Name is not set — your name was derived from your account alias. Fix: Configuration > Profile > Full Name."
    }

    # Role: fallback is 'member of my team'
    if ([string]::IsNullOrWhiteSpace($txtRole.Text) -and $msg -match [regex]::Escape('member of my team')) {
        $warnings += "Role is not set — message is using the generic fallback 'member of my team'. Fix: Configuration > Profile > Role."
    }

    # Backup Contact: fallback is 'our support team'
    if ([string]::IsNullOrWhiteSpace($txtBackupContact.Text) -and $msg -match [regex]::Escape('our support team')) {
        $warnings += "Backup Contact is not set — message is using the generic fallback 'our support team'. Fix: Configuration > Profile > Backup."
    }

    # Team Alias: fallback is 'Azure Networking Support'
    if ([string]::IsNullOrWhiteSpace($txtTeamAlias.Text) -and $msg -match [regex]::Escape('Azure Networking Support')) {
        $warnings += "Team Alias is not set — message is using the default fallback 'Azure Networking Support'. Fix: Configuration > Profile > Team Alias."
    }

    # Support Link: fallback is 'AzureBU@microsoft.com'
    if ([string]::IsNullOrWhiteSpace($txtSupportLink.Text) -and $msg -match [regex]::Escape('AzureBU@microsoft.com')) {
        $warnings += "Support Link is not set — message is using the default fallback 'AzureBU@microsoft.com'. Fix: Configuration > Profile > Support Link."
    }

    # Account email: genuinely blocking — [EMAIL] and signature links will be blank
    if ([string]::IsNullOrWhiteSpace($script:UserAlias)) {
        $warnings += "Account email is not set — [EMAIL] and signature link will be blank. Fix: Quick Actions > Account."
    }

    return $warnings
}

# Show-TemplateWarningDialog: Present placeholder/config warnings before applying a message.
# Returns 'Yes' if the user chooses to proceed, 'No' otherwise.
function Show-TemplateWarningDialog {
    param([string[]]$Warnings)
    $numbered = for ($index = 0; $index -lt $Warnings.Count; $index++) {
        "  $($index + 1). $($Warnings[$index])"
    }
    $body = "Review the following before applying:`n`n$($numbered -join "`n")`n`nApply the message anyway?"
    return [System.Windows.MessageBox]::Show($body, "Template Warnings ($($Warnings.Count))", 'YesNo', 'Warning')
}

# Export-MessageToFile: Write an HTML message body to disk.
function Export-MessageToFile($FilePath, $Content) {
    $Content | Out-File -FilePath $FilePath -Encoding utf8
}

# Get-LastWorkDayEndOfShift: Find the end-of-shift time on the most recent work day.
# If today is a work day and we haven't passed end-of-shift yet, use today.
# Otherwise, walk backwards to find the last work day.
# This ensures vacation OOF starts from when the user actually left the office.
function Get-LastWorkDayEndOfShift {
    if ($null -eq $script:EndOfShift -or $null -eq $script:WorkDays) { return $script:EndOfShift }

    $Now = Get-Date
    $TodayEndOfShift = (Get-Date).Date.Add($script:EndOfShift.TimeOfDay)

    # If today is a work day and we haven't passed end-of-shift, use today
    if ($Now.DayOfWeek -in $script:WorkDays -and $Now -le $TodayEndOfShift) {
        return $TodayEndOfShift
    }

    # If today is a work day but we've already passed end-of-shift, use today's end-of-shift
    if ($Now.DayOfWeek -in $script:WorkDays) {
        return $TodayEndOfShift
    }

    # Today is not a work day — walk backwards to find the last work day
    $CheckDate = $Now.AddDays(-1)
    for ($i = 0; $i -lt 7; $i++) {
        if ($CheckDate.DayOfWeek -in $script:WorkDays) {
            return $CheckDate.Date.Add($script:EndOfShift.TimeOfDay)
        }
        $CheckDate = $CheckDate.AddDays(-1)
    }

    # Fallback (shouldn't happen if WorkDays has at least one day)
    return $TodayEndOfShift
}

# Set-VacationAutoReply: Configure an extended/vacation OOF that runs from the
# last work day's end-of-shift until the given return date at shift-start time.
# This ensures the OOF starts from when the user was last in the office, not
# from the current day which may be a non-work day.
function Set-VacationAutoReply($ReturnDate) {
    if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift) { return }
    $ParsedDate = [datetime]$ReturnDate
    $EndTime = $ParsedDate + $script:StartOfShift.TimeOfDay
    $VacationStartTime = Get-LastWorkDayEndOfShift
    Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState "Scheduled" -StartTime $VacationStartTime -EndTime $EndTime
    Save-AutoReplyConfigToFile
}

# Disable-VacationAutoReply: Turn off the vacation/extended OOF by setting auto-reply to Disabled.
function Disable-VacationAutoReply {
    Set-AutoReplyState 'Disabled'
}

# Test-IsCurrentSessionElevated: Return $true when the current PowerShell session is running as Administrator.
function Test-IsCurrentSessionElevated {
    return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Get-PreferredTaskScriptPath: Pick the best available script path for task registration and repair.
function Get-PreferredTaskScriptPath {
    $candidateScriptPaths = @(
        (Join-Path $env:USERPROFILE "AAOOF-GUI.ps1"),
        $PSCommandPath,
        (Join-Path $ScriptDir "AAOOF-GUI.ps1")
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    return $candidateScriptPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
}

# Test-SamePath: Compare two file-system paths without requiring exact string formatting.
function Test-SamePath($PathA, $PathB) {
    if ([string]::IsNullOrWhiteSpace($PathA) -or [string]::IsNullOrWhiteSpace($PathB)) {
        return $false
    }

    try {
        $normalizedPathA = [System.IO.Path]::GetFullPath($PathA).TrimEnd('\\')
        $normalizedPathB = [System.IO.Path]::GetFullPath($PathB).TrimEnd('\\')
        return [string]::Equals($normalizedPathA, $normalizedPathB, [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch {
        return [string]::Equals($PathA, $PathB, [System.StringComparison]::OrdinalIgnoreCase)
    }
}

# Get-PowerShellExecutablePath: Resolve the current PowerShell host executable path.
function Get-PowerShellExecutablePath {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $psExe = Join-Path $PSHOME "pwsh.exe"
    }
    else {
        $psExe = Join-Path $PSHOME "powershell.exe"
    }

    if (-not (Test-Path $psExe)) {
        throw "PowerShell executable not found at '$psExe'. Cannot register scheduled task."
    }
    return $psExe
}

# Register-DailyScheduledTask: Create or update the 'AAOOF' scheduled task.
# The task runs this script daily in CLI mode with parameter '1'.
# Admin rights are required to create/update registration, but not required for task execution.
function Register-DailyScheduledTask {
    # Check elevation before attempting to register or update the task.
    $isAdmin = Test-IsCurrentSessionElevated
    if (-not $isAdmin) {
        throw "This action requires Administrator privileges.`n`nPlease close the app and re-run the script as Administrator, then try again."
    }

    if ($script:StartOfShift -ge $script:EndOfShift) {
        throw "Shift start must be earlier than shift end before creating the scheduled task."
    }
    if (-not $script:WorkDays -or $script:WorkDays.Count -eq 0) {
        throw "Select at least one work day before creating the scheduled task."
    }

    $scriptPath = Get-PreferredTaskScriptPath
    if (-not (Test-Path $scriptPath)) {
        $candidateScriptPaths = @(
            (Join-Path $env:USERPROFILE "AAOOF-GUI.ps1"),
            $PSCommandPath,
            (Join-Path $ScriptDir "AAOOF-GUI.ps1")
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        throw "Script not found at any expected location. Checked:`n - $($candidateScriptPaths -join "`n - ")"
    }

    $scriptWorkingDir = Split-Path -Parent $scriptPath

    # Resolve the PowerShell executable reliably across PS5, PS7, ISE, VS Code, etc.
    $psExe = Get-PowerShellExecutablePath

    $taskname = "AAOOF"
    $action = New-ScheduledTaskAction -Execute $psExe -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" 1" -WorkingDirectory $scriptWorkingDir
    $date = Get-Date -Date (Get-Date).Date
    $TriggerTime = $script:StartOfShift.TimeOfDay
    $TriggerTime = $date.AddMinutes($script:TaskStartOffsetMinutes) + $TriggerTime
    $trigger = New-ScheduledTaskTrigger -Daily -At $TriggerTime

    $settings = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -RestartCount 1 `
        -RestartInterval (New-TimeSpan -Minutes 1) `
        -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
        -MultipleInstances Queue

    $existing = Get-ScheduledTask -TaskName $taskname -ErrorAction SilentlyContinue
    if ($existing) {
        # Update the existing task so stale configurations are corrected
        Set-ScheduledTask -TaskName $taskname -Trigger $trigger -Action $action -Settings $settings -ErrorAction Stop | Out-Null
    }
    else {
        Register-ScheduledTask -TaskName $taskname -Trigger $trigger -Action $action -Settings $settings -RunLevel Highest -ErrorAction Stop | Out-Null
    }
    return $scriptPath
}

# Repair-DailyScheduledTaskScriptPath: Update the task action so it points to the preferred live script path.
function Repair-DailyScheduledTaskScriptPath {
    if (-not (Test-IsCurrentSessionElevated)) {
        throw "This action requires Administrator privileges.`n`nPlease close the app and re-run the script as Administrator, then try again."
    }

    $task = Get-ScheduledTask -TaskName "AAOOF" -ErrorAction SilentlyContinue
    if ($null -eq $task) {
        throw "Scheduled task 'AAOOF' has not been created yet."
    }

    $preferredScriptPath = Get-PreferredTaskScriptPath
    if (-not (Test-Path $preferredScriptPath)) {
        throw "Could not locate a valid script path to repair the task action."
    }

    $psExe = Get-PowerShellExecutablePath
    $scriptWorkingDir = Split-Path -Parent $preferredScriptPath
    $action = New-ScheduledTaskAction -Execute $psExe -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$preferredScriptPath`" 1" -WorkingDirectory $scriptWorkingDir
    Set-ScheduledTask -TaskName "AAOOF" -Action $action -ErrorAction Stop | Out-Null
    return $preferredScriptPath
}

# Get-ScheduledTaskResultText: Convert common task result codes into readable status text.
function Get-ScheduledTaskResultText($ResultCode) {
    switch ([int]$ResultCode) {
        0 { return "Success (0x0)" }
        1 { return "Incorrect function (0x1)" }
        267008 { return "Task is ready to run (0x41300)" }
        267009 { return "Task is currently running (0x41301)" }
        267010 { return "Task is disabled (0x41302)" }
        2147942402 { return "File not found (0x80070002)" }
        default { return ("0x{0:X}" -f [int]$ResultCode) }
    }
}

# Get-ScheduledTaskScriptPath: Extract the configured script path from the scheduled task action.
function Get-ScheduledTaskScriptPath($Task) {
    if ($null -eq $Task) { return $null }
    $action = $Task.Actions | Select-Object -First 1
    if ($null -eq $action -or [string]::IsNullOrWhiteSpace($action.Arguments)) { return $null }

    $match = [regex]::Match($action.Arguments, '-File\s+"([^"]+)"')
    if ($match.Success) {
        return $match.Groups[1].Value
    }
    return $null
}

# Get-DailyScheduledTaskStatus: Return a UI-friendly status object for the AAOOF task.
function Get-DailyScheduledTaskStatus {
    $task = Get-ScheduledTask -TaskName "AAOOF" -ErrorAction SilentlyContinue
    $preferredScriptPath = Get-PreferredTaskScriptPath
    if ($null -eq $task) {
        return [PSCustomObject]@{
            Exists = $false
            State = "Not created"
            NextRunTime = "-"
            LastRunTime = "-"
            LastResult = "-"
            ScriptPath = "-"
            Summary = "Task not created yet. Click Create Scheduled Task to register daily automation."
            SummaryBrush = [System.Windows.Media.Brushes]::DarkOrange
            CreateButtonLabel = "Create Scheduled Task"
            CanRunNow = $false
            CanEnable = $false
            CanDisable = $false
            CanRepairPath = $false
        }
    }

    $taskInfo = Get-ScheduledTaskInfo -TaskName "AAOOF" -ErrorAction SilentlyContinue
    $state = [string]$task.State
    $lastResult = if ($taskInfo) { Get-ScheduledTaskResultText $taskInfo.LastTaskResult } else { "-" }
    $taskScriptPath = Get-ScheduledTaskScriptPath $task
    $isPathMismatch = (-not [string]::IsNullOrWhiteSpace($preferredScriptPath)) -and (-not (Test-SamePath $taskScriptPath $preferredScriptPath))
    $summary = "Task is ready for daily automation."
    $summaryBrush = [System.Windows.Media.Brushes]::DarkGreen

    if ($state -eq "Running") {
        $summary = "Task is currently running. Wait for completion before starting it again."
        $summaryBrush = [System.Windows.Media.Brushes]::DarkOrange
    }
    elseif ($state -eq "Disabled") {
        $summary = "Task is disabled. Enable it to resume automatic runs."
        $summaryBrush = [System.Windows.Media.Brushes]::DarkOrange
    }
    elseif ($taskInfo -and ($taskInfo.LastTaskResult -ne 0) -and ($taskInfo.LastTaskResult -ne 267008) -and ($taskInfo.LastTaskResult -ne 267009)) {
        $summary = "Last result was $lastResult. Review Task Scheduler if runs are failing."
        $summaryBrush = [System.Windows.Media.Brushes]::DarkOrange
    }
    if ($isPathMismatch) {
        $summary = "Task points to a different script path. Click Repair Task Path to align with the preferred live script."
        $summaryBrush = [System.Windows.Media.Brushes]::DarkOrange
    }

    return [PSCustomObject]@{
        Exists = $true
        State = $state
        NextRunTime = if ($taskInfo -and $taskInfo.NextRunTime -and $taskInfo.NextRunTime.Year -gt 1900) { $taskInfo.NextRunTime.ToString("g") } else { "-" }
        LastRunTime = if ($taskInfo -and $taskInfo.LastRunTime -and $taskInfo.LastRunTime.Year -gt 1900) { $taskInfo.LastRunTime.ToString("g") } else { "-" }
        LastResult = $lastResult
        ScriptPath = $taskScriptPath
        Summary = $summary
        SummaryBrush = $summaryBrush
        CreateButtonLabel = "Update Scheduled Task"
        CanRunNow = ($state -ne "Running" -and $state -ne "Disabled")
        CanEnable = ($state -eq "Disabled")
        CanDisable = ($state -ne "Disabled")
        CanRepairPath = $isPathMismatch
    }
}

# Update-ScheduledTaskStatusUI: Refresh the scheduled task section in the Automation tab.
function Update-ScheduledTaskStatusUI {
    $taskStatus = Get-DailyScheduledTaskStatus
    if ($taskStatus.Exists) {
        $txtTaskExists.Text = "Created"
        $txtTaskExists.Foreground = [System.Windows.Media.Brushes]::Green
    }
    else {
        $txtTaskExists.Text = "Not created"
        $txtTaskExists.Foreground = [System.Windows.Media.Brushes]::DarkOrange
    }

    $txtTaskState.Text = $taskStatus.State
    $txtTaskNextRun.Text = $taskStatus.NextRunTime
    $txtTaskLastRun.Text = $taskStatus.LastRunTime
    $txtTaskLastResult.Text = $taskStatus.LastResult
    $txtTaskScriptPath.Text = if ([string]::IsNullOrWhiteSpace($taskStatus.ScriptPath)) { "-" } else { $taskStatus.ScriptPath }
    $txtTaskSummary.Text = $taskStatus.Summary
    $txtTaskSummary.Foreground = $taskStatus.SummaryBrush
    $btnCreateTask.Content = $taskStatus.CreateButtonLabel
    $btnRunTaskNow.IsEnabled = [bool]$taskStatus.CanRunNow
    $btnEnableTask.IsEnabled = [bool]$taskStatus.CanEnable
    $btnDisableTask.IsEnabled = [bool]$taskStatus.CanDisable
    $btnRepairTaskPath.IsEnabled = [bool]$taskStatus.CanRepairPath
}

# Export-AppConfiguration: Persist all global settings to config.json.
function Export-AppConfiguration {
    $startOfShiftStr = $null
    $endOfShiftStr = $null
    if ($null -ne $script:StartOfShift) { $startOfShiftStr = $script:StartOfShift.ToString("o") }
    if ($null -ne $script:EndOfShift) { $endOfShiftStr = $script:EndOfShift.ToString("o") }
    
    $cfg = @{
        StartOfShift    = $startOfShiftStr
        EndOfShift      = $endOfShiftStr
        WorkDays        = $script:WorkDays
        UserAlias       = $script:UserAlias
        UserAliasSuffix = $script:UserAliasSuffix
        FullName        = $script:FullName
        Role            = $script:Role
        BackupContact   = $script:BackupContact
        BackupEngineerEmail = $script:BackupEngineerEmail
        BackupEmail     = $script:BackupEmail
        TeamAlias       = $script:TeamAlias
        SupportLink     = $script:SupportLink
        OverrideAccount = $script:OverrideAccount
        TaskStartOffsetMinutes = $script:TaskStartOffsetMinutes
    }
    $cfg | ConvertTo-Json -Depth 5 | Set-Content $ConfigFile -Encoding utf8
}

# ===================== CLI Mode (for scheduled task / automation) =====================
# When invoked with a parameter, skip the GUI and run headless.
#   '1'   — Daily scheduled OOF update. Checks for active vacation before overwriting.
#   <date> — Set vacation/extended OOF until that return date.
if ($InputParameter) {
    Write-MissingHeadlessFiles

    if ($InputParameter -eq '1') {
        if (-not (Confirm-HeadlessConfigAvailable)) { exit }
        if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift -or $null -eq $script:WorkDays) {
            Write-Host "Configuration incomplete. Please run the GUI, configure Start/End shift and Work Days, then save." -ForegroundColor Red
            exit
        }
        Connect-ExchangeOnlineSession | Out-Null

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
    }
    if ($InputParameter -as [datetime]) {
        if (-not (Confirm-HeadlessConfigAvailable)) { exit }
        if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift) {
            Write-Host "Configuration incomplete. Please run the GUI, configure Start/End shift, then save." -ForegroundColor Red
            exit
        }
        Connect-ExchangeOnlineSession | Out-Null
        Set-VacationAutoReply $InputParameter
        $arc = Get-AutoReplyConfiguration
        Write-Host "Auto Reply: $($arc.AutoReplyState) | Start: $($arc.StartTime) | End: $($arc.EndTime)"
        Disconnect-ExchangeOnlineSession
    }

    # Check for updates after applying OOF changes so reply updates are not blocked.
    try {
        $remoteVer = Get-RemoteScriptVersion
        if (Test-IsRemoteVersionNewer -RemoteVersion $remoteVer -LocalVersion $script:ScriptVersion) {
            Write-Host "Update available: v$($script:ScriptVersion) -> v$remoteVer" -ForegroundColor Cyan
            $updated = Invoke-ScriptSelfUpdateExternal $InputParameter
            if ($updated) {
                Write-Host "Script update launched in a separate process." -ForegroundColor Green
            }
        }
        else {
            Write-Host "Script is up to date (v$($script:ScriptVersion))." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Auto-update skipped: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    exit
}

# ===================== Load XAML GUI from File =====================
# Parse the external XAML layout file and build the WPF window.
if (!(Test-Path $XamlFile)) {
    Write-Host "FATAL: XAML file not found at $XamlFile" -ForegroundColor Red
    [System.Windows.MessageBox]::Show(
        "The UI layout file was not found and could not be downloaded:`n`n$XamlFile`n`nPlease check your internet connection and try again, or manually place the AAOOF-GUI.xaml file in the config folder.",
        "AAOOF - Missing UI File",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    ) | Out-Null
    exit 1
}

# XAML version check: extract the embedded XamlVersion token and compare to the
# running script version. If stale and auto-download is enabled, pull a fresh copy
# from GitHub so layout changes (button sizing, new controls) take effect immediately.
$xamlRaw = Get-Content $XamlFile -Raw
$xamlVersionMatch = [regex]::Match($xamlRaw, '<!--\s*XamlVersion:\s*([\d\.]+)\s*-->')
$xamlVersion = if ($xamlVersionMatch.Success) { $xamlVersionMatch.Groups[1].Value } else { 'unknown' }
if ($xamlVersion -ne $script:ScriptVersion -and $script:EnableTemplateAutoDownload) {
    Write-Host "XAML version mismatch (XAML: $xamlVersion / Script: $($script:ScriptVersion)) — refreshing layout from GitHub..." -ForegroundColor Yellow
    try {
        $xamlUrl = "$RepoBaseUrl/AAOOF-GUI.xaml"
        Invoke-WebRequest -Uri $xamlUrl -OutFile $XamlFile -UseBasicParsing -TimeoutSec 10 -Headers @{ 'Cache-Control' = 'no-cache' }
        $xamlRaw = Get-Content $XamlFile -Raw
        Write-Host "XAML updated successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Warning: Could not refresh XAML layout: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}
[xml]$XAML = $xamlRaw

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

# Ensure parity for adjacent status buttons even when an older XAML is present locally.
$quickStatusButtonWidth = 170
$quickStatusButtonHeight = 36
$quickStatusButtonMargin = [System.Windows.Thickness]::new(4)
$btnRefreshStatus.Width = $quickStatusButtonWidth
$btnViewCurrentMsg.Width = $quickStatusButtonWidth
$btnRefreshStatus.MinWidth = $quickStatusButtonWidth
$btnViewCurrentMsg.MinWidth = $quickStatusButtonWidth
$btnRefreshStatus.MaxWidth = $quickStatusButtonWidth
$btnViewCurrentMsg.MaxWidth = $quickStatusButtonWidth
$btnRefreshStatus.Height = $quickStatusButtonHeight
$btnViewCurrentMsg.Height = $quickStatusButtonHeight
$btnRefreshStatus.Margin = $quickStatusButtonMargin
$btnViewCurrentMsg.Margin = $quickStatusButtonMargin

# --- Configuration and Automation tab controls ---
$txtFullName = $Window.FindName("txtFullName")
$txtRole = $Window.FindName("txtRole")
$txtBackupContact = $Window.FindName("txtBackupContact")
$txtBackupEngineerEmail = $Window.FindName("txtBackupEngineerEmail")
$txtTeamAlias = $Window.FindName("txtTeamAlias")
$txtSupportLink = $Window.FindName("txtSupportLink")
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
$btnRepairTaskPath = $Window.FindName("btnRepairTaskPath")
$btnEnableTask = $Window.FindName("btnEnableTask")
$btnDisableTask = $Window.FindName("btnDisableTask")
$btnRefreshTaskStatus = $Window.FindName("btnRefreshTaskStatus")
$btnRunTaskNow = $Window.FindName("btnRunTaskNow")
$btnOpenTaskScheduler = $Window.FindName("btnOpenTaskScheduler")
$btnCheckForUpdates = $Window.FindName("btnCheckForUpdates")
$btnExportDiagnostics = $Window.FindName("btnExportDiagnostics")
$txtTaskOffsetMinutes = $Window.FindName("txtTaskOffsetMinutes")
$txtTaskExists = $Window.FindName("txtTaskExists")
$txtTaskState = $Window.FindName("txtTaskState")
$txtTaskNextRun = $Window.FindName("txtTaskNextRun")
$txtTaskLastRun = $Window.FindName("txtTaskLastRun")
$txtTaskLastResult = $Window.FindName("txtTaskLastResult")
$txtTaskScriptPath = $Window.FindName("txtTaskScriptPath")
$txtTaskSummary = $Window.FindName("txtTaskSummary")
$txtLocalVersion = $Window.FindName("txtLocalVersion")
$txtRemoteVersion = $Window.FindName("txtRemoteVersion")

# Keep the update/diagnostics action buttons visually identical even when older XAML is present.
$updateActionButtonWidth = 180
$updateActionButtonHeight = 36
$updateActionButtonMargin = [System.Windows.Thickness]::new(4)
if ($null -ne $btnCheckForUpdates) {
    $btnCheckForUpdates.Width = $updateActionButtonWidth
    $btnCheckForUpdates.MinWidth = $updateActionButtonWidth
    $btnCheckForUpdates.MaxWidth = $updateActionButtonWidth
    $btnCheckForUpdates.Height = $updateActionButtonHeight
    $btnCheckForUpdates.Margin = $updateActionButtonMargin
}
if ($null -ne $btnExportDiagnostics) {
    $btnExportDiagnostics.Width = $updateActionButtonWidth
    $btnExportDiagnostics.MinWidth = $updateActionButtonWidth
    $btnExportDiagnostics.MaxWidth = $updateActionButtonWidth
    $btnExportDiagnostics.Height = $updateActionButtonHeight
    $btnExportDiagnostics.Margin = $updateActionButtonMargin
}

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
$btnBackupMessage = $Window.FindName("btnBackupMessage")

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
    if ($script:IsConnectedToEXO -and -not $script:OOFReplyEnabled) {
        $borderStatusBar.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xD8, 0x3B, 0x01))
        $txtStatusBar.Text = [char]0x26A0 + " Out of Office reply is not enabled | $Message"
    }
    else {
        $borderStatusBar.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
    }
    $Window.Dispatcher.Invoke([action] {}, [Windows.Threading.DispatcherPriority]::Render)
}

# Show-InfoDialog: Display an informational popup.
function Show-InfoDialog($Title, $Message) {
    [System.Windows.MessageBox]::Show($Message, $Title, 'OK', 'Information')
}

# Show-ErrorDialog: Display an error popup.
function Show-ErrorDialog($Title, $Message) {
    [System.Windows.MessageBox]::Show($Message, $Title, 'OK', 'Error')
}

# Get-DiagnosticsReport: Collect a snapshot of current app state for troubleshooting.
# Returns a formatted string covering identity, connection, OOF, task, config, and versions.
function Get-DiagnosticsReport {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $lines = @()
    $lines += "===== DailyOOF Diagnostics ====="
    $lines += "Generated : $ts"
    $lines += ""

    # --- Identity ---
    $lines += "--- Identity ---"
    $lines += "Windows User  : $($env:USERDOMAIN)\$($env:USERNAME)"
    $lines += "EXO Alias     : $(if ($script:UserAlias) { $script:UserAlias } else { '(not set)' })"
    $lines += "Full Name     : $(if ($script:FullName) { $script:FullName } else { '(not set)' })"
    $lines += "Role          : $(if ($script:Role) { $script:Role } else { '(not set)' })"
    $lines += "Backup Contact: $(if ($script:BackupContact) { $script:BackupContact } else { '(not set)' })"
    $lines += "Team Alias    : $(if ($script:TeamAlias) { $script:TeamAlias } else { '(not set)' })"
    $lines += "Support Link  : $(if ($script:SupportLink) { $script:SupportLink } else { '(not set)' })"
    $lines += ""

    # --- EXO Connection ---
    $lines += "--- Exchange Online Connection ---"
    $lines += "Connected     : $($script:IsConnectedToEXO)"
    try {
        $sessionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like 'ExchangeOnline_*' } |
            Select-Object -First 1
        if ($sessionInfo) {
            $lines += "Session Name  : $($sessionInfo.Name)"
            $lines += "Session State : $($sessionInfo.State)"
            $lines += "Token Expires : $($sessionInfo.TokenExpiryTimeUTC)"
        }
        else {
            $lines += "Session Info  : No active EXO session found"
        }
    }
    catch {
        $lines += "Session Info  : Could not retrieve ($($_.Exception.Message))"
    }
    $lines += ""

    # --- OOF State ---
    $lines += "--- OOF State ---"
    $lines += "OOF Reply Enabled : $($script:OOFReplyEnabled)"
    try {
        $arcPath = Get-AutoReplyConfigPath
        if (Test-Path $arcPath) {
            $arc = Get-Content $arcPath -Raw | ConvertFrom-Json
            $lines += "AutoReplyState    : $($arc.AutoReplyState)"
            $lines += "Start Time        : $($arc.StartTime)"
            $lines += "End Time          : $($arc.EndTime)"
        }
        else {
            $lines += "Cached ARC        : Not found (connect to populate)"
        }
    }
    catch {
        $lines += "Cached ARC        : Could not read ($($_.Exception.Message))"
    }
    $lines += ""

    # --- Scheduled Task ---
    $lines += "--- Scheduled Task ---"
    try {
        $task = Get-ScheduledTask -TaskName "AAOOF" -ErrorAction SilentlyContinue
        if ($task) {
            $taskInfo = Get-ScheduledTaskInfo -TaskName "AAOOF" -ErrorAction SilentlyContinue
            $lines += "Task Exists   : Yes"
            $lines += "Task State    : $($task.State)"
            $lines += "Next Run      : $(if ($taskInfo) { $taskInfo.NextRunTime } else { '-' })"
            $lines += "Last Run      : $(if ($taskInfo) { $taskInfo.LastRunTime } else { '-' })"
            $lines += "Last Result   : $(if ($taskInfo) { Get-ScheduledTaskResultText $taskInfo.LastTaskResult } else { '-' })"
            $lines += "Script Path   : $(Get-ScheduledTaskScriptPath $task)"
        }
        else {
            $lines += "Task Exists   : No — task not yet created"
        }
    }
    catch {
        $lines += "Task Info     : Could not retrieve ($($_.Exception.Message))"
    }
    $lines += ""

    # --- Configuration Files ---
    $lines += "--- Configuration Files ---"
    $lines += "Config Dir    : $ConfigDir"
    $lines += "Config File   : $ConfigFile ($(if (Test-Path $ConfigFile) { 'exists' } else { 'MISSING' }))"
    $lines += "XAML File     : $XamlFile ($(if (Test-Path $XamlFile) { 'exists' } else { 'MISSING' }))"
    foreach ($f in $DefaultConfigFiles) {
        $fp = Join-Path $ConfigDir $f
        $lines += "  $f : $(if (Test-Path $fp) { 'exists' } else { 'MISSING' })"
    }
    $lines += ""

    # --- Versions ---
    $lines += "--- Versions ---"
    $lines += "Script Version : $($script:ScriptVersion)"
    try {
        $remote = Get-RemoteScriptVersion
        $lines += "GitHub Version : $remote"
        $lines += "Update State   : $(Get-UpdateVersionState $remote $script:ScriptVersion)"
    }
    catch {
        $lines += "GitHub Version : Could not retrieve"
        $lines += "Update State   : Unknown"
    }
    $lines += ""

    # --- Environment ---
    $lines += "--- Environment ---"
    $lines += "OS             : $([System.Environment]::OSVersion.VersionString)"
    $lines += "PowerShell     : $($PSVersionTable.PSVersion)"
    $lines += "Script Dir     : $ScriptDir"
    $lines += "Is First Run   : $($script:IsFirstRun)"

    return $lines -join "`r`n"
}

# Show-DiagnosticsDialog: Display the diagnostics report in a scrollable window with
# Copy to Clipboard and Save to File options.
function Show-DiagnosticsDialog {
    $report = Get-DiagnosticsReport

    $diagXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Diagnostics Report" Width="620" Height="540"
        WindowStartupLocation="CenterOwner" ResizeMode="CanResize">
    <DockPanel Margin="12">
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="btnCopy"  Content="Copy to Clipboard" Padding="10,5" Margin="0,0,6,0"/>
            <Button x:Name="btnSave"  Content="Save to File..."   Padding="10,5" Margin="0,0,6,0"/>
            <Button x:Name="btnClose" Content="Close"             Padding="10,5" IsDefault="True" IsCancel="True"/>
        </StackPanel>
        <Border BorderBrush="#CCC" BorderThickness="1" CornerRadius="4">
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                <TextBox x:Name="txtReport" IsReadOnly="True" FontFamily="Consolas" FontSize="12"
                         Background="#FAFAFA" BorderThickness="0" Padding="8"
                         TextWrapping="NoWrap" VerticalAlignment="Stretch" AcceptsReturn="True"/>
            </ScrollViewer>
        </Border>
    </DockPanel>
</Window>
'@
    $reader  = [System.Xml.XmlReader]::Create([System.IO.StringReader]$diagXaml)
    $diagWin = [System.Windows.Markup.XamlReader]::Load($reader)
    $diagWin.Owner = $Window

    $txtReport = $diagWin.FindName('txtReport')
    $btnCopy   = $diagWin.FindName('btnCopy')
    $btnSave   = $diagWin.FindName('btnSave')
    $btnClose  = $diagWin.FindName('btnClose')

    $txtReport.Text = $report

    $btnCopy.Add_Click({
        [System.Windows.Clipboard]::SetText($txtReport.Text)
        $btnCopy.Content = "Copied!"
    })

    $btnSave.Add_Click({
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Title      = "Save Diagnostics Report"
        $dlg.Filter     = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        $dlg.FileName   = "AAOOF-Diagnostics-$(Get-Date -Format yyyyMMdd-HHmmss).txt"
        if ($dlg.ShowDialog($diagWin) -eq $true) {
            $txtReport.Text | Set-Content -Path $dlg.FileName -Encoding UTF8
            $btnSave.Content = "Saved!"
        }
    })

    $btnClose.Add_Click({ $diagWin.Close() })

    $diagWin.ShowDialog() | Out-Null
}


# Update-ConnectionUiState: Keep status text and Connect button visuals in sync.
function Update-ConnectionUiState {
    param(
        [ValidateSet('Connected', 'Disconnected', 'Failed')]
        [string]$State = 'Disconnected'
    )

    if ($State -eq 'Connected') {
        $txtConnectionStatus.Text = 'Connected'
        $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Green
        $btnConnect.Content = 'Connected'
        $btnConnect.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x10, 0x7C, 0x10))
        $btnConnect.Foreground = [System.Windows.Media.Brushes]::White
        $script:IsConnectedToEXO = $true
        return
    }

    if ($State -eq 'Failed') {
        $txtConnectionStatus.Text = 'Connection Failed'
        $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Red
    }
    else {
        $txtConnectionStatus.Text = 'Disconnected'
        $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::DarkOrange
    }

    $btnConnect.Content = 'Connect'
    $btnConnect.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x00, 0x78, 0xD4))
    $btnConnect.Foreground = [System.Windows.Media.Brushes]::White
    $script:IsConnectedToEXO = $false
}

# Assert-ExchangeConnection: Check for an active Exchange Online session and auto-connect
# if none exists. Updates the connection status UI. Returns $true if connected, $false on failure.
function Assert-ExchangeConnection {
    $session = Get-ConnectionInformation -ErrorAction SilentlyContinue
    $connected = $null -ne ($session | Where-Object { $_.Name -like "ExchangeOnline_*" })
    if ($connected) { return $true }

    Update-StatusBar "Not connected — attempting to connect..."
    if (-not $chkOverrideAccount.IsChecked) {
        Resolve-UserAlias
        $txtAccount.Text = $script:UserAlias
    }
    else {
        $script:UserAlias = $txtAccount.Text
    }
    Connect-ExchangeOnlineSession
    Update-ConnectionUiState -State Connected
    return $true
}

# ===================== Populate Combos =====================
# Fill the hour, minute, and AM/PM dropdown lists for shift time selection.
1..12 | ForEach-Object { $cmbStartHour.Items.Add($_.ToString()) | Out-Null; $cmbEndHour.Items.Add($_.ToString()) | Out-Null }
@("00", "15", "30", "45") | ForEach-Object { $cmbStartMin.Items.Add($_) | Out-Null; $cmbEndMin.Items.Add($_) | Out-Null }
@("AM", "PM") | ForEach-Object { $cmbStartAmPm.Items.Add($_) | Out-Null; $cmbEndAmPm.Items.Add($_) | Out-Null }

# ===================== Load Saved Config into UI =====================
# Initialize-UIFromConfig: Populate controls from saved config and apply defaults where needed.
function Initialize-UIFromConfig {
    # Full Name
    if (![string]::IsNullOrEmpty($script:FullName)) {
        $txtFullName.Text = $script:FullName
    }
    # Role
    if (![string]::IsNullOrEmpty($script:Role)) {
        $txtRole.Text = $script:Role
    }
    if (![string]::IsNullOrEmpty($script:BackupContact)) {
        $txtBackupContact.Text = $script:BackupContact
    }
    if (![string]::IsNullOrEmpty($script:BackupEngineerEmail)) {
        $txtBackupEngineerEmail.Text = $script:BackupEngineerEmail
    }
    if (![string]::IsNullOrEmpty($script:TeamAlias)) {
        $txtTeamAlias.Text = $script:TeamAlias
    }
    if (![string]::IsNullOrEmpty($script:SupportLink)) {
        $txtSupportLink.Text = $script:SupportLink
    }
    # Account override checkbox
    $chkOverrideAccount.IsChecked = $script:OverrideAccount
    $txtAccount.IsEnabled = $script:OverrideAccount
    # Account
    if ($script:OverrideAccount) {
        $txtAccount.Text = $script:UserAlias
    }
    else {
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
        $nearestMin = @("00", "15", "30", "45") | Sort-Object { [Math]::Abs([int]$_ - $m) } | Select-Object -First 1
        $cmbStartMin.SelectedItem = $nearestMin
        $cmbStartAmPm.SelectedItem = $ampm
    }
    else {
        $cmbStartHour.SelectedIndex = 8; $cmbStartMin.SelectedIndex = 0; $cmbStartAmPm.SelectedIndex = 0 # 9 AM
    }

    if ($null -ne $script:EndOfShift) {
        $h = (Get-Date $script:EndOfShift).Hour
        $m = (Get-Date $script:EndOfShift).Minute
        $ampm = if ($h -ge 12) { "PM" } else { "AM" }
        $displayH = if ($h -gt 12) { $h - 12 } elseif ($h -eq 0) { 12 } else { $h }
        $cmbEndHour.SelectedItem = $displayH.ToString()
        $nearestMin = @("00", "15", "30", "45") | Sort-Object { [Math]::Abs([int]$_ - $m) } | Select-Object -First 1
        $cmbEndMin.SelectedItem = $nearestMin
        $cmbEndAmPm.SelectedItem = $ampm
    }
    else {
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
    }
    else {
        # Default Mon-Fri
        $chkMon.IsChecked = $true; $chkTue.IsChecked = $true; $chkWed.IsChecked = $true
        $chkThu.IsChecked = $true; $chkFri.IsChecked = $true
    }

    $txtTaskOffsetMinutes.Text = [string]$script:TaskStartOffsetMinutes
}

# Resolve-TemplateFilePath: Map a template display name to its file path in the config directory.
function Resolve-TemplateFilePath($TemplateName) {
    switch ($TemplateName) {
        "Normal OOF" { return Join-Path $ConfigDir "normal_oof.html" }
        "Vacation OOF" { return Join-Path $ConfigDir "vacation_oof.html" }
        "Sick OOF" { return Join-Path $ConfigDir "sick_oof.html" }
        "Holiday OOF" { return Join-Path $ConfigDir "holiday_oof.html" }
        "Placeholder Examples" { return Join-Path $ConfigDir "placeholder_examples.html" }
        default { return $null }
    }
}

# Get-LegacyMessageBackupPath: Legacy path used before message backups were promoted to
# first-class HTML templates.
function Get-LegacyMessageBackupPath {
    return Join-Path $ConfigDir "message.html.bak"
}

# Get-LastMessageTemplatePath: Dedicated HTML template path used to preserve the current
# editor content before loading another template.
function Get-LastMessageTemplatePath {
    $templatePath = Join-Path $ConfigDir "last_message_template.html"
    $legacyPath = Get-LegacyMessageBackupPath

    if ((-not (Test-Path $templatePath)) -and (Test-Path $legacyPath)) {
        try {
            Move-Item -Path $legacyPath -Destination $templatePath -Force
        }
        catch {
            try { Copy-Item -Path $legacyPath -Destination $templatePath -Force } catch { }
        }
    }

    return $templatePath
}

# Resolve-TemplatePlaceholders: Process an HTML template string, replacing:
#   [RETURN DATE]   — with the selected return date from the date picker
#   [HOLIDAY NAME]  — with the selected holiday name
#   [ROLE]          — with the user's configured role (or default)
#   [BACKUP CONTACT] — with the configured backup contact (or default)
#   [TEAM ALIAS]    — with the configured team alias (or default)
#   [SUPPORT LINK]  — with the configured support link (or default)
#   [FULL NAME]     — with the user's display name
#   [FIRST NAME]    — first name derived from display name
#   [LAST NAME]     — last name derived from display name
#   [EMAIL]         — with the user's email address
#   [OFFICE HOURS]  — with configured shift start–end times
#   [WORK DAYS]     — with configured work days
#   [TIMEZONE]      — with the local timezone
#   [SIGNATURE]     — with an auto-generated signature block (or removed if unchecked)
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

    # Replace 1.8.0 contact placeholders with configured values or safe defaults.
    $backupContact = if (![string]::IsNullOrWhiteSpace($txtBackupContact.Text)) { $txtBackupContact.Text } else { 'our support team' }
    $backupEngineerEmail = if (![string]::IsNullOrWhiteSpace($txtBackupEngineerEmail.Text)) { $txtBackupEngineerEmail.Text } else { '' }
    $teamAlias = if (![string]::IsNullOrWhiteSpace($txtTeamAlias.Text)) { $txtTeamAlias.Text } else { 'Azure Networking Support' }
    $supportLink = if (![string]::IsNullOrWhiteSpace($txtSupportLink.Text)) { $txtSupportLink.Text } else { 'AzureBU@microsoft.com' }
    $text = $text -replace '\[BACKUP CONTACT\]', $backupContact
    $text = $text -replace '\[BACKUP ENGINEER EMAIL\]', $backupEngineerEmail
    $text = $text -replace '\[TEAM ALIAS\]', $teamAlias
    $text = $text -replace '\[SUPPORT LINK\]', $supportLink

    # Derive display name for name-based placeholders
    if (![string]::IsNullOrWhiteSpace($txtFullName.Text)) {
        $displayName = $txtFullName.Text
    }
    else {
        $aliasLocal = ($script:UserAlias -split '@')[0]
        if ($aliasLocal) {
            if ($aliasLocal -match '\.' ) {
                $nameParts = $aliasLocal -split '\.'
            }
            else {
                $nameParts = [regex]::Split($aliasLocal, '(?<=[a-z])(?=[A-Z])')
            }
            $displayName = ($nameParts | ForEach-Object { (Get-Culture).TextInfo.ToTitleCase($_.ToLower()) }) -join ' '
        }
        else {
            $displayName = $env:USERNAME
        }
    }

    # Replace [FULL NAME], [FIRST NAME], [LAST NAME]
    $text = $text -replace '\[FULL NAME\]', $displayName
    $nameTokens = $displayName -split '\s+', 2
    $firstName = $nameTokens[0]
    $lastName = if ($nameTokens.Count -gt 1) { $nameTokens[1] } else { '' }
    $text = $text -replace '\[FIRST NAME\]', $firstName
    $text = $text -replace '\[LAST NAME\]', $lastName

    # Replace [EMAIL]
    if (![string]::IsNullOrWhiteSpace($script:UserAlias)) {
        $text = $text -replace '\[EMAIL\]', $script:UserAlias
    }

    # Replace [OFFICE HOURS]
    if ($null -ne $script:StartOfShift -and $null -ne $script:EndOfShift) {
        $hoursStr = "$($script:StartOfShift.ToString('h:mm tt')) - $($script:EndOfShift.ToString('h:mm tt'))"
        $text = $text -replace '\[OFFICE HOURS\]', $hoursStr
    }

    # Replace [WORK DAYS]
    if ($script:WorkDays -and $script:WorkDays.Count -gt 0) {
        $weekOrder = @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')
        $sorted = $script:WorkDays | Sort-Object { $weekOrder.IndexOf($_) }
        $text = $text -replace '\[WORK DAYS\]', ($sorted -join ', ')
    }

    # Replace [TIMEZONE]
    $text = $text -replace '\[TIMEZONE\]', [System.TimeZoneInfo]::Local.DisplayName

    # Auto-generate signature block: the greeting/name is conditional on the
    # "Include Signature" checkbox, but office details are always included per their own toggles.
    $sigLines = @()

    if ($chkIncludeSignature.IsChecked) {
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
        $weekOrder = @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')
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
    }
    else {
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

# Read-TaskSettingsFromUI: Parse scheduled task settings from the Automation tab.
function Read-TaskSettingsFromUI {
    $parsedOffset = 0
    if ([int]::TryParse($txtTaskOffsetMinutes.Text, [ref]$parsedOffset)) {
        $script:TaskStartOffsetMinutes = $parsedOffset
    }
}

# ===================== Event Handlers =====================
# Wire up button clicks, checkbox changes, and other UI events to their logic.

# Connect: Resolve the user alias and establish an Exchange Online session.
$btnConnect.Add_Click({
        try {
            Update-StatusBar "Connecting to Exchange Online..."
            if (-not $chkOverrideAccount.IsChecked) {
                Resolve-UserAlias
                $txtAccount.Text = $script:UserAlias
            }
            else {
                $script:UserAlias = $txtAccount.Text
            }
            $connectCtx = Show-ConnectingWindow
            try {
                Connect-ExchangeOnlineSession
            }
            finally {
                Close-ConnectingWindow $connectCtx
            }
            # If the user hit Cancel during auth, abandon the connect flow gracefully.
            if ($connectCtx.SyncHash.Cancelled) {
                Update-ConnectionUiState -State Disconnected
                try { Disconnect-ExchangeOnlineSession } catch {}
                Update-StatusBar "Connection cancelled"
                return
            }
            Update-ConnectionUiState -State Connected

            # On first connect, pull current OOF config and message and save locally
            try {
                $arc = Get-AutoReplyConfiguration
                $txtARCState.Text = $arc.AutoReplyState
                $txtARCStart.Text = $arc.StartTime.ToString()
                $txtARCEnd.Text = $arc.EndTime.ToString()

                # Check if OOF auto-reply is enabled (not Disabled)
                $script:OOFReplyEnabled = ($arc.AutoReplyState -ne 'Disabled')

                # Save the current online messages to template files if we don't already have a saved message
                $savedMsgFile = Join-Path $ConfigDir "message.html"
                if (!(Test-Path $savedMsgFile) -and ![string]::IsNullOrWhiteSpace($arc.ExternalMessage)) {
                    Export-MessageToFile $savedMsgFile $arc.ExternalMessage
                }
            }
            catch { }

            # Auto-populate profile fields (Full Name, Role) from EXO on first run or when blank.
            # If EXO lookup fails and name is still blank, prompt the user to enter it.
            try {
                $nameResolved = Resolve-ProfileFromEXO
                if (-not $nameResolved -and [string]::IsNullOrWhiteSpace($script:FullName)) {
                    Show-NameInputDialog
                }
            }
            catch { }

            Export-AppConfiguration
            Update-StatusBar "Connected as $($script:UserAlias)"
        }
        catch {
            Update-ConnectionUiState -State Failed
            Show-ErrorDialog "Connection Error" $_.Exception.Message
            Update-StatusBar "Connection failed"
        }
    })

# Disconnect: Tear down the Exchange Online session and update the UI.
$btnDisconnect.Add_Click({
        try {
            Disconnect-ExchangeOnlineSession
            Update-ConnectionUiState -State Disconnected
            $script:OOFReplyEnabled = $true
            Update-StatusBar "Disconnected from Exchange Online"
        }
        catch {
            Show-ErrorDialog "Disconnect Error" $_.Exception.Message
        }
    })

# Enable Scheduled Auto Reply: Read shift/work-day settings and apply Scheduled mode.
$btnEnableScheduled.Add_Click({
        try {
            Assert-ExchangeConnection
            Update-StatusBar "Setting scheduled auto reply..."
            Read-ShiftTimesFromUI
            $script:WorkDays = Read-WorkDaysFromUI
            Set-AutoReplyState 'Scheduled'
            $arc = Get-AutoReplyConfiguration
            $txtARCState.Text = $arc.AutoReplyState
            $txtARCStart.Text = $arc.StartTime.ToString()
            $txtARCEnd.Text = $arc.EndTime.ToString()
            $script:OOFReplyEnabled = ($arc.AutoReplyState -ne 'Disabled')
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
            Assert-ConfigurationValid -RequireShiftTimes -RequireFutureReturnDate -RequireOverrideEmail
            Assert-ExchangeConnection
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
            $script:OOFReplyEnabled = ($arc.AutoReplyState -ne 'Disabled')
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
            Assert-ExchangeConnection
            Update-StatusBar "Cancelling vacation OOF..."
            Disable-VacationAutoReply
            $arc = Get-AutoReplyConfiguration
            $txtARCState.Text = $arc.AutoReplyState
            $txtARCStart.Text = $arc.StartTime.ToString()
            $txtARCEnd.Text = $arc.EndTime.ToString()
            $script:OOFReplyEnabled = ($arc.AutoReplyState -ne 'Disabled')
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
            Assert-ExchangeConnection
            Update-StatusBar "Refreshing status..."
            $arc = Get-AutoReplyConfiguration
            $txtARCState.Text = $arc.AutoReplyState
            $txtARCStart.Text = $arc.StartTime.ToString()
            $txtARCEnd.Text = $arc.EndTime.ToString()
            $script:OOFReplyEnabled = ($arc.AutoReplyState -ne 'Disabled')
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
                    Resolve-UserAlias
                    $txtAccount.Text = $script:UserAlias
                }
                else {
                    $script:UserAlias = $txtAccount.Text
                }
                Connect-ExchangeOnlineSession
                Update-ConnectionUiState -State Connected
            }

            $arc = Get-AutoReplyConfiguration
            $msg = if (![string]::IsNullOrWhiteSpace($arc.ExternalMessage)) { $arc.ExternalMessage }
            elseif (![string]::IsNullOrWhiteSpace($arc.InternalMessage)) { $arc.InternalMessage }
            else { $null }
            if ($null -eq $msg) {
                $wbCurrentOOF.NavigateToString("<html><body style='font-family:Segoe UI;padding:20px;color:#888;'><h3>No OOF message is currently set.</h3></body></html>")
                $txtCurrentOOFStatus.Text = "No message set"
                Update-StatusBar "No current OOF message"
            }
            else {
                $wbCurrentOOF.NavigateToString($msg)
                $txtCurrentOOFStatus.Text = "State: $($arc.AutoReplyState) | Loaded $(Get-Date -Format 'h:mm tt')"
                Update-StatusBar "Current OOF message loaded"
            }

            $txtARCState.Text = $arc.AutoReplyState
            $txtARCStart.Text = $arc.StartTime.ToString()
            $txtARCEnd.Text = $arc.EndTime.ToString()

            $script:OOFReplyEnabled = ($arc.AutoReplyState -ne 'Disabled')

            # Switch to the Current OOF tab
            $tcMain.SelectedIndex = 4
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
                        Resolve-UserAlias
                        $txtAccount.Text = $script:UserAlias
                    }
                    else {
                        $script:UserAlias = $txtAccount.Text
                    }
                    Connect-ExchangeOnlineSession
                    Update-ConnectionUiState -State Connected
                }
                catch {
                    Update-ConnectionUiState -State Failed
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
            }
            else {
                $wbCurrentOOF.NavigateToString($msg)
                $txtCurrentOOFStatus.Text = "State: $($arc.AutoReplyState) | Loaded $(Get-Date -Format 'h:mm tt')"
                Update-StatusBar "Current OOF message loaded"
            }

            $script:OOFReplyEnabled = ($arc.AutoReplyState -ne 'Disabled')
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
        if ($tcMain.SelectedIndex -eq 3) {
            # Sync config globals from UI before template rendering
            Read-ShiftTimesFromUI
            $script:WorkDays = Read-WorkDaysFromUI
            $script:FullName = $txtFullName.Text
            $script:Role = $txtRole.Text
            & $optionReloadHandler
            return
        }
        if ($tcMain.SelectedIndex -ne 4) { return }
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
                    Resolve-UserAlias
                    $txtAccount.Text = $script:UserAlias
                }
                else {
                    $script:UserAlias = $txtAccount.Text
                }
                Connect-ExchangeOnlineSession
                Update-ConnectionUiState -State Connected
            }

            $arc = Get-AutoReplyConfiguration
            $msg = if (![string]::IsNullOrWhiteSpace($arc.ExternalMessage)) { $arc.ExternalMessage }
            elseif (![string]::IsNullOrWhiteSpace($arc.InternalMessage)) { $arc.InternalMessage }
            else { $null }
            if ($null -eq $msg) {
                $wbCurrentOOF.NavigateToString("<html><body style='font-family:Segoe UI;padding:20px;color:#888;'><h3>No OOF message is currently set.</h3></body></html>")
                $txtCurrentOOFStatus.Text = "No message set"
                Update-StatusBar "No current OOF message"
            }
            else {
                $wbCurrentOOF.NavigateToString($msg)
                $txtCurrentOOFStatus.Text = "State: $($arc.AutoReplyState) | Loaded $(Get-Date -Format 'h:mm tt')"
                Update-StatusBar "Current OOF message loaded"
            }

            $txtARCState.Text = $arc.AutoReplyState
            $txtARCStart.Text = $arc.StartTime.ToString()
            $txtARCEnd.Text = $arc.EndTime.ToString()

            $script:OOFReplyEnabled = ($arc.AutoReplyState -ne 'Disabled')
        }
        catch {
            $script:CurrentOOFLoaded = $false
            $wbCurrentOOF.NavigateToString("<html><body style='font-family:Segoe UI;padding:20px;color:red;'><h3>Error</h3><p>Could not load message.</p><p style='color:#888;font-size:10pt;'>$([System.Web.HttpUtility]::HtmlEncode($_.Exception.Message))</p></body></html>")
            $txtCurrentOOFStatus.Text = "Error loading"
            Update-StatusBar "Failed to load current OOF message"
        }
    })

# Debounced auto-save: When any config control changes, restart a 0.5-second timer.
# When the timer fires (user stopped making changes), read all UI fields and save.
$script:ConfigSaveTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:ConfigSaveTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:ConfigSaveTimer.Add_Tick({
    $timer = $this
    if ($null -ne $timer) {
        try { $timer.Stop() } catch { }
    }
    $script:FullName = $txtFullName.Text
    $script:Role = $txtRole.Text
    $script:BackupContact = $txtBackupContact.Text
    $script:BackupEngineerEmail = $txtBackupEngineerEmail.Text
    $script:TeamAlias = $txtTeamAlias.Text
    $script:SupportLink = $txtSupportLink.Text
    if ($chkOverrideAccount.IsChecked) {
        $script:UserAlias = $txtAccount.Text
    }
    else {
        Resolve-UserAlias
    }
    Read-ShiftTimesFromUI
    Read-TaskSettingsFromUI
    $script:WorkDays = Read-WorkDaysFromUI
    Export-AppConfiguration
    Update-StatusBar "💾 Settings saved"
})

function Request-DebouncedConfigSave {
    Read-TaskSettingsFromUI
    $script:ConfigSaveTimer.Stop()
    $script:ConfigSaveTimer.Start()
}

# Wire config controls to trigger debounced auto-save

$cmbStartHour.Add_SelectionChanged({ Request-DebouncedConfigSave })
$cmbStartMin.Add_SelectionChanged({ Request-DebouncedConfigSave })
$cmbStartAmPm.Add_SelectionChanged({ Request-DebouncedConfigSave })
$cmbEndHour.Add_SelectionChanged({ Request-DebouncedConfigSave })
$cmbEndMin.Add_SelectionChanged({ Request-DebouncedConfigSave })
$cmbEndAmPm.Add_SelectionChanged({ Request-DebouncedConfigSave })
$txtTaskOffsetMinutes.Add_TextChanged({ Request-DebouncedConfigSave })
$chkSun.Add_Checked({ Request-DebouncedConfigSave }); $chkSun.Add_Unchecked({ Request-DebouncedConfigSave })
$chkMon.Add_Checked({ Request-DebouncedConfigSave }); $chkMon.Add_Unchecked({ Request-DebouncedConfigSave })
$chkTue.Add_Checked({ Request-DebouncedConfigSave }); $chkTue.Add_Unchecked({ Request-DebouncedConfigSave })
$chkWed.Add_Checked({ Request-DebouncedConfigSave }); $chkWed.Add_Unchecked({ Request-DebouncedConfigSave })
$chkThu.Add_Checked({ Request-DebouncedConfigSave }); $chkThu.Add_Unchecked({ Request-DebouncedConfigSave })
$chkFri.Add_Checked({ Request-DebouncedConfigSave }); $chkFri.Add_Unchecked({ Request-DebouncedConfigSave })
$chkSat.Add_Checked({ Request-DebouncedConfigSave }); $chkSat.Add_Unchecked({ Request-DebouncedConfigSave })

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
            Assert-ExchangeConnection
            Update-StatusBar "Setting auto reply to Enabled..."
            Set-AutoReplyState 'Enabled'
            $script:OOFReplyEnabled = $true
            Update-StatusBar "Auto reply set to Enabled"
            Show-InfoDialog "Done" "Auto Reply State set to Enabled"
        }
        catch { Show-ErrorDialog "Error" $_.Exception.Message }
    })

$btnStateDisabled.Add_Click({
        try {
            Assert-ExchangeConnection
            Update-StatusBar "Setting auto reply to Disabled..."
            Set-AutoReplyState 'Disabled'
            $script:OOFReplyEnabled = $false
            Update-StatusBar "Auto reply set to Disabled"
            Show-InfoDialog "Done" "Auto Reply State set to Disabled"
        }
        catch { Show-ErrorDialog "Error" $_.Exception.Message }
    })

$btnStateScheduled.Add_Click({
        try {
            Assert-ConfigurationValid -RequireShiftTimes -RequireWorkDays -RequireOverrideEmail
            Assert-ExchangeConnection
            Update-StatusBar "Setting auto reply to Scheduled..."
            Read-ShiftTimesFromUI
            $script:WorkDays = Read-WorkDaysFromUI
            Set-AutoReplyState 'Scheduled'
            $script:OOFReplyEnabled = $true
            Update-StatusBar "Auto reply set to Scheduled"
            Show-InfoDialog "Done" "Auto Reply State set to Scheduled"
        }
        catch { Show-ErrorDialog "Error" $_.Exception.Message }
    })

# Create Scheduled Task: Register a Windows Task Scheduler job to run this script daily.
$btnCreateTask.Add_Click({
        try {
            $existingTask = Get-ScheduledTask -TaskName "AAOOF" -ErrorAction SilentlyContinue
            Assert-ConfigurationValid -RequireShiftTimes -RequireWorkDays -RequireOverrideEmail -RequireTaskOffset
            $taskScriptPath = Register-DailyScheduledTask
            if ($taskScriptPath) {
                Update-ScheduledTaskStatusUI
                $taskAction = if ($null -eq $existingTask) { "created" } else { "updated" }
                Show-InfoDialog "Success" "Scheduled task 'AAOOF' $taskAction successfully.`n`nScript path:`n$taskScriptPath"
                Update-StatusBar "Scheduled task ready"
            }
        }
        catch {
            Show-ErrorDialog "Error" "Failed to create task.`n`n$($_.Exception.Message)"
            Update-StatusBar "Task creation failed"
        }
    })

# Refresh Task Status: Query the current AAOOF task registration and update the UI.
$btnRefreshTaskStatus.Add_Click({
        try {
            Update-ScheduledTaskStatusUI
            Update-StatusBar "Scheduled task status refreshed"
        }
        catch {
            Show-ErrorDialog "Task Error" "Could not refresh scheduled task status.`n`n$($_.Exception.Message)"
            Update-StatusBar "Task status refresh failed"
        }
    })

# Run Task Now: Start the registered AAOOF scheduled task immediately.
$btnRunTaskNow.Add_Click({
        try {
            $task = Get-ScheduledTask -TaskName "AAOOF" -ErrorAction SilentlyContinue
            if ($null -eq $task) {
                throw "Scheduled task 'AAOOF' has not been created yet."
            }
            if ([string]$task.State -eq "Disabled") {
                throw "Scheduled task 'AAOOF' is disabled. Enable it before running it manually."
            }
            if ([string]$task.State -eq "Running") {
                throw "Scheduled task 'AAOOF' is already running."
            }
            Start-ScheduledTask -TaskName "AAOOF" -ErrorAction Stop
            Update-ScheduledTaskStatusUI
            Update-StatusBar "Scheduled task started"
            Show-InfoDialog "Task Started" "Scheduled task 'AAOOF' was started successfully."
        }
        catch {
            Show-ErrorDialog "Task Error" "Could not start scheduled task.`n`n$($_.Exception.Message)"
            Update-StatusBar "Task start failed"
        }
    })

# Repair Task Path: Repoint task action to the preferred live script path.
$btnRepairTaskPath.Add_Click({
        try {
            $repairedPath = Repair-DailyScheduledTaskScriptPath
            Update-ScheduledTaskStatusUI
            Update-StatusBar "Scheduled task path repaired"
            Show-InfoDialog "Task Path Repaired" "Scheduled task 'AAOOF' now points to:`n$repairedPath"
        }
        catch {
            Show-ErrorDialog "Task Error" "Could not repair scheduled task path.`n`n$($_.Exception.Message)"
            Update-StatusBar "Task path repair failed"
        }
    })

# Enable Task: Re-enable the AAOOF scheduled task without recreating it.
$btnEnableTask.Add_Click({
        try {
            $task = Get-ScheduledTask -TaskName "AAOOF" -ErrorAction SilentlyContinue
            if ($null -eq $task) {
                throw "Scheduled task 'AAOOF' has not been created yet."
            }
            Enable-ScheduledTask -TaskName "AAOOF" -ErrorAction Stop | Out-Null
            Update-ScheduledTaskStatusUI
            Update-StatusBar "Scheduled task enabled"
        }
        catch {
            Show-ErrorDialog "Task Error" "Could not enable scheduled task.`n`n$($_.Exception.Message)"
            Update-StatusBar "Task enable failed"
        }
    })

# Disable Task: Pause the AAOOF scheduled task without deleting it.
$btnDisableTask.Add_Click({
        try {
            $task = Get-ScheduledTask -TaskName "AAOOF" -ErrorAction SilentlyContinue
            if ($null -eq $task) {
                throw "Scheduled task 'AAOOF' has not been created yet."
            }
            Disable-ScheduledTask -TaskName "AAOOF" -ErrorAction Stop | Out-Null
            Update-ScheduledTaskStatusUI
            Update-StatusBar "Scheduled task disabled"
        }
        catch {
            Show-ErrorDialog "Task Error" "Could not disable scheduled task.`n`n$($_.Exception.Message)"
            Update-StatusBar "Task disable failed"
        }
    })

# Open Task Scheduler: Launch the Windows Task Scheduler MMC for task review and editing.
$btnOpenTaskScheduler.Add_Click({
        try {
            Start-Process "taskschd.msc"
            Update-StatusBar "Task Scheduler opened"
        }
        catch {
            Show-ErrorDialog "Task Scheduler Error" "Could not open Task Scheduler.`n`n$($_.Exception.Message)"
            Update-StatusBar "Task Scheduler open failed"
        }
    })

# Check for Updates: If a newer version exists, launch the external updater and prompt for restart.
$btnCheckForUpdates.Add_Click({
        try {
            Update-StatusBar "Checking for updates..."
            # Show remote version
            $remoteVer = Get-RemoteScriptVersion
            $txtRemoteVersion.Text = $remoteVer
            $updateState = Get-UpdateVersionState -RemoteVersion $remoteVer -LocalVersion $script:ScriptVersion
            switch ($updateState) {
                'Unknown' {
                    $txtRemoteVersion.Foreground = [System.Windows.Media.Brushes]::DarkOrange
                    Update-StatusBar "Could not determine GitHub version"
                    Show-InfoDialog "No Update" "Could not determine a newer version from GitHub right now."
                    return
                }
                'LocalNewer' {
                    $txtRemoteVersion.Foreground = [System.Windows.Media.Brushes]::Green
                    Update-StatusBar "Local version is newer than GitHub"
                    Show-InfoDialog "No Update" "Your local version (v$($script:ScriptVersion)) is newer than GitHub (v$remoteVer)."
                    return
                }
                'UpToDate' {
                    $txtRemoteVersion.Foreground = [System.Windows.Media.Brushes]::Green
                    Update-StatusBar "Already up to date"
                    Show-InfoDialog "No Update" "You are already running the latest version."
                    return
                }
                default {
                    $txtRemoteVersion.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xD8, 0x3B, 0x01))
                    Update-StatusBar "Update available: v$remoteVer"
                }
            }

            $updated = Invoke-ScriptSelfUpdateExternal
            if ($updated) {
                Update-StatusBar "Update downloaded successfully. Please restart the script now."
                Show-InfoDialog "Update Complete" "The application has been updated successfully. Please restart the script now."
                $Window.Close()
            }
        }
        catch {
            Show-ErrorDialog "Update Error" $_.Exception.Message
            Update-StatusBar "Update check failed"
        }
    })

# Export Diagnostics: Build and display a full app state snapshot.
if ($null -ne $btnExportDiagnostics) {
    $btnExportDiagnostics.Add_Click({
        try {
            Show-DiagnosticsDialog
        }
        catch {
            Show-ErrorDialog "Diagnostics Error" $_.Exception.Message
        }
    })
}

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

        # Check for the last loaded-message template snapshot
        $lastMessageTemplate = Get-LastMessageTemplatePath
        if (Test-Path $lastMessageTemplate) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = "Last Loaded Message"
            $item.Tag = $lastMessageTemplate
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
            # Save current message as a reusable HTML template before overwriting.
            if (![string]::IsNullOrWhiteSpace($txtMessage.Text)) {
                $backupFile = Get-LastMessageTemplatePath
                Export-MessageToFile $backupFile $txtMessage.Text
            }
            $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $path -Raw)
            # Refresh preview if on Preview tab
            if ($tcMessageView.SelectedIndex -eq 1) {
                $wbPreview.NavigateToString($txtMessage.Text)
            }
            Update-StatusBar "Template loaded: $selected"
        }
        else {
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
                $backupFile = Get-LastMessageTemplatePath
                Export-MessageToFile $backupFile $txtMessage.Text
            }
            $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $path -Raw)
            if ($tcMessageView.SelectedIndex -eq 1) {
                $wbPreview.NavigateToString($txtMessage.Text)
            }
            Update-StatusBar "Template loaded: $selected"
        }
        else {
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
        $script:IsLoadingTemplate = $true
        $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $path -Raw)
        $script:IsLoadingTemplate = $false
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
        Request-DebouncedConfigSave
        & $optionReloadHandler
    })
$txtRole.Add_TextChanged({
        $script:Role = $txtRole.Text
        Request-DebouncedConfigSave
        & $optionReloadHandler
    })
$txtBackupContact.Add_TextChanged({
        $script:BackupContact = $txtBackupContact.Text
        Request-DebouncedConfigSave
        & $optionReloadHandler
    })
$txtBackupEngineerEmail.Add_TextChanged({
        $script:BackupEngineerEmail = $txtBackupEngineerEmail.Text
        Request-DebouncedConfigSave
        & $optionReloadHandler
    })
$txtTeamAlias.Add_TextChanged({
        $script:TeamAlias = $txtTeamAlias.Text
        Request-DebouncedConfigSave
        & $optionReloadHandler
    })
$txtSupportLink.Add_TextChanged({
        $script:SupportLink = $txtSupportLink.Text
        Request-DebouncedConfigSave
        & $optionReloadHandler
    })

# Override Account checkbox: enable/disable account text box
$chkOverrideAccount.Add_Checked({
        $script:OverrideAccount = $true
        $txtAccount.IsEnabled = $true
        Request-DebouncedConfigSave
    })
$chkOverrideAccount.Add_Unchecked({
        $script:OverrideAccount = $false
        $txtAccount.IsEnabled = $false
        Resolve-UserAlias
        $txtAccount.Text = $script:UserAlias
        Request-DebouncedConfigSave
        & $optionReloadHandler
    })
# Save edited account while typing, then persist after a brief pause
$txtAccount.Add_TextChanged({
        if ($chkOverrideAccount.IsChecked) {
            $script:UserAlias = $txtAccount.Text
            Request-DebouncedConfigSave
        }
    })

# Save edited account when user tabs out
$txtAccount.Add_LostFocus({
        if ($chkOverrideAccount.IsChecked) {
            $script:UserAlias = $txtAccount.Text
            Request-DebouncedConfigSave
            & $optionReloadHandler
        }
    })

# ===================== Template Editor: Auto-Save and Live Preview =====================
# Flag to suppress auto-save when text is set programmatically (e.g. on config change).
$script:IsLoadingTemplate = $false

# Debounced timer for template auto-save and preview refresh.
# Updates preview in real-time and saves the current template after user stops editing.
$script:TemplateEditTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:TemplateEditTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:TemplateEditTimer.Add_Tick({
    $timer = $this
    if ($null -ne $timer) {
        try { $timer.Stop() } catch { }
    }
    
    # Update preview in real-time
    $html = $txtMessage.Text
    if ([string]::IsNullOrWhiteSpace($html)) {
        $html = "<html><body><p style='color:#888;font-family:Segoe UI;'>No message to preview.</p></body></html>"
    }
    # Only update preview if we're on the Preview tab to avoid unnecessary rendering
    if ($tcMessageView.SelectedIndex -eq 1) {
        $wbPreview.NavigateToString($html)
    }
    
    # Auto-save the current template
    $selected = $cmbTemplate.SelectedItem.Content
    if ($selected -and $selected -ne "Custom..." -and $selected -ne "") {
        $path = Resolve-TemplateFilePath $selected
        if ($path) {
            try {
                Export-MessageToFile $path $txtMessage.Text
                Update-StatusBar "💾 Template '$selected' auto-saved"
            }
            catch {
                # Silently fail on auto-save to avoid interrupting the user
            }
        }
    }
})

function Request-DebouncedTemplateSave {
    $script:TemplateEditTimer.Stop()
    $script:TemplateEditTimer.Start()
}

# Auto-save and live preview when message text changes — skip when loading programmatically
$txtMessage.Add_TextChanged({
    if (-not $script:IsLoadingTemplate) {
        Request-DebouncedTemplateSave
    }
})

# ===================== HTML Formatting Toolbar Handlers =====================
# Add-HtmlTag: Wrap selected text in an HTML tag, or insert the tag pair at the cursor.
function Add-HtmlTag($openTag, $closeTag) {
    $selStart = $txtMessage.SelectionStart
    $selLen = $txtMessage.SelectionLength
    if ($selLen -gt 0) {
        $selected = $txtMessage.Text.Substring($selStart, $selLen)
        $replacement = "$openTag$selected$closeTag"
        $txtMessage.Text = $txtMessage.Text.Remove($selStart, $selLen).Insert($selStart, $replacement)
        $txtMessage.SelectionStart = $selStart
        $txtMessage.SelectionLength = $replacement.Length
    }
    else {
        $insert = "$openTag$closeTag"
        $txtMessage.Text = $txtMessage.Text.Insert($selStart, $insert)
        $txtMessage.SelectionStart = $selStart + $openTag.Length
    }
    $txtMessage.Focus()
}

# Add-HtmlSnippet: Insert an HTML snippet at the current cursor position.
function Add-HtmlSnippet($snippet) {
    $selStart = $txtMessage.SelectionStart
    $txtMessage.Text = $txtMessage.Text.Insert($selStart, $snippet)
    $txtMessage.SelectionStart = $selStart + $snippet.Length
    $txtMessage.Focus()
}

$btnFmtBold.Add_Click({ Add-HtmlTag '<b>' '</b>' })
$btnFmtItalic.Add_Click({ Add-HtmlTag '<i>' '</i>' })
$btnFmtUnderline.Add_Click({ Add-HtmlTag '<u>' '</u>' })
$btnFmtH3.Add_Click({ Add-HtmlTag '<h3>' '</h3>' })
$btnFmtP.Add_Click({ Add-HtmlTag '<p>' '</p>' })
$btnFmtBr.Add_Click({ Add-HtmlSnippet '<br/>' })

$btnFmtLink.Add_Click({
        $selStart = $txtMessage.SelectionStart
        $selLen = $txtMessage.SelectionLength
        $linkText = if ($selLen -gt 0) { $txtMessage.Text.Substring($selStart, $selLen) } else { 'link text' }
        $snippet = "<a href=`"https://`">$linkText</a>"
        if ($selLen -gt 0) {
            $txtMessage.Text = $txtMessage.Text.Remove($selStart, $selLen).Insert($selStart, $snippet)
        }
        else {
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
        }
        else {
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
        }
        else {
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
                if ((Show-TemplateWarningDialog $warnings) -ne 'Yes') { return }
            }
            Assert-ExchangeConnection
            Update-StatusBar "Applying internal message..."
            Set-AutoReplyMessage $txtMessage.Text 'Internal'
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
                if ((Show-TemplateWarningDialog $warnings) -ne 'Yes') { return }
            }
            Assert-ExchangeConnection
            Update-StatusBar "Applying external message..."
            Set-AutoReplyMessage $txtMessage.Text 'External'
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
                if ((Show-TemplateWarningDialog $warnings) -ne 'Yes') { return }
            }
            Assert-ExchangeConnection
            Update-StatusBar "Applying message to both internal and external..."
            Set-AutoReplyMessage $txtMessage.Text 'Both'
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
        }
        else {
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

# Backup Message: Save the current editor content as a new backup template file
# with a timestamp-based filename (e.g., backup_2024-01-15_14-30-45.html)
$btnBackupMessage.Add_Click({
        try {
            $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
            $backupFileName = "backup_$timestamp.html"
            $backupPath = Join-Path $ConfigDir $backupFileName
            Export-MessageToFile $backupPath $txtMessage.Text
            Update-StatusBar "Message backed up to $backupFileName"
            Show-InfoDialog "Backed Up" "Message saved as template:`n$backupFileName"
        }
        catch {
            Show-ErrorDialog "Error" "Could not backup message.`n$($_.Exception.Message)"
            Update-StatusBar "Failed to backup message"
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
Update-ConnectionUiState -State Disconnected
Update-ScheduledTaskStatusUI

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

# Holiday picker: Set the return date and holiday name when a holiday is selected.
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
# Uses screen capture (CopyFromScreen) so WebBrowser content is included.
Add-Type -AssemblyName System.Drawing
$ScreenshotsDir = Join-Path $ScriptDir "screenshots"

function Wait-WebBrowserReady($browser, [int]$timeoutMs = 5000) {
    # Wait for the WebBrowser's LoadCompleted event before proceeding.
    $handler = [System.Windows.Navigation.LoadCompletedEventHandler] { $script:_wbLoaded = $true }
    $script:_wbLoaded = $false
    $browser.Add_LoadCompleted($handler)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not $script:_wbLoaded -and $sw.ElapsedMilliseconds -lt $timeoutMs) {
        $Window.Dispatcher.Invoke([action] {}, [Windows.Threading.DispatcherPriority]::Background)
        Start-Sleep -Milliseconds 50
    }
    $browser.Remove_LoadCompleted($handler)
    # Extra render pass to ensure the layout is painted
    $Window.Dispatcher.Invoke([action] {}, [Windows.Threading.DispatcherPriority]::Render)
    Start-Sleep -Milliseconds 300
}

function Save-WindowScreenshot($filePath) {
    # Flush WPF render queue and let the window paint
    $Window.Dispatcher.Invoke([action] {}, [Windows.Threading.DispatcherPriority]::Render)
    Start-Sleep -Milliseconds 300

    # Get window position and size in physical pixels
    $source = [System.Windows.PresentationSource]::FromVisual($Window)
    [double]$dpiX = $source.CompositionTarget.TransformToDevice.M11
    [double]$dpiY = $source.CompositionTarget.TransformToDevice.M22

    [int]$left = [Math]::Round($Window.Left * $dpiX)
    [int]$top = [Math]::Round($Window.Top * $dpiY)
    [int]$width = [Math]::Round($Window.ActualWidth * $dpiX)
    [int]$height = [Math]::Round($Window.ActualHeight * $dpiY)

    # Capture from screen — includes WebBrowser and all hosted Win32 content
    $bmp = New-Object System.Drawing.Bitmap($width, $height)
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    $gfx.CopyFromScreen($left, $top, 0, 0, (New-Object System.Drawing.Size($width, $height)))
    $gfx.Dispose()
    $bmp.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

$Window.Add_KeyDown({
        if ($_.Key -eq 'F12') {
            $_.Handled = $true
            # Screenshot capture is disabled by default. Enable via "EnableScreenshots": true in config.json.
            $cfg = @{}
            $cfgPath = Join-Path $ScriptDir "config\config.json"
            if (Test-Path $cfgPath) { $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json }
            if (-not ($cfg.PSObject.Properties.Name -contains 'EnableScreenshots' -and $cfg.EnableScreenshots -eq $true)) { return }
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

                # Tab 2: Automation
                $tcMain.SelectedIndex = 2
                Save-WindowScreenshot (Join-Path $ScreenshotsDir "automation.png")

                # Tab 3: Message Templates — Edit
                $tcMain.SelectedIndex = 3
                $tcMessageView.SelectedIndex = 0
                Save-WindowScreenshot (Join-Path $ScreenshotsDir "message-templates-edit.png")

                # Tab 3: Message Templates — Preview
                $tcMessageView.SelectedIndex = 1
                Wait-WebBrowserReady $wbPreview
                Save-WindowScreenshot (Join-Path $ScreenshotsDir "message-templates-preview.png")

                # Tab 4: Current OOF
                $tcMain.SelectedIndex = 4
                Wait-WebBrowserReady $wbCurrentOOF
                Save-WindowScreenshot (Join-Path $ScreenshotsDir "current-oof.png")

                # Restore original tab positions
                $tcMessageView.SelectedIndex = $originalSubTab
                $tcMain.SelectedIndex = $originalTab

                Update-StatusBar "Screenshots saved to screenshots/ folder"
                Show-InfoDialog "Screenshots Captured" "6 screenshots saved to:`n$ScreenshotsDir`n`n- quick-actions.png`n- configuration.png`n- automation.png`n- message-templates-edit.png`n- message-templates-preview.png`n- current-oof.png"
            }
            catch {
                Show-ErrorDialog "Screenshot Error" "Failed to capture screenshots:`n$($_.Exception.Message)"
                Update-StatusBar "Screenshot capture failed"
            }
        }
    })

# ===================== Show the Window =====================
# Display the WPF window until it is closed.
# On close, release the Exchange Online session and warn if OOF is currently disabled.

# Silent startup update check — notify via status bar if a newer version exists
$txtLocalVersion.Text = $script:ScriptVersion
try {
    $remoteVer = Get-RemoteScriptVersion
    $txtRemoteVersion.Text = $remoteVer
    $updateState = Get-UpdateVersionState -RemoteVersion $remoteVer -LocalVersion $script:ScriptVersion
    switch ($updateState) {
        'RemoteNewer' {
            $txtRemoteVersion.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xD8, 0x3B, 0x01))
            Update-StatusBar "Update available: v$remoteVer — go to Configuration > Check for Updates"
        }
        'LocalNewer' {
            $txtRemoteVersion.Foreground = [System.Windows.Media.Brushes]::Green
            Update-StatusBar "Local version is newer than GitHub (v$remoteVer)."
        }
        'Unknown' {
            $txtRemoteVersion.Foreground = [System.Windows.Media.Brushes]::DarkOrange
            Update-StatusBar "Could not determine GitHub version at startup."
        }
        default {
            $txtRemoteVersion.Foreground = [System.Windows.Media.Brushes]::Green
        }
    }
}
catch { }
$Window.Add_Closing({
        if ($script:IsConnectedToEXO -and -not $script:OOFReplyEnabled) {
            $result = [System.Windows.MessageBox]::Show(
                "Your Out of Office reply is not currently enabled.`n`nAre you sure you want to exit without enabling it?",
                "OOF Not Enabled",
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
        # Clean up background update check job and timer
        if ($script:UpdateCheckTimer) {
            try { $script:UpdateCheckTimer.Stop() } catch { }
        }
        if ($script:UpdateCheckJob -and $script:UpdateCheckJob.State -eq 'Running') {
            $script:UpdateCheckJob | Stop-Job -Force -ErrorAction SilentlyContinue
        }
    })

# ===================== Background Update Check =====================
# Check for script updates once when the GUI starts.
# If an update is available, highlight the Check for Updates button.
$script:UpdateCheckJob = Start-Job -ScriptBlock {
    param($UpdateUrl, $LocalVersion)
    try {
        $tempFile = [System.IO.Path]::GetTempFileName()
        Invoke-WebRequest -Uri $UpdateUrl -OutFile $tempFile -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop -Headers @{ 'Cache-Control' = 'no-cache' }
        $line = Select-String -Path $tempFile -Pattern '^\$script:ScriptVersion\s*=\s*"(.+)"' | Select-Object -First 1
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        $remoteVersion = if ($line) { $line.Matches[0].Groups[1].Value } else { $null }

        # Only flag an update when remote version is strictly newer than local.
        $isNewer = $false
        if ($remoteVersion) {
            try {
                $isNewer = ([version]$remoteVersion -gt [version]$LocalVersion)
            }
            catch {
                $isNewer = $false
            }
        }
        if ($isNewer) {
            Write-Output "UPDATE_AVAILABLE"
        }
    }
    catch {
        # Silently skip on error
    }
} -ArgumentList $ScriptUpdateUrl, $script:ScriptVersion

# Track whether an update has been signaled
$script:UpdateSignaled = $false

# Set up a timer to check for background job completion once
$script:UpdateCheckTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:UpdateCheckTimer.Interval = [TimeSpan]::FromMilliseconds(500)  # Check every 500ms until job completes
$updateCheckTimer_Tick = {
    $timer = $this
    $job = $script:UpdateCheckJob
    if ($null -eq $job) {
        if ($null -ne $timer) {
            try { $timer.Stop() } catch { }
        }
        return
    }

    if (-not $script:UpdateSignaled -and $job.State -ne 'Running') {
        if ($null -ne $timer) {
            try { $timer.Stop() } catch { }
        }
        $jobOutput = $job | Receive-Job -ErrorAction SilentlyContinue
        if ($jobOutput -contains "UPDATE_AVAILABLE") {
            $script:UpdateSignaled = $true
            # Highlight and focus the Check for Updates button
            $btnCheckForUpdates.Background = [System.Windows.Media.Brushes]::LightYellow
            $btnCheckForUpdates.Foreground = [System.Windows.Media.Brushes]::Red
            $btnCheckForUpdates.Content = "📥 Update Available!"
            $btnCheckForUpdates.Focus()
        }
    }
}
$script:UpdateCheckTimer.Add_Tick($updateCheckTimer_Tick)
$script:UpdateCheckTimer.Start()

# First-run auto-connect: if no config.json existed at startup, automatically trigger the
# Connect flow once the window is visible so EXO can populate the Full Name and profile fields.
$Window.Add_ContentRendered({
    if ($script:IsFirstRun) {
        Update-StatusBar "First run detected — connecting to Exchange Online to populate your profile..."
        $btnConnect.RaiseEvent(
            [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)
        )
    }
})

$Window.ShowDialog() | Out-Null

# Final cleanup: ensure Exchange session is released even if the Closed event didn't fire.
try { Disconnect-ExchangeOnlineSession } catch { }
