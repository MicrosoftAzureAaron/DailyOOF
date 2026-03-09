param([string]$InputParm)

# ===================== WPF GUI Setup =====================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigDir = Join-Path $ScriptDir "config"
$ConfigFile = Join-Path $ConfigDir "config.json"
$XamlFile = Join-Path $ConfigDir "AAOOF-GUI.xaml"

# Ensure config directory exists
if (!(Test-Path $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir | Out-Null }

# ===================== Auto-Download Missing Config Files =====================
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
$global:StartOfShift = $null
$global:EndOfShift = $null
$global:WorkDays = $null
$global:UserAlias = ""
$global:UserAliasSuffix = "@microsoft.com"
$global:UserSignature = ""

function Import-ConfigFromFile {
    if (Test-Path $ConfigFile) {
        $cfg = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        if ($cfg.StartOfShift)    { $global:StartOfShift = [datetime]$cfg.StartOfShift }
        if ($cfg.EndOfShift)      { $global:EndOfShift = [datetime]$cfg.EndOfShift }
        if ($cfg.WorkDays)        { $global:WorkDays = @($cfg.WorkDays) }
        if ($cfg.UserAlias)       { $global:UserAlias = $cfg.UserAlias }
        if ($cfg.UserAliasSuffix) { $global:UserAliasSuffix = $cfg.UserAliasSuffix }
    }
}

# Load config immediately on script start
Import-ConfigFromFile

# ===================== Core Functions =====================

function Get-UserAlias {
    if ([string]::IsNullOrEmpty($global:UserAliasSuffix)) {
        $global:UserAliasSuffix = "@microsoft.com"
    }
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem
    if ($cs.Username) {
        $CurrentUser = $cs.Username.Split('\')[-1]
    } else {
        $CurrentUser = $env:USERNAME
    }
    $global:UserAlias = "$CurrentUser$global:UserAliasSuffix"
}

function Get-ARCFilePath {
    return Join-Path $ConfigDir "AutoReplyConfig.json"
}

function Set-ARCFile {
    $ARCFilePath = Get-ARCFilePath
    Get-ARC | ConvertTo-Json -Depth 100 | Set-Content $ARCFilePath
}

function Get-ARC {
    return Get-MailboxAutoReplyConfiguration -Identity $global:UserAlias
}

function Get-ARCFile {
    $ARCFilePath = Get-ARCFilePath
    return Get-Content $ARCFilePath -Raw | ConvertFrom-Json
}

function Set-ARCState($S) {
    switch ($S) {
        'Enabled'   { Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -AutoReplyState "Enabled" }
        'Disabled'  { Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -AutoReplyState "Disabled" }
        'Scheduled' {
            Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -AutoReplyState "Scheduled"
            Set-ARCTimes
        }
    }
    Set-ARCFile
}

function Set-ARCTimes {
    if ($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift) { return }
    if ($null -eq $global:WorkDays) { return }

    $daysToAdd = Get-NextWorkDay

    $hours = Get-Date $global:StartOfShift
    $startTime = [datetime](Get-Date).Date.AddHours($hours.Hour)
    $startTime = $startTime.AddDays($daysToAdd)

    $hours = Get-Date $global:EndOfShift
    $endTime = [datetime](Get-Date).Date.AddHours($hours.Hour)

    if ($daysToAdd -eq 0) { $endTime = $endTime.AddDays(-1) }

    Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -StartTime $endTime -EndTime $startTime
    Set-ARCFile
}

function Set-ARCMessage($message, $IOE) {
    switch ($IOE) {
        'Internal' { Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -InternalMessage $message }
        'External' { Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -ExternalMessage $message }
        default    { Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -ExternalMessage $message -InternalMessage $message }
    }
}

function Get-NextWorkDay {
    if ($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift) { return 1 }
    if (!$global:WorkDays) { return 1 }

    $CTime = [datetime](Get-Date)

    if (!($CTime.DayOfWeek -in $global:WorkDays)) {
        $i = 0
        while (!($CTime.DayOfWeek -in $global:WorkDays)) {
            $i += 1
            $CTime = $CTime.AddDays(1)
        }
        return $i
    }
    else {
        $CTime2 = $CTime.AddDays(1)
        $i = 1
        while (!($CTime2.DayOfWeek -in $global:WorkDays)) {
            $i += 1
            $CTime2 = $CTime2.AddDays(1)
        }
        if ($i -gt 1) {
            return $i
        }

        $CTime = [datetime](Get-Date)
        if ($CTime -lt $global:StartOfShift) {
            return 0
        }
        else {
            return 1
        }
    }
}

function Get-EXOConnection {
    if ([string]::IsNullOrEmpty($global:UserAlias)) { Get-UserAlias }

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
    Connect-ExchangeOnline -UserPrincipalName $global:UserAlias
    return $true
}

function Set-EXODisconnect {
    try { Disconnect-ExchangeOnline -Confirm:$false } catch { }
}

function Save-MessageToFile($filePath, $content) {
    $content | Out-File -FilePath $filePath -Encoding utf8
}

function Get-VacationDate($returnDate) {
    if ($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift) { return }
    $parsedDate = [datetime]$returnDate
    $endTime = $parsedDate + $global:StartOfShift.TimeOfDay
    Set-MailboxAutoReplyConfiguration -Identity $global:UserAlias -AutoReplyState "Scheduled" -StartTime $global:EndOfShift -EndTime $endTime
    Set-ARCFile
}

function Set-DailyScriptTask {
    if (!(Get-ScheduledTask -TaskName "AAOOF" -ErrorAction SilentlyContinue)) {
        $scriptPath = Join-Path $ScriptDir "AAOOF-GUI.ps1"
        $taskname = "AAOOF"
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "`"$scriptPath`" 1"
        $date = Get-Date -Date (Get-Date).Date
        $TriggerTime = $global:StartOfShift.TimeOfDay
        $TriggerTime = $date.AddMinutes(15) + $TriggerTime
        $trigger = New-ScheduledTaskTrigger -Daily -At $TriggerTime

        try {
            Register-ScheduledTask -TaskName $taskname -Trigger $trigger -Action $action -RunLevel Highest -ErrorAction Stop
        }
        catch {
            # Not running as admin — elevate and retry
            $triggerTimeStr = $TriggerTime.ToString("o")
            $elevatedCmd = @"
\`$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '`"$scriptPath`" 1'
\`$trigger = New-ScheduledTaskTrigger -Daily -At '$triggerTimeStr'
Register-ScheduledTask -TaskName '$taskname' -Trigger \`$trigger -Action \`$action -RunLevel Highest -ErrorAction Stop
"@
            $encodedCmd = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($elevatedCmd))
            $proc = Start-Process powershell.exe -ArgumentList "-NoProfile -EncodedCommand $encodedCmd" -Verb RunAs -Wait -PassThru
            if ($proc.ExitCode -ne 0) {
                throw "Scheduled task creation failed even after elevation."
            }
        }
        return $true
    }
    return $false
}

function Save-ConfigToFile {
    $cfg = @{
        StartOfShift    = if ($null -ne $global:StartOfShift) { $global:StartOfShift.ToString("o") } else { $null }
        EndOfShift      = if ($null -ne $global:EndOfShift) { $global:EndOfShift.ToString("o") } else { $null }
        WorkDays        = $global:WorkDays
        UserAlias       = $global:UserAlias
        UserAliasSuffix = $global:UserAliasSuffix
    }
    $cfg | ConvertTo-Json -Depth 5 | Set-Content $ConfigFile -Encoding utf8
}

# ===================== CLI Mode (for scheduled task / automation) =====================
if ($InputParm) {
    if ($InputParm -eq '1') {
        if ($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift -or $null -eq $global:WorkDays) {
            Write-Host "Configuration not set. Please run the GUI first to configure." -ForegroundColor Red
            exit
        }
        Get-EXOConnection

        # Skip if a vacation/extended OOF is active (end time is more than 1 day out)
        $arc = Get-ARC
        if ($arc.AutoReplyState -eq 'Scheduled') {
            $endTime = [datetime]$arc.EndTime
            if ($endTime -gt (Get-Date).AddDays(1)) {
                Write-Host "Vacation/extended OOF is active until $endTime — skipping daily update." -ForegroundColor Yellow
                Set-EXODisconnect
                exit
            }
        }

        Set-ARCState 'Scheduled'
        $arc = Get-ARC
        Write-Host "Auto Reply: $($arc.AutoReplyState) | Start: $($arc.StartTime) | End: $($arc.EndTime)"
        Set-EXODisconnect
        exit
    }
    if ($InputParm -as [datetime]) {
        if ($null -eq $global:StartOfShift -or $null -eq $global:EndOfShift) {
            Write-Host "Configuration not set. Please run the GUI first to configure." -ForegroundColor Red
            exit
        }
        Get-EXOConnection
        Get-VacationDate $InputParm
        $arc = Get-ARC
        Write-Host "Auto Reply: $($arc.AutoReplyState) | Start: $($arc.StartTime) | End: $($arc.EndTime)"
        Set-EXODisconnect
        exit
    }
}

# ===================== Load XAML GUI from File =====================
if (!(Test-Path $XamlFile)) {
    Write-Host "FATAL: XAML file not found at $XamlFile" -ForegroundColor Red
    exit 1
}
[xml]$XAML = Get-Content $XamlFile -Raw

# ===================== Build the Window =====================
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Get all named controls
$txtAccount = $Window.FindName("txtAccount")
$txtConnectionStatus = $Window.FindName("txtConnectionStatus")
$btnConnect = $Window.FindName("btnConnect")
$btnDisconnect = $Window.FindName("btnDisconnect")
$btnEnableScheduled = $Window.FindName("btnEnableScheduled")
$dpReturnDate = $Window.FindName("dpReturnDate")
$btnSetVacation = $Window.FindName("btnSetVacation")
$txtARCState = $Window.FindName("txtARCState")
$txtARCStart = $Window.FindName("txtARCStart")
$txtARCEnd = $Window.FindName("txtARCEnd")
$btnRefreshStatus = $Window.FindName("btnRefreshStatus")
$btnViewCurrentMsg = $Window.FindName("btnViewCurrentMsg")
$txtSuffix = $Window.FindName("txtSuffix")
$btnSaveSuffix = $Window.FindName("btnSaveSuffix")
$cmbStartHour = $Window.FindName("cmbStartHour")
$cmbStartMin = $Window.FindName("cmbStartMin")
$cmbStartAmPm = $Window.FindName("cmbStartAmPm")
$cmbEndHour = $Window.FindName("cmbEndHour")
$cmbEndMin = $Window.FindName("cmbEndMin")
$cmbEndAmPm = $Window.FindName("cmbEndAmPm")
$btnSaveHours = $Window.FindName("btnSaveHours")
$chkMon = $Window.FindName("chkMon")
$chkTue = $Window.FindName("chkTue")
$chkWed = $Window.FindName("chkWed")
$chkThu = $Window.FindName("chkThu")
$chkFri = $Window.FindName("chkFri")
$chkSat = $Window.FindName("chkSat")
$chkSun = $Window.FindName("chkSun")
$btnSaveDays = $Window.FindName("btnSaveDays")
$btnPresetMF = $Window.FindName("btnPresetMF")
$btnPresetSunWed = $Window.FindName("btnPresetSunWed")
$btnPresetWedSat = $Window.FindName("btnPresetWedSat")
$btnStateEnabled = $Window.FindName("btnStateEnabled")
$btnStateDisabled = $Window.FindName("btnStateDisabled")
$btnStateScheduled = $Window.FindName("btnStateScheduled")
$btnCreateTask = $Window.FindName("btnCreateTask")
$cmbTemplate = $Window.FindName("cmbTemplate")
$btnLoadTemplate = $Window.FindName("btnLoadTemplate")
$btnBrowseFile = $Window.FindName("btnBrowseFile")
$txtMessage = $Window.FindName("txtMessage")
$btnApplyInternal = $Window.FindName("btnApplyInternal")
$btnApplyExternal = $Window.FindName("btnApplyExternal")
$btnApplyBoth = $Window.FindName("btnApplyBoth")
$btnSaveTemplate = $Window.FindName("btnSaveTemplate")
$btnSaveOnlineMsg = $Window.FindName("btnSaveOnlineMsg")
$chkIncludeSignature = $Window.FindName("chkIncludeSignature")
$chkIncludeOfficeHours = $Window.FindName("chkIncludeOfficeHours")
$chkIncludeWorkDays = $Window.FindName("chkIncludeWorkDays")
$chkIncludeTimezone = $Window.FindName("chkIncludeTimezone")
$tcMessageView = $Window.FindName("tcMessageView")
$wbPreview = $Window.FindName("wbPreview")
$txtStatusBar = $Window.FindName("txtStatusBar")

# ===================== Helper: Update Status Bar =====================
function Update-Status($msg) {
    $txtStatusBar.Text = $msg
    $Window.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
}

function Show-Result($title, $msg) {
    [System.Windows.MessageBox]::Show($msg, $title, 'OK', 'Information')
}

function Show-Error($title, $msg) {
    [System.Windows.MessageBox]::Show($msg, $title, 'OK', 'Error')
}

# ===================== Populate Combos =====================
1..12 | ForEach-Object { $cmbStartHour.Items.Add($_.ToString()) | Out-Null; $cmbEndHour.Items.Add($_.ToString()) | Out-Null }
@("00","15","30","45") | ForEach-Object { $cmbStartMin.Items.Add($_) | Out-Null; $cmbEndMin.Items.Add($_) | Out-Null }
@("AM","PM") | ForEach-Object { $cmbStartAmPm.Items.Add($_) | Out-Null; $cmbEndAmPm.Items.Add($_) | Out-Null }

# ===================== Load Saved Config into UI =====================
function Import-UIFromConfig {
    # Suffix
    if (![string]::IsNullOrEmpty($global:UserAliasSuffix)) {
        $txtSuffix.Text = $global:UserAliasSuffix
    }
    # Account
    if (![string]::IsNullOrEmpty($global:UserAlias)) {
        $txtAccount.Text = $global:UserAlias
    } else {
        Get-UserAlias
        $txtAccount.Text = $global:UserAlias
    }

    # Shift times
    if ($null -ne $global:StartOfShift) {
        $h = (Get-Date $global:StartOfShift).Hour
        $m = (Get-Date $global:StartOfShift).Minute
        $ampm = if ($h -ge 12) { "PM" } else { "AM" }
        $displayH = if ($h -gt 12) { $h - 12 } elseif ($h -eq 0) { 12 } else { $h }
        $cmbStartHour.SelectedItem = $displayH.ToString()
        $nearestMin = @("00","15","30","45") | Sort-Object { [Math]::Abs([int]$_ - $m) } | Select-Object -First 1
        $cmbStartMin.SelectedItem = $nearestMin
        $cmbStartAmPm.SelectedItem = $ampm
    } else {
        $cmbStartHour.SelectedIndex = 8; $cmbStartMin.SelectedIndex = 0; $cmbStartAmPm.SelectedIndex = 0 # 9 AM
    }

    if ($null -ne $global:EndOfShift) {
        $h = (Get-Date $global:EndOfShift).Hour
        $m = (Get-Date $global:EndOfShift).Minute
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
    if ($global:WorkDays) {
        $chkMon.IsChecked = ('Monday' -in $global:WorkDays)
        $chkTue.IsChecked = ('Tuesday' -in $global:WorkDays)
        $chkWed.IsChecked = ('Wednesday' -in $global:WorkDays)
        $chkThu.IsChecked = ('Thursday' -in $global:WorkDays)
        $chkFri.IsChecked = ('Friday' -in $global:WorkDays)
        $chkSat.IsChecked = ('Saturday' -in $global:WorkDays)
        $chkSun.IsChecked = ('Sunday' -in $global:WorkDays)
    } else {
        # Default Mon-Fri
        $chkMon.IsChecked = $true; $chkTue.IsChecked = $true; $chkWed.IsChecked = $true
        $chkThu.IsChecked = $true; $chkFri.IsChecked = $true
    }
}

function Get-TemplateFilePath($templateName) {
    switch ($templateName) {
        "Normal OOF"   { return Join-Path $ConfigDir "normal_oof.html" }
        "Vacation OOF" { return Join-Path $ConfigDir "vacation_oof.html" }
        "Sick OOF"     { return Join-Path $ConfigDir "sick_oof.html" }
        "Holiday OOF"  { return Join-Path $ConfigDir "holiday_oof.html" }
        default        { return $null }
    }
}

function Resolve-TemplatePlaceholders($text) {
    if ($null -ne $global:StartOfShift -and $null -ne $global:EndOfShift) {
        $hours = "$($global:StartOfShift.ToString('h:mm tt')) - $($global:EndOfShift.ToString('h:mm tt'))"
    } else {
        $hours = "not configured"
    }
    if ($global:WorkDays) {
        $days = ($global:WorkDays -join ', ')
    } else {
        $days = "not configured"
    }
    $tz = [System.TimeZoneInfo]::Local.DisplayName

    # Office hours — include or strip
    if ($chkIncludeOfficeHours.IsChecked) {
        $text = $text -replace '\[OFFICE HOURS\]', $hours
    } else {
        $text = $text -replace '\[OFFICE HOURS\]', ''
    }

    # Work days
    if ($chkIncludeWorkDays.IsChecked) {
        $text = $text -replace '\[WORK DAYS\]', $days
    } else {
        $text = $text -replace '\[WORK DAYS\]', ''
    }

    # Timezone
    if ($chkIncludeTimezone.IsChecked) {
        $text = $text -replace '\[TIMEZONE\]', $tz
    } else {
        $text = $text -replace '\[TIMEZONE\]', ''
    }

    # Rewrite or remove the footer line based on which options are active
    $hasHours = $chkIncludeOfficeHours.IsChecked
    $hasDays  = $chkIncludeWorkDays.IsChecked
    $hasTZ    = $chkIncludeTimezone.IsChecked

    if (!$hasHours -and !$hasDays -and !$hasTZ) {
        # All unchecked — remove the entire footer line
        $text = $text -replace '(?m)^\s*<p[^>]*>My regular office hours are.*?</p>\s*\r?\n?', ''
    }
    elseif (!$hasHours -and $hasDays -and !$hasTZ) {
        # Only work days — rewrite to "My days in office are ..."
        $text = $text -replace '(?m)(<p[^>]*>)My regular office hours are\s*(<b>\s*</b>\s*)*,?\s*(<b>[^<]+</b>)\s*\.(</p>)', '$1My days in office are $3.$4'
    }
    elseif (!$hasHours -and !$hasDays -and $hasTZ) {
        # Only timezone — rewrite to "My timezone is ..."
        $text = $text -replace '(?m)(<p[^>]*>)My regular office hours are\s*(<b>[^<]+</b>)\s*,?\s*(<b>\s*</b>)?\s*\.(</p>)', '$1My timezone is $2.$4'
    }
    else {
        # Clean up empty <b></b> tags and stray commas/spaces left by disabled options
        $text = $text -replace '<b>\s*</b>', ''
        $text = $text -replace '(?m)(<p[^>]*>My regular office hours are)\s+,', '$1'
        $text = $text -replace ',\s*\.', '.'
        $text = $text -replace '\s{2,}', ' '
    }

    # Signature — include or strip (also remove "Best regards" when signature is excluded)
    if ($chkIncludeSignature.IsChecked -and ![string]::IsNullOrWhiteSpace($global:UserSignature)) {
        $text = $text -replace '\[SIGNATURE\]', $global:UserSignature
    } else {
        $text = $text -replace '(?m)^\s*<p>\s*Best regards\s*</p>\s*\r?\n?', ''
        $text = $text -replace '(?m)^\s*\[SIGNATURE\]\s*\r?\n?', ''
    }
    return $text
}

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

function Read-ShiftTimesFromUI {
    $sH = [int]$cmbStartHour.SelectedItem
    $sM = [int]$cmbStartMin.SelectedItem
    $sAP = $cmbStartAmPm.SelectedItem
    if ($sAP -eq "PM" -and $sH -ne 12) { $sH += 12 }
    if ($sAP -eq "AM" -and $sH -eq 12) { $sH = 0 }
    $global:StartOfShift = [datetime](Get-Date).Date.AddHours($sH).AddMinutes($sM)

    $eH = [int]$cmbEndHour.SelectedItem
    $eM = [int]$cmbEndMin.SelectedItem
    $eAP = $cmbEndAmPm.SelectedItem
    if ($eAP -eq "PM" -and $eH -ne 12) { $eH += 12 }
    if ($eAP -eq "AM" -and $eH -eq 12) { $eH = 0 }
    $global:EndOfShift = [datetime](Get-Date).Date.AddHours($eH).AddMinutes($eM)
}

# ===================== Event Handlers =====================

# Connect
$btnConnect.Add_Click({
    try {
        Update-Status "Connecting to Exchange Online..."
        $global:UserAliasSuffix = $txtSuffix.Text
        Get-UserAlias
        $txtAccount.Text = $global:UserAlias
        Get-EXOConnection
        $txtConnectionStatus.Text = "Connected"
        $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Green

        # On first connect, pull current OOF config and message and save locally
        try {
            $arc = Get-ARC
            $txtARCState.Text = $arc.AutoReplyState
            $txtARCStart.Text = $arc.StartTime.ToString()
            $txtARCEnd.Text = $arc.EndTime.ToString()

            # Save the current online messages to template files if we don't already have a saved message
            $savedMsgFile = Join-Path $ConfigDir "message.html"
            if (!(Test-Path $savedMsgFile) -and ![string]::IsNullOrWhiteSpace($arc.ExternalMessage)) {
                Save-MessageToFile $savedMsgFile $arc.ExternalMessage
            }
        } catch { }

        # Fetch user's OWA email signature for template placeholder
        try {
            $msgConfig = Get-MailboxMessageConfiguration -Identity $global:UserAlias
            if (![string]::IsNullOrWhiteSpace($msgConfig.SignatureHtml)) {
                $global:UserSignature = $msgConfig.SignatureHtml
                $chkIncludeSignature.IsEnabled = $true
                $chkIncludeSignature.ToolTip = $null
            } else {
                $global:UserSignature = ""
                $chkIncludeSignature.IsChecked = $false
                $chkIncludeSignature.IsEnabled = $false
                $chkIncludeSignature.ToolTip = "No OWA signature found. Paste your signature into the message template and save it."
                Update-Status "Connected as $($global:UserAlias) — No OWA signature found"
            }
        } catch {
            $global:UserSignature = ""
            $chkIncludeSignature.IsChecked = $false
            $chkIncludeSignature.IsEnabled = $false
            $chkIncludeSignature.ToolTip = "No OWA signature found. Paste your signature into the message template and save it."
            Update-Status "Connected as $($global:UserAlias) — Could not retrieve signature"
        }

        Save-ConfigToFile
        Update-Status "Connected as $($global:UserAlias)"
    }
    catch {
        $txtConnectionStatus.Text = "Connection Failed"
        $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Red
        Show-Error "Connection Error" $_.Exception.Message
        Update-Status "Connection failed"
    }
})

# Disconnect
$btnDisconnect.Add_Click({
    try {
        Set-EXODisconnect
        $txtConnectionStatus.Text = "Disconnected"
        $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Red
        Update-Status "Disconnected from Exchange Online"
    }
    catch {
        Show-Error "Disconnect Error" $_.Exception.Message
    }
})

# Enable Scheduled Auto Reply
$btnEnableScheduled.Add_Click({
    try {
        Update-Status "Setting scheduled auto reply..."
        Read-ShiftTimesFromUI
        $global:WorkDays = Read-WorkDaysFromUI
        Set-ARCState 'Scheduled'
        $arc = Get-ARC
        $txtARCState.Text = $arc.AutoReplyState
        $txtARCStart.Text = $arc.StartTime.ToString()
        $txtARCEnd.Text = $arc.EndTime.ToString()
        Update-Status "Scheduled auto reply enabled"
        Show-Result "Success" "Scheduled Auto Reply enabled.`nStart: $($arc.StartTime)`nEnd: $($arc.EndTime)"
    }
    catch {
        Show-Error "Error" $_.Exception.Message
        Update-Status "Failed to set scheduled auto reply"
    }
})

# Set Vacation OOF
$btnSetVacation.Add_Click({
    try {
        if ($null -eq $dpReturnDate.SelectedDate) {
            Show-Error "Missing Date" "Please select a return date."
            return
        }
        Update-Status "Setting vacation OOF..."
        Read-ShiftTimesFromUI
        $returnDate = $dpReturnDate.SelectedDate.ToString("yyyy/MM/dd")
        Get-VacationDate $returnDate

        # If there's a vacation template loaded, apply it
        $vacPath = Get-TemplateFilePath "Vacation OOF"
        if ((Test-Path $vacPath) -and [string]::IsNullOrWhiteSpace($txtMessage.Text) -eq $false) {
            # User may want to apply the loaded message - handled separately via Apply buttons
        }

        $arc = Get-ARC
        $txtARCState.Text = $arc.AutoReplyState
        $txtARCStart.Text = $arc.StartTime.ToString()
        $txtARCEnd.Text = $arc.EndTime.ToString()
        Update-Status "Vacation OOF set until $returnDate"
        Show-Result "Success" "Vacation OOF enabled until $returnDate`nStart: $($arc.StartTime)`nEnd: $($arc.EndTime)"
    }
    catch {
        Show-Error "Error" $_.Exception.Message
        Update-Status "Failed to set vacation OOF"
    }
})

# Refresh Status
$btnRefreshStatus.Add_Click({
    try {
        Update-Status "Refreshing status..."
        $arc = Get-ARC
        $txtARCState.Text = $arc.AutoReplyState
        $txtARCStart.Text = $arc.StartTime.ToString()
        $txtARCEnd.Text = $arc.EndTime.ToString()
        Update-Status "Status refreshed"
    }
    catch {
        Show-Error "Error" "Could not refresh. Are you connected?`n$($_.Exception.Message)"
        Update-Status "Refresh failed"
    }
})

# View Current OOF Message
$btnViewCurrentMsg.Add_Click({
    try {
        Update-Status "Fetching current OOF message..."
        $arc = Get-ARC
        if ([string]::IsNullOrWhiteSpace($arc.InternalMessage) -and [string]::IsNullOrWhiteSpace($arc.ExternalMessage)) {
            Show-Result "Info" "No OOF message is currently set."
            Update-Status "No current OOF message"
            return
        }
        $txtMessage.Text = $arc.InternalMessage
        if ($tcMessageView.SelectedIndex -eq 1) {
            $wbPreview.NavigateToString($txtMessage.Text)
        }
        Update-Status "Current OOF message loaded into editor"
    }
    catch {
        Show-Error "Error" "Could not fetch message. Are you connected?`n$($_.Exception.Message)"
        Update-Status "Failed to fetch OOF message"
    }
})

# Save Suffix
$btnSaveSuffix.Add_Click({
    $global:UserAliasSuffix = $txtSuffix.Text
    Get-UserAlias
    $txtAccount.Text = $global:UserAlias
    Save-ConfigToFile
    Update-Status "Suffix saved: $($global:UserAliasSuffix)"
})

# Save Office Hours
$btnSaveHours.Add_Click({
    Read-ShiftTimesFromUI
    Save-ConfigToFile
    Update-Status "Office hours saved: $($global:StartOfShift.ToString('h:mm tt')) - $($global:EndOfShift.ToString('h:mm tt'))"
    Show-Result "Saved" "Office hours saved.`nStart: $($global:StartOfShift.ToString('h:mm tt'))`nEnd: $($global:EndOfShift.ToString('h:mm tt'))"
})

# Save Work Days
$btnSaveDays.Add_Click({
    $global:WorkDays = Read-WorkDaysFromUI
    Save-ConfigToFile
    Update-Status "Work days saved: $($global:WorkDays -join ', ')"
    Show-Result "Saved" "Work days saved: $($global:WorkDays -join ', ')"
})

# Preset Mon-Fri
$btnPresetMF.Add_Click({
    $chkMon.IsChecked = $true; $chkTue.IsChecked = $true; $chkWed.IsChecked = $true
    $chkThu.IsChecked = $true; $chkFri.IsChecked = $true
    $chkSat.IsChecked = $false; $chkSun.IsChecked = $false
})

# Preset Sun-Wed (4x10)
$btnPresetSunWed.Add_Click({
    $chkSun.IsChecked = $true; $chkMon.IsChecked = $true; $chkTue.IsChecked = $true
    $chkWed.IsChecked = $true; $chkThu.IsChecked = $false
    $chkFri.IsChecked = $false; $chkSat.IsChecked = $false
})

# Preset Wed-Sat (4x10)
$btnPresetWedSat.Add_Click({
    $chkWed.IsChecked = $true; $chkThu.IsChecked = $true
    $chkFri.IsChecked = $true; $chkSat.IsChecked = $true
    $chkSun.IsChecked = $false; $chkMon.IsChecked = $false; $chkTue.IsChecked = $false
})

# Auto Reply State buttons
$btnStateEnabled.Add_Click({
    try {
        Update-Status "Setting auto reply to Enabled..."
        Set-ARCState 'Enabled'
        Update-Status "Auto reply set to Enabled"
        Show-Result "Done" "Auto Reply State set to Enabled"
    }
    catch { Show-Error "Error" $_.Exception.Message }
})

$btnStateDisabled.Add_Click({
    try {
        Update-Status "Setting auto reply to Disabled..."
        Set-ARCState 'Disabled'
        Update-Status "Auto reply set to Disabled"
        Show-Result "Done" "Auto Reply State set to Disabled"
    }
    catch { Show-Error "Error" $_.Exception.Message }
})

$btnStateScheduled.Add_Click({
    try {
        Update-Status "Setting auto reply to Scheduled..."
        Read-ShiftTimesFromUI
        $global:WorkDays = Read-WorkDaysFromUI
        Set-ARCState 'Scheduled'
        Update-Status "Auto reply set to Scheduled"
        Show-Result "Done" "Auto Reply State set to Scheduled"
    }
    catch { Show-Error "Error" $_.Exception.Message }
})

# Create Scheduled Task
$btnCreateTask.Add_Click({
    try {
        Read-ShiftTimesFromUI
        $result = Set-DailyScriptTask
        if ($result) {
            Show-Result "Success" "Scheduled task 'AAOOF' created successfully."
            Update-Status "Scheduled task created"
        } else {
            Show-Result "Info" "Scheduled task 'AAOOF' already exists."
            Update-Status "Task already exists"
        }
    }
    catch {
        Show-Error "Error" "Failed to create task. Run as Administrator.`n$($_.Exception.Message)"
        Update-Status "Task creation failed - need admin"
    }
})

# Populate dynamic items (custom files, backups) when dropdown opens
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
    $selected = $cmbTemplate.SelectedItem.Content
    if ($selected -eq "Custom...") {
        Update-Status "Use 'Browse File...' to load a custom template"
        return
    }
    $path = if ($cmbTemplate.SelectedItem.Tag) { $cmbTemplate.SelectedItem.Tag } else { Get-TemplateFilePath $selected }
    if ($path -and (Test-Path $path)) {
        # Save current message to backup before overwriting
        if (![string]::IsNullOrWhiteSpace($txtMessage.Text)) {
            $backupFile = Join-Path $ConfigDir "message.html.bak"
            Save-MessageToFile $backupFile $txtMessage.Text
        }
        $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $path -Raw)
        # Refresh preview if on Preview tab
        if ($tcMessageView.SelectedIndex -eq 1) {
            $wbPreview.NavigateToString($txtMessage.Text)
        }
        Update-Status "Template loaded: $selected"
    } else {
        Show-Error "Not Found" "Template file not found: $path"
        Update-Status "Template file not found"
    }
})

# Load Template button (also loads selected template)
$btnLoadTemplate.Add_Click({
    $selected = $cmbTemplate.SelectedItem.Content
    if ($selected -eq "Custom...") {
        Update-Status "Use 'Browse File...' to load a custom template"
        return
    }
    $path = if ($cmbTemplate.SelectedItem.Tag) { $cmbTemplate.SelectedItem.Tag } else { Get-TemplateFilePath $selected }
    if ($path -and (Test-Path $path)) {
        if (![string]::IsNullOrWhiteSpace($txtMessage.Text)) {
            $backupFile = Join-Path $ConfigDir "message.html.bak"
            Save-MessageToFile $backupFile $txtMessage.Text
        }
        $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $path -Raw)
        if ($tcMessageView.SelectedIndex -eq 1) {
            $wbPreview.NavigateToString($txtMessage.Text)
        }
        Update-Status "Template loaded: $selected"
    } else {
        Show-Error "Not Found" "Template file not found: $path"
        Update-Status "Template file not found"
    }
})

# Re-resolve template when any option checkbox changes
$optionReloadHandler = {
    $selected = $cmbTemplate.SelectedItem.Content
    if ($selected -eq "Custom...") { return }
    $path = if ($cmbTemplate.SelectedItem.Tag) { $cmbTemplate.SelectedItem.Tag } else { Get-TemplateFilePath $selected }
    if ($path -and (Test-Path $path)) {
        $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $path -Raw)
        $wbPreview.NavigateToString($txtMessage.Text)
    }
}
$chkIncludeSignature.Add_Checked($optionReloadHandler)
$chkIncludeSignature.Add_Unchecked($optionReloadHandler)
$chkIncludeOfficeHours.Add_Checked($optionReloadHandler)
$chkIncludeOfficeHours.Add_Unchecked($optionReloadHandler)
$chkIncludeWorkDays.Add_Checked($optionReloadHandler)
$chkIncludeWorkDays.Add_Unchecked($optionReloadHandler)
$chkIncludeTimezone.Add_Checked($optionReloadHandler)
$chkIncludeTimezone.Add_Unchecked($optionReloadHandler)

# Browse File
$btnBrowseFile.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*"
    $dialog.InitialDirectory = $ConfigDir
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtMessage.Text = Get-Content $dialog.FileName -Raw
        Update-Status "Loaded message from $($dialog.FileName)"
    }
})

# Apply Internal Message
$btnApplyInternal.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($txtMessage.Text)) {
            Show-Error "Empty Message" "Please enter or load a message first."
            return
        }
        Update-Status "Applying internal message..."
        Set-ARCMessage $txtMessage.Text 'Internal'
        Update-Status "Internal message applied"
        Show-Result "Done" "Internal auto-reply message updated."
    }
    catch { Show-Error "Error" $_.Exception.Message }
})

# Apply External Message
$btnApplyExternal.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($txtMessage.Text)) {
            Show-Error "Empty Message" "Please enter or load a message first."
            return
        }
        Update-Status "Applying external message..."
        Set-ARCMessage $txtMessage.Text 'External'
        Update-Status "External message applied"
        Show-Result "Done" "External auto-reply message updated."
    }
    catch { Show-Error "Error" $_.Exception.Message }
})

# Apply Both Messages
$btnApplyBoth.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($txtMessage.Text)) {
            Show-Error "Empty Message" "Please enter or load a message first."
            return
        }
        Update-Status "Applying message to both internal and external..."
        Set-ARCMessage $txtMessage.Text 'Both'
        Update-Status "Both messages applied"
        Show-Result "Done" "Internal and External auto-reply messages updated."
    }
    catch { Show-Error "Error" $_.Exception.Message }
})

# Save Template
$btnSaveTemplate.Add_Click({
    $selected = $cmbTemplate.SelectedItem.Content
    if ($selected -eq "Custom...") {
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "HTML Files (*.html)|*.html"
        $dialog.InitialDirectory = $ConfigDir
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Save-MessageToFile $dialog.FileName $txtMessage.Text
            Update-Status "Message saved to $($dialog.FileName)"
            Show-Result "Saved" "Message saved to $($dialog.FileName)"
        }
    } else {
        $path = Get-TemplateFilePath $selected
        if ($path) {
            Save-MessageToFile $path $txtMessage.Text
            Update-Status "Template saved: $selected"
            Show-Result "Saved" "Template '$selected' updated."
        }
    }
})

# Save Online Message to File
$btnSaveOnlineMsg.Add_Click({
    try {
        Update-Status "Fetching current online message..."
        $arc = Get-ARC
        $msgFile = Join-Path $ConfigDir "message.html"
        Save-MessageToFile $msgFile $arc.ExternalMessage
        $txtMessage.Text = $arc.ExternalMessage
        Update-Status "Online message saved to message.html"
        Show-Result "Saved" "Current online auto-reply message saved to:`n$msgFile"
    }
    catch {
        Show-Error "Error" "Could not fetch message. Are you connected?`n$($_.Exception.Message)"
        Update-Status "Failed to save online message"
    }
})

# ===================== HTML Preview Tab Handler =====================
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
Import-UIFromConfig

# Load default template into message box
$defaultTemplate = Get-TemplateFilePath "Normal OOF"
if (Test-Path $defaultTemplate) {
    $txtMessage.Text = Resolve-TemplatePlaceholders (Get-Content $defaultTemplate -Raw)
}

# ===================== Show the Window =====================
$Window.Add_Closed({
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
})
$Window.ShowDialog() | Out-Null

# Cleanup: disconnect on window close
try { Set-EXODisconnect } catch { }
