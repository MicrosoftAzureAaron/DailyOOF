<#
.SYNOPSIS
    Capture README screenshots for each main tab in the Daily OOF GUI.

.DESCRIPTION
    Maintainer-focused helper that launches the XAML window, switches through tabs,
    captures window screenshots, and saves PNG files into the screenshots folder.

.PARAMETER OutputDirectory
    Directory where screenshots are written.

.PARAMETER KeepOpen
    Keep the window open after captures for visual verification.

.PARAMETER CaptureDelayMs
    Delay in milliseconds used between tab switches and captures.

.PARAMETER UseBlankProfile
    Ignore config.json and use generic sample values for all fields.
#>
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot "screenshots"),
    [switch]$KeepOpen,
    [int]$CaptureDelayMs = 650,
    [switch]$UseBlankProfile
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Drawing

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigDir = Join-Path $ScriptDir "config"
$XamlFile = Join-Path $ConfigDir "AAOOF-GUI.xaml"
$ConfigFile = Join-Path $ConfigDir "config.json"
$TemplateFile = Join-Path $ConfigDir "normal_oof.html"

if (-not (Test-Path $XamlFile)) {
    throw "XAML file not found: $XamlFile"
}

function Get-ConfigValue {
    param(
        [object]$Config,
        [string]$Name,
        [object]$DefaultValue = ""
    )

    if ($null -eq $Config) { return $DefaultValue }
    if ($Config.PSObject.Properties.Name -contains $Name) {
        $value = $Config.$Name
        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
            return $value
        }
    }

    return $DefaultValue
}

function Get-OfficeHoursText {
    param([object]$Config)

    $startText = "9:00 AM"
    $endText = "5:00 PM"

    try {
        $startRaw = Get-ConfigValue -Config $Config -Name "StartOfShift" -DefaultValue $null
        if ($startRaw) { $startText = ([datetime]$startRaw).ToString("h:mm tt") }
    }
    catch { }

    try {
        $endRaw = Get-ConfigValue -Config $Config -Name "EndOfShift" -DefaultValue $null
        if ($endRaw) { $endText = ([datetime]$endRaw).ToString("h:mm tt") }
    }
    catch { }

    return "$startText - $endText"
}

function Resolve-TemplatePlaceholdersForCapture {
    param(
        [string]$TemplateContent,
        [object]$Config
    )

    $fullName = [string](Get-ConfigValue -Config $Config -Name "FullName" -DefaultValue $env:USERNAME)
    $nameParts = $fullName -split '\\s+'
    $firstName = if ($nameParts.Count -gt 0) { $nameParts[0] } else { $fullName }
    $lastName = if ($nameParts.Count -gt 1) { $nameParts[$nameParts.Count - 1] } else { "" }

    $workDays = Get-ConfigValue -Config $Config -Name "WorkDays" -DefaultValue @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    if ($workDays -isnot [System.Array]) {
        $workDays = @([string]$workDays)
    }

    $userAlias = [string](Get-ConfigValue -Config $Config -Name "UserAlias" -DefaultValue "")
    if ([string]::IsNullOrWhiteSpace($userAlias)) {
        $suffix = [string](Get-ConfigValue -Config $Config -Name "UserAliasSuffix" -DefaultValue "@microsoft.com")
        $userAlias = "$env:USERNAME$suffix"
    }

    $backupEngineer = [string](Get-ConfigValue -Config $Config -Name "BackupEngineer" -DefaultValue "Backup Engineer")
    if ([string]::IsNullOrWhiteSpace($backupEngineer)) {
        $backupEngineer = [string](Get-ConfigValue -Config $Config -Name "BackupContact" -DefaultValue "Backup Engineer")
    }

    $officeHoursText = Get-OfficeHoursText -Config $Config
    $timezoneText = [System.TimeZoneInfo]::Local.DisplayName
    $workDaysText = ($workDays -join ", ")

    $signatureLines = @(
        "<p><b>Best Regards,</b><br/>",
        "$fullName</p>",
        "<p style='color: #555; font-size: 10pt;'>$officeHoursText | $timezoneText | $workDaysText</p>",
        "<p><a href='mailto:$userAlias'>$userAlias</a></p>"
    )
    $signatureHtml = $signatureLines -join "`n"

    $tokenMap = [ordered]@{
        "[FULL NAME]" = $fullName
        "[FIRST NAME]" = $firstName
        "[LAST NAME]" = $lastName
        "[ROLE]" = [string](Get-ConfigValue -Config $Config -Name "Role" -DefaultValue "member of my team")
        "[EMAIL]" = $userAlias
        "[OFFICE HOURS]" = $officeHoursText
        "[WORK DAYS]" = $workDaysText
        "[TIMEZONE]" = $timezoneText
        "[RETURN DATE]" = (Get-Date).AddDays(5).ToString("MMMM d, yyyy")
        "[BACKUP ENGINEER]" = $backupEngineer
        "[BACKUP CONTACT]" = $backupEngineer
        "[BACKUP ENGINEER EMAIL]" = [string](Get-ConfigValue -Config $Config -Name "BackupEngineerEmail" -DefaultValue "backup@example.com")
        "[TEAM ALIAS]" = [string](Get-ConfigValue -Config $Config -Name "TeamAlias" -DefaultValue "Azure Support Team")
        "[SUPPORT LINK]" = [string](Get-ConfigValue -Config $Config -Name "SupportLink" -DefaultValue "https://portal.azure.com")
        "[SIGNATURE]" = $signatureHtml
    }

    $resolved = $TemplateContent
    foreach ($token in $tokenMap.Keys) {
        $resolved = $resolved -replace [regex]::Escape($token), [string]$tokenMap[$token]
    }

    return $resolved
}

function Set-ComboBoxByContent {
    param(
        [object]$ComboBox,
        [string]$DesiredText
    )

    if ($null -eq $ComboBox -or [string]::IsNullOrWhiteSpace($DesiredText)) { return }

    foreach ($item in $ComboBox.Items) {
        if ([string]$item -eq $DesiredText) {
            $ComboBox.SelectedItem = $item
            return
        }
    }
}

function Initialize-UiFieldsForCapture {
    param([object]$Config)

    $sampleUserAlias = [string](Get-ConfigValue -Config $Config -Name "UserAlias" -DefaultValue "")
    if ([string]::IsNullOrWhiteSpace($sampleUserAlias)) {
        $sampleSuffix = [string](Get-ConfigValue -Config $Config -Name "UserAliasSuffix" -DefaultValue "@microsoft.com")
        $sampleUserAlias = "$env:USERNAME$sampleSuffix"
    }

    $backupEngineer = [string](Get-ConfigValue -Config $Config -Name "BackupEngineer" -DefaultValue "")
    if ([string]::IsNullOrWhiteSpace($backupEngineer)) {
        $backupEngineer = [string](Get-ConfigValue -Config $Config -Name "BackupContact" -DefaultValue "Backup Engineer")
    }

    if ($null -ne $txtAccount) { $txtAccount.Text = $sampleUserAlias }
    if ($null -ne $txtConnectionStatus) { $txtConnectionStatus.Text = "Disconnected" }
    if ($null -ne $txtARCState) { $txtARCState.Text = "Scheduled" }
    if ($null -ne $txtARCStart) { $txtARCStart.Text = (Get-Date).AddHours(-1).ToString("g") }
    if ($null -ne $txtARCEnd) { $txtARCEnd.Text = (Get-Date).AddHours(7).ToString("g") }

    if ($null -ne $txtFullName) { $txtFullName.Text = [string](Get-ConfigValue -Config $Config -Name "FullName" -DefaultValue "Alex Example") }
    if ($null -ne $txtRole) { $txtRole.Text = [string](Get-ConfigValue -Config $Config -Name "Role" -DefaultValue "Support Engineer") }
    if ($null -ne $txtBackupContact) { $txtBackupContact.Text = $backupEngineer }
    if ($null -ne $txtBackupEngineerEmail) { $txtBackupEngineerEmail.Text = [string](Get-ConfigValue -Config $Config -Name "BackupEngineerEmail" -DefaultValue "backup@example.com") }
    if ($null -ne $txtTeamAlias) { $txtTeamAlias.Text = [string](Get-ConfigValue -Config $Config -Name "TeamAlias" -DefaultValue "Azure Support Team") }
    if ($null -ne $txtSupportLink) { $txtSupportLink.Text = [string](Get-ConfigValue -Config $Config -Name "SupportLink" -DefaultValue "https://portal.azure.com") }

    # Use stable demo shift values for screenshot consistency across environments.
    Set-ComboBoxByContent -ComboBox $cmbStartHour -DesiredText "9"
    Set-ComboBoxByContent -ComboBox $cmbStartMin -DesiredText "00"
    Set-ComboBoxByContent -ComboBox $cmbStartAmPm -DesiredText "AM"
    Set-ComboBoxByContent -ComboBox $cmbEndHour -DesiredText "5"
    Set-ComboBoxByContent -ComboBox $cmbEndMin -DesiredText "00"
    Set-ComboBoxByContent -ComboBox $cmbEndAmPm -DesiredText "PM"

    $workDays = Get-ConfigValue -Config $Config -Name "WorkDays" -DefaultValue @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
    if ($workDays -isnot [System.Array]) {
        $workDays = @([string]$workDays)
    }

    if ($null -ne $chkSun) { $chkSun.IsChecked = ($workDays -contains "Sunday") }
    if ($null -ne $chkMon) { $chkMon.IsChecked = ($workDays -contains "Monday") }
    if ($null -ne $chkTue) { $chkTue.IsChecked = ($workDays -contains "Tuesday") }
    if ($null -ne $chkWed) { $chkWed.IsChecked = ($workDays -contains "Wednesday") }
    if ($null -ne $chkThu) { $chkThu.IsChecked = ($workDays -contains "Thursday") }
    if ($null -ne $chkFri) { $chkFri.IsChecked = ($workDays -contains "Friday") }
    if ($null -ne $chkSat) { $chkSat.IsChecked = ($workDays -contains "Saturday") }

    if ($null -ne $txtTaskOffsetMinutes) {
        $txtTaskOffsetMinutes.Text = [string](Get-ConfigValue -Config $Config -Name "TaskStartOffsetMinutes" -DefaultValue 15)
    }
    if ($null -ne $txtTaskExists) { $txtTaskExists.Text = "Created" }
    if ($null -ne $txtTaskState) { $txtTaskState.Text = "Ready" }
    if ($null -ne $txtTaskNextRun) { $txtTaskNextRun.Text = (Get-Date).AddMinutes(45).ToString("g") }
    if ($null -ne $txtTaskLastRun) { $txtTaskLastRun.Text = (Get-Date).AddHours(-18).ToString("g") }
    if ($null -ne $txtTaskLastResult) { $txtTaskLastResult.Text = "Success (0x0)" }
    if ($null -ne $txtTaskScriptPath) { $txtTaskScriptPath.Text = (Join-Path $ScriptDir "AAOOF-GUI.ps1") }
    if ($null -ne $txtTaskSummary) { $txtTaskSummary.Text = "Task is ready for daily automation." }
    if ($null -ne $txtLocalVersion) { $txtLocalVersion.Text = "1.9.26" }
    if ($null -ne $txtRemoteVersion) { $txtRemoteVersion.Text = "1.9.26" }
}

function Invoke-UiPump {
    param([int]$Milliseconds = 250)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $Milliseconds) {
        $Window.Dispatcher.Invoke([action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
        Start-Sleep -Milliseconds 30
    }
}

function Wait-WebBrowserReady {
    param(
        [object]$Browser,
        [int]$TimeoutMs = 5000
    )

    $script:BrowserLoaded = $false
    $handler = [System.Windows.Navigation.LoadCompletedEventHandler] { $script:BrowserLoaded = $true }
    $Browser.Add_LoadCompleted($handler)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not $script:BrowserLoaded -and $sw.ElapsedMilliseconds -lt $TimeoutMs) {
        Invoke-UiPump -Milliseconds 100
    }

    $Browser.Remove_LoadCompleted($handler)
    $Window.Dispatcher.Invoke([action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
    Start-Sleep -Milliseconds 250
}

function Save-WindowCapture {
    param([string]$Path)

    $Window.Dispatcher.Invoke([action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
    Start-Sleep -Milliseconds 200

    $source = [System.Windows.PresentationSource]::FromVisual($Window)
    if ($null -eq $source) {
        throw "Could not resolve window presentation source for capture."
    }

    [double]$dpiX = $source.CompositionTarget.TransformToDevice.M11
    [double]$dpiY = $source.CompositionTarget.TransformToDevice.M22

    [int]$left = [Math]::Round($Window.Left * $dpiX)
    [int]$top = [Math]::Round($Window.Top * $dpiY)
    [int]$width = [Math]::Round($Window.ActualWidth * $dpiX)
    [int]$height = [Math]::Round($Window.ActualHeight * $dpiY)

    if ($width -le 0 -or $height -le 0) {
        throw "Window size is invalid for screenshot capture."
    }

    $bitmap = New-Object System.Drawing.Bitmap($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($left, $top, 0, 0, (New-Object System.Drawing.Size($width, $height)))
    $graphics.Dispose()

    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
}

[xml]$xaml = Get-Content $XamlFile -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

$tcMain = $Window.FindName("tcMain")
$tcMessageView = $Window.FindName("tcMessageView")
$txtMessage = $Window.FindName("txtMessage")
$wbPreview = $Window.FindName("wbPreview")
$wbCurrentOOF = $Window.FindName("wbCurrentOOF")
$txtCurrentOOFStatus = $Window.FindName("txtCurrentOOFStatus")
$txtStatusBar = $Window.FindName("txtStatusBar")

if ($null -eq $tcMain -or $null -eq $tcMessageView -or $null -eq $txtMessage -or $null -eq $wbPreview -or $null -eq $wbCurrentOOF) {
    throw "One or more required controls were not found in the XAML."
}

$txtAccount = $Window.FindName("txtAccount")
$txtConnectionStatus = $Window.FindName("txtConnectionStatus")
$txtARCState = $Window.FindName("txtARCState")
$txtARCStart = $Window.FindName("txtARCStart")
$txtARCEnd = $Window.FindName("txtARCEnd")
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
$chkSun = $Window.FindName("chkSun")
$chkMon = $Window.FindName("chkMon")
$chkTue = $Window.FindName("chkTue")
$chkWed = $Window.FindName("chkWed")
$chkThu = $Window.FindName("chkThu")
$chkFri = $Window.FindName("chkFri")
$chkSat = $Window.FindName("chkSat")
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

$config = $null
if (-not $UseBlankProfile -and (Test-Path $ConfigFile)) {
    try {
        $config = Get-Content $ConfigFile -Raw | ConvertFrom-Json
    }
    catch {
        Write-Warning "Could not parse config.json. Using fallback placeholder values."
    }
}

$templateRaw = if (Test-Path $TemplateFile) {
    Get-Content $TemplateFile -Raw
}
else {
    "<html><body><h3>Out of Office</h3><p>Hello, I am currently away and will respond as soon as possible.</p></body></html>"
}

if ($UseBlankProfile) {
    Write-Host "UseBlankProfile enabled: using generic sample values instead of config.json" -ForegroundColor Yellow
}

Initialize-UiFieldsForCapture -Config $config

$txtMessage.Text = $templateRaw
$resolvedTemplate = Resolve-TemplatePlaceholdersForCapture -TemplateContent $templateRaw -Config $config
$wbPreview.NavigateToString($resolvedTemplate)
$wbCurrentOOF.NavigateToString($resolvedTemplate)
$txtCurrentOOFStatus.Text = "Captured in maintainer screenshot mode"
$txtStatusBar.Text = "Preparing screenshot capture..."

if (-not (Test-Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory | Out-Null
}

$script:CaptureError = $null
$script:CapturedFiles = @()

function Invoke-CaptureRun {
    $originalTab = $tcMain.SelectedIndex
    $originalSubTab = $tcMessageView.SelectedIndex

    try {
        $tabShots = @(
            @{ Index = 0; File = "quick-actions.png" },
            @{ Index = 1; File = "configuration.png" },
            @{ Index = 2; File = "automation.png" }
        )

        foreach ($shot in $tabShots) {
            $tcMain.SelectedIndex = $shot.Index
            Invoke-UiPump -Milliseconds $CaptureDelayMs
            $path = Join-Path $OutputDirectory $shot.File
            Save-WindowCapture -Path $path
            $script:CapturedFiles += $path
        }

        $tcMain.SelectedIndex = 3
        $tcMessageView.SelectedIndex = 0
        Invoke-UiPump -Milliseconds $CaptureDelayMs
        $editPath = Join-Path $OutputDirectory "message-templates-edit.png"
        Save-WindowCapture -Path $editPath
        $script:CapturedFiles += $editPath

        $tcMessageView.SelectedIndex = 1
        Wait-WebBrowserReady -Browser $wbPreview
        $previewPath = Join-Path $OutputDirectory "message-templates-preview.png"
        Save-WindowCapture -Path $previewPath
        $script:CapturedFiles += $previewPath

        $tcMain.SelectedIndex = 4
        Wait-WebBrowserReady -Browser $wbCurrentOOF
        $currentPath = Join-Path $OutputDirectory "current-oof.png"
        Save-WindowCapture -Path $currentPath
        $script:CapturedFiles += $currentPath

        $txtStatusBar.Text = "Screenshot capture complete"
        Write-Host "Saved screenshots:" -ForegroundColor Green
        foreach ($file in $script:CapturedFiles) {
            Write-Host " - $file"
        }
    }
    finally {
        $tcMessageView.SelectedIndex = $originalSubTab
        $tcMain.SelectedIndex = $originalTab
    }
}

$Window.Add_ContentRendered({
    try {
        Invoke-CaptureRun
    }
    catch {
        $script:CaptureError = $_.Exception.Message
        Write-Error "Screenshot capture failed: $($script:CaptureError)"
    }
    finally {
        if (-not $KeepOpen) {
            $Window.Close()
        }
    }
})

$Window.ShowDialog() | Out-Null

if ($script:CaptureError) {
    exit 1
}

exit 0
