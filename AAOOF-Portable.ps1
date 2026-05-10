<#!
.SYNOPSIS
    Daily OOF Portable - Simplified GUI for auto-reply state management.
.DESCRIPTION
    Standalone single-file GUI for managing only Exchange Online auto-reply state.
    This portable mode intentionally excludes template editing, message apply flows,
    and scheduled task automation features.

    Manage OOF message content in Outlook or Outlook on the web.
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$script:PortableVersion = "1.1.3"
$script:IsConnectedToEXO = $false
$script:UserAlias = ""
$script:UserAliasSuffix = ""
$script:StartOfShift = $null
$script:EndOfShift = $null
$script:WorkDays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
$script:AutoReplyAudience = 'Both'

function Show-ErrorDialog($Title, $Message) {
    [System.Windows.MessageBox]::Show($Message, $Title, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
}

function Show-InfoDialog($Title, $Message) {
    [System.Windows.MessageBox]::Show($Message, $Title, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
}

function Update-StatusBar($Message) {
    if ($null -ne $txtStatusBar) {
        $txtStatusBar.Text = $Message
    }
}

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

function Get-InstallCommand {
    if (Test-IsAdmin) {
        return "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber"
    }
    return "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser"
}

function Resolve-UserAlias {
    if ([string]::IsNullOrEmpty($script:UserAliasSuffix)) {
        if ($env:USERDNSDOMAIN) {
            $dnsDomain = $env:USERDNSDOMAIN.ToLower()
            if ($dnsDomain -match '\.?microsoft\.com$') {
                $script:UserAliasSuffix = '@microsoft.com'
            }
            else {
                $script:UserAliasSuffix = "@$dnsDomain"
            }
        }
        else {
            $script:UserAliasSuffix = '@microsoft.com'
        }
    }

    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
    if ($computerSystem -and $computerSystem.Username) {
        $currentUser = $computerSystem.Username.Split('\\')[-1]
    }
    else {
        $currentUser = $env:USERNAME
    }

    $script:UserAlias = "$currentUser$script:UserAliasSuffix"
    return $script:UserAlias
}

function Install-ExchangeModuleIfMissing {
    $moduleInstalled = Get-Module -ListAvailable -Name ExchangeOnlineManagement
    if ($moduleInstalled) { return }

    $installCommand = Get-InstallCommand

    $install = [System.Windows.MessageBox]::Show(
        "ExchangeOnlineManagement module is required but not installed.`n`nInstall now?`n`nRecommended command:`n$installCommand",
        "Module Required",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )

    if ($install -ne [System.Windows.MessageBoxResult]::Yes) {
        throw "ExchangeOnlineManagement module is required to continue."
    }

    if (Test-IsAdmin) {
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -ErrorAction Stop
    }
    else {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
}

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
        $win = [System.Windows.Markup.XamlReader]::Load($reader)
        $lbl = $win.FindName('lblElapsed')
        $btn = $win.FindName('btnCancel')
        $syncHash.Window = $win

        $btn.Add_Click({
            $syncHash.Cancelled = $true
            $syncHash.Done = $true
            $win.Close()
        })

        $startTime = [datetime]::Now
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(500)
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
    $waited = 0
    while ($null -eq $syncHash.Window -and $waited -lt 3000) {
        Start-Sleep -Milliseconds 50
        $waited += 50
    }

    return @{ SyncHash = $syncHash; PS = $ps; Runspace = $runspace; Handle = $handle }
}

function Close-ConnectingWindow {
    param($ctx)
    if ($null -eq $ctx) { return }
    $ctx.SyncHash.Done = $true
    try { $null = $ctx.PS.EndInvoke($ctx.Handle) } catch {}
    try { $ctx.PS.Dispose() } catch {}
    try { $ctx.Runspace.Close(); $ctx.Runspace.Dispose() } catch {}
}

function Connect-ExchangeOnlineSession {
    Install-ExchangeModuleIfMissing

    if (!(Get-Module -Name ExchangeOnlineManagement)) {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
    }

    if (-not (Test-ValidEmailAddress $script:UserAlias)) {
        throw "Please enter a valid mailbox email address before connecting."
    }

    $session = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -ne $session) {
        $exchangeSession = $session | Where-Object { $_.Name -like "ExchangeOnline_*" }
        if ($null -ne $exchangeSession) {
            $script:IsConnectedToEXO = $true
            $txtConnectionStatus.Text = "Connected"
            $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Green
            Update-StatusBar "Connected"
            return
        }
    }

    Update-StatusBar "Connecting to Exchange Online..."
    Connect-ExchangeOnline -UserPrincipalName $script:UserAlias -ShowBanner:$false -CommandName Get-MailboxAutoReplyConfiguration, Set-MailboxAutoReplyConfiguration -ErrorAction Stop
    $script:IsConnectedToEXO = $true
    $txtConnectionStatus.Text = "Connected"
    $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Green
    Update-StatusBar "Connected"
}

function Disconnect-ExchangeOnlineSession {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    $script:IsConnectedToEXO = $false
    $txtConnectionStatus.Text = "Disconnected"
    $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::DarkOrange
    Update-StatusBar "Disconnected"
}

function Assert-ExchangeConnection {
    if (-not $script:IsConnectedToEXO) {
        throw "Connect to Exchange Online first."
    }
}

function Resolve-AutoReplyAudience($Audience) {
    if ([string]::IsNullOrWhiteSpace($Audience)) { return 'Both' }

    switch ($Audience.ToString().Trim()) {
        'Both' { return 'Both' }
        'InternalOnly' { return 'InternalOnly' }
        'ExternalOnly' { return 'ExternalOnly' }
        default { return 'Both' }
    }
}

function Get-AutoReplyAudienceFromUI {
    if ($null -eq $cmbAudience -or $null -eq $cmbAudience.SelectedItem) {
        return 'Both'
    }

    $selectedItem = $cmbAudience.SelectedItem
    if ($null -ne $selectedItem.Tag) {
        return Resolve-AutoReplyAudience($selectedItem.Tag.ToString())
    }

    return 'Both'
}

function Get-ExternalAudienceFromSelection($Audience) {
    if ($Audience -eq 'InternalOnly') {
        return 'None'
    }
    return 'All'
}

function ConvertTo-TimeOfDayFromInput($TimeText, $FieldLabel) {
    $parsedTime = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($TimeText.Text, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedTime)) {
        throw "$FieldLabel time must be in HH:mm format (24-hour)."
    }

    return $parsedTime
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

function Set-WorkDayPreset([string[]]$Days) {
    $chkSun.IsChecked = ('Sunday' -in $Days)
    $chkMon.IsChecked = ('Monday' -in $Days)
    $chkTue.IsChecked = ('Tuesday' -in $Days)
    $chkWed.IsChecked = ('Wednesday' -in $Days)
    $chkThu.IsChecked = ('Thursday' -in $Days)
    $chkFri.IsChecked = ('Friday' -in $Days)
    $chkSat.IsChecked = ('Saturday' -in $Days)
}

function Get-NextWorkDayOffset {
    if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift) { return 1 }
    if (-not $script:WorkDays -or $script:WorkDays.Count -eq 0) { return 1 }

    $currentTime = [datetime](Get-Date)

    if (!($currentTime.DayOfWeek -in $script:WorkDays)) {
        $daysAhead = 0
        while (!($currentTime.DayOfWeek -in $script:WorkDays)) {
            $daysAhead += 1
            $currentTime = $currentTime.AddDays(1)
        }
        return $daysAhead
    }

    $nextDay = $currentTime.AddDays(1)
    $daysAhead = 1
    while (!($nextDay.DayOfWeek -in $script:WorkDays)) {
        $daysAhead += 1
        $nextDay = $nextDay.AddDays(1)
    }
    if ($daysAhead -gt 1) {
        return $daysAhead
    }

    if ($currentTime.TimeOfDay -lt $script:StartOfShift.TimeOfDay) {
        return 0
    }

    return 1
}

function Get-AutoReplyScheduleTimes {
    if ($null -eq $script:StartOfShift -or $null -eq $script:EndOfShift) {
        throw "Shift start and end times are required."
    }
    if (-not $script:WorkDays -or $script:WorkDays.Count -eq 0) {
        throw "Select at least one work day."
    }

    $daysToAdd = Get-NextWorkDayOffset
    $oofEndTime = (Get-Date).Date.Add($script:StartOfShift.TimeOfDay).AddDays($daysToAdd)
    $oofStartTime = (Get-Date).Date.Add($script:EndOfShift.TimeOfDay)
    if ($daysToAdd -eq 0) {
        $oofStartTime = $oofStartTime.AddDays(-1)
    }

    if ($oofStartTime -ge $oofEndTime) {
        throw "Calculated schedule is invalid. Check your shift times and work days."
    }

    return @{ StartTime = $oofStartTime; EndTime = $oofEndTime }
}

function Update-AutoReplyStatus {
    Assert-ExchangeConnection
    $arc = Get-MailboxAutoReplyConfiguration -Identity $script:UserAlias -ErrorAction Stop

    $txtARCState.Text = [string]$arc.AutoReplyState
    $txtARCStart.Text = if ($arc.StartTime) { $arc.StartTime.ToString("g") } else { "-" }
    $txtARCEnd.Text = if ($arc.EndTime) { $arc.EndTime.ToString("g") } else { "-" }

    Update-StatusBar "Auto-reply status refreshed"
}

function Set-AutoReplyEnabled {
    Assert-ExchangeConnection
    $script:AutoReplyAudience = Get-AutoReplyAudienceFromUI
    $externalAudience = Get-ExternalAudienceFromSelection $script:AutoReplyAudience

    if ($script:AutoReplyAudience -eq 'ExternalOnly') {
        Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState Enabled -ExternalAudience $externalAudience -InternalMessage '' -ErrorAction Stop
    }
    else {
        Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState Enabled -ExternalAudience $externalAudience -ErrorAction Stop
    }

    Update-AutoReplyStatus
    Show-InfoDialog "Done" "Auto-reply state set to Enabled ($($script:AutoReplyAudience))."
}

function Set-AutoReplyDisabled {
    Assert-ExchangeConnection
    Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState Disabled -ExternalAudience None -ErrorAction Stop
    Update-AutoReplyStatus
    Show-InfoDialog "Done" "Auto-reply state set to Disabled."
}

function Set-AutoReplyScheduled {
    Assert-ExchangeConnection
    $script:AutoReplyAudience = Get-AutoReplyAudienceFromUI

    $script:StartOfShift = ConvertTo-TimeOfDayFromInput -TimeText $txtStartTime -FieldLabel "Start"
    $script:EndOfShift = ConvertTo-TimeOfDayFromInput -TimeText $txtEndTime -FieldLabel "End"
    $script:WorkDays = Read-WorkDaysFromUI

    if ($script:EndOfShift.TimeOfDay -eq $script:StartOfShift.TimeOfDay) {
        throw "End time must be different from start time."
    }

    $schedule = Get-AutoReplyScheduleTimes
    $externalAudience = Get-ExternalAudienceFromSelection $script:AutoReplyAudience

    if ($script:AutoReplyAudience -eq 'ExternalOnly') {
        Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState Scheduled -StartTime $schedule.StartTime -EndTime $schedule.EndTime -ExternalAudience $externalAudience -InternalMessage '' -ErrorAction Stop
    }
    else {
        Set-MailboxAutoReplyConfiguration -Identity $script:UserAlias -AutoReplyState Scheduled -StartTime $schedule.StartTime -EndTime $schedule.EndTime -ExternalAudience $externalAudience -ErrorAction Stop
    }

    Update-AutoReplyStatus
    Show-InfoDialog "Done" "Auto-reply state set to Scheduled ($($script:AutoReplyAudience)).`nStart: $($schedule.StartTime)`nEnd: $($schedule.EndTime)"
}

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Daily OOF Portable - State Manager"
    Height="820" Width="760"
    MinHeight="820" MinWidth="720"
        WindowStartupLocation="CenterScreen"
        Background="#F4F6F8">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" Padding="14">
            <StackPanel>
                <Border Background="White" BorderBrush="#D8D8D8" BorderThickness="1" CornerRadius="6" Padding="12" Margin="0,0,0,10">
                    <StackPanel>
                        <TextBlock Text="Daily OOF Portable (v$($script:PortableVersion))" FontSize="18" FontWeight="Bold" Foreground="#1F2937"/>
                        <TextBlock Text="Simplified standalone GUI for auto-reply state only." Foreground="#4B5563" Margin="0,4,0,0"/>
                        <TextBlock Text="Manage message content in Outlook or Outlook on the web." Foreground="#B45309" FontWeight="SemiBold" Margin="0,8,0,0"/>
                    </StackPanel>
                </Border>

                <GroupBox Header="Connection" Margin="0,0,0,10">
                    <StackPanel Margin="10">
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                            <TextBlock Text="Mailbox:" Width="90" VerticalAlignment="Center"/>
                            <TextBox x:Name="txtAccount" Width="320" VerticalContentAlignment="Center" Padding="4"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                            <TextBlock Text="Status:" Width="90" VerticalAlignment="Center"/>
                            <TextBlock x:Name="txtConnectionStatus" Text="Disconnected" Foreground="DarkOrange" FontWeight="SemiBold" VerticalAlignment="Center"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal">
                            <Button x:Name="btnConnect" Content="Connect" Padding="14,6" Margin="0,0,8,0"/>
                            <Button x:Name="btnDisconnect" Content="Disconnect" Padding="14,6"/>
                        </StackPanel>
                    </StackPanel>
                </GroupBox>

                <GroupBox Header="Auto-Reply State" Margin="0,0,0,10">
                    <StackPanel Margin="10">
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                            <TextBlock Text="Audience:" Width="110" VerticalAlignment="Center"/>
                            <ComboBox x:Name="cmbAudience" Width="260" SelectedIndex="0">
                                <ComboBoxItem Content="Internal and External (default)" Tag="Both"/>
                                <ComboBoxItem Content="Internal Only" Tag="InternalOnly"/>
                                <ComboBoxItem Content="External Only" Tag="ExternalOnly"/>
                            </ComboBox>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                            <TextBlock Text="State:" Width="110"/>
                            <TextBlock x:Name="txtARCState" Text="Unknown" FontWeight="Bold"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                            <TextBlock Text="Start:" Width="110"/>
                            <TextBlock x:Name="txtARCStart" Text="-"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <TextBlock Text="End:" Width="110"/>
                            <TextBlock x:Name="txtARCEnd" Text="-"/>
                        </StackPanel>

                        <StackPanel Orientation="Horizontal">
                            <Button x:Name="btnStateEnabled" Content="Set Enabled" Padding="12,6" Margin="0,0,8,0"/>
                            <Button x:Name="btnStateDisabled" Content="Set Disabled" Padding="12,6" Margin="0,0,8,0"/>
                            <Button x:Name="btnRefreshStatus" Content="Refresh Status" Padding="12,6"/>
                        </StackPanel>
                    </StackPanel>
                </GroupBox>

                <GroupBox Header="Work Schedule" Margin="0,0,0,10">
                    <StackPanel Margin="10">
                        <TextBlock Text="Set your normal shift start/end times and work days. The tool will calculate OOF from end of shift to the next work-day start." Foreground="#4B5563" Margin="0,0,0,10" TextWrapping="Wrap"/>

                        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                            <TextBlock Text="Start Time:" Width="90" VerticalAlignment="Center"/>
                            <TextBox x:Name="txtStartTime" Width="90" Text="09:00"/>
                            <TextBlock Text="End Time:" Width="80" VerticalAlignment="Center" Margin="16,0,0,0"/>
                            <TextBox x:Name="txtEndTime" Width="90" Text="17:00"/>
                        </StackPanel>

                        <TextBlock Text="Work Days:" Margin="0,0,0,6"/>
                        <WrapPanel Margin="0,0,0,8">
                            <CheckBox x:Name="chkSun" Content="Sun" Margin="0,0,12,4"/>
                            <CheckBox x:Name="chkMon" Content="Mon" Margin="0,0,12,4"/>
                            <CheckBox x:Name="chkTue" Content="Tue" Margin="0,0,12,4"/>
                            <CheckBox x:Name="chkWed" Content="Wed" Margin="0,0,12,4"/>
                            <CheckBox x:Name="chkThu" Content="Thu" Margin="0,0,12,4"/>
                            <CheckBox x:Name="chkFri" Content="Fri" Margin="0,0,12,4"/>
                            <CheckBox x:Name="chkSat" Content="Sat" Margin="0,0,12,4"/>
                        </WrapPanel>

                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <Button x:Name="btnPresetMF" Content="Mon-Fri" Padding="10,5" Margin="0,0,8,0"/>
                            <Button x:Name="btnPresetSunWed" Content="Sun-Wed" Padding="10,5" Margin="0,0,8,0"/>
                            <Button x:Name="btnPresetWedSat" Content="Wed-Sat" Padding="10,5"/>
                        </StackPanel>

                        <Button x:Name="btnStateScheduled" Content="Set Scheduled" Padding="12,6" HorizontalAlignment="Left"/>
                    </StackPanel>
                </GroupBox>
            </StackPanel>
        </ScrollViewer>

        <Border Grid.Row="1" Background="#0F62FE" Height="28">
            <TextBlock x:Name="txtStatusBar" Text="Ready" Foreground="White" VerticalAlignment="Center" Margin="12,0,0,0"/>
        </Border>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($xaml)))
$Window = [Windows.Markup.XamlReader]::Load($reader)

$txtAccount = $Window.FindName("txtAccount")
$txtConnectionStatus = $Window.FindName("txtConnectionStatus")
$btnConnect = $Window.FindName("btnConnect")
$btnDisconnect = $Window.FindName("btnDisconnect")
$txtARCState = $Window.FindName("txtARCState")
$txtARCStart = $Window.FindName("txtARCStart")
$txtARCEnd = $Window.FindName("txtARCEnd")
$cmbAudience = $Window.FindName("cmbAudience")
$btnStateEnabled = $Window.FindName("btnStateEnabled")
$btnStateDisabled = $Window.FindName("btnStateDisabled")
$btnStateScheduled = $Window.FindName("btnStateScheduled")
$btnRefreshStatus = $Window.FindName("btnRefreshStatus")
$txtStartTime = $Window.FindName("txtStartTime")
$txtEndTime = $Window.FindName("txtEndTime")
$chkSun = $Window.FindName("chkSun")
$chkMon = $Window.FindName("chkMon")
$chkTue = $Window.FindName("chkTue")
$chkWed = $Window.FindName("chkWed")
$chkThu = $Window.FindName("chkThu")
$chkFri = $Window.FindName("chkFri")
$chkSat = $Window.FindName("chkSat")
$btnPresetMF = $Window.FindName("btnPresetMF")
$btnPresetSunWed = $Window.FindName("btnPresetSunWed")
$btnPresetWedSat = $Window.FindName("btnPresetWedSat")
$txtStatusBar = $Window.FindName("txtStatusBar")

$defaultAlias = Resolve-UserAlias
$txtAccount.Text = $defaultAlias
$script:UserAlias = $defaultAlias
$cmbAudience.SelectedIndex = 0
Set-WorkDayPreset @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')

$cmbAudience.Add_SelectionChanged({
        $script:AutoReplyAudience = Get-AutoReplyAudienceFromUI
        Update-StatusBar "Audience set to $script:AutoReplyAudience"
    })

$btnPresetMF.Add_Click({ Set-WorkDayPreset @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday') })
$btnPresetSunWed.Add_Click({ Set-WorkDayPreset @('Sunday', 'Monday', 'Tuesday', 'Wednesday') })
$btnPresetWedSat.Add_Click({ Set-WorkDayPreset @('Wednesday', 'Thursday', 'Friday', 'Saturday') })

$btnConnect.Add_Click({
        try {
            $script:UserAlias = $txtAccount.Text.Trim()
            $connectCtx = Show-ConnectingWindow
            try {
                Connect-ExchangeOnlineSession
            }
            finally {
                Close-ConnectingWindow $connectCtx
            }
            if ($connectCtx.SyncHash.Cancelled) {
                Disconnect-ExchangeOnlineSession
                Update-StatusBar "Connection cancelled"
                return
            }
            Update-AutoReplyStatus
        }
        catch {
            $script:IsConnectedToEXO = $false
            $txtConnectionStatus.Text = "Connection failed"
            $txtConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Red
            Show-ErrorDialog "Connection Error" $_.Exception.Message
            Update-StatusBar "Connection failed"
        }
    })

$btnDisconnect.Add_Click({
        try {
            Disconnect-ExchangeOnlineSession
        }
        catch {
            Show-ErrorDialog "Disconnect Error" $_.Exception.Message
            Update-StatusBar "Disconnect failed"
        }
    })

$btnRefreshStatus.Add_Click({
        try {
            Update-AutoReplyStatus
        }
        catch {
            Show-ErrorDialog "Refresh Error" $_.Exception.Message
            Update-StatusBar "Refresh failed"
        }
    })

$btnStateEnabled.Add_Click({
        try {
            Set-AutoReplyEnabled
        }
        catch {
            Show-ErrorDialog "Set Enabled Error" $_.Exception.Message
            Update-StatusBar "Set enabled failed"
        }
    })

$btnStateDisabled.Add_Click({
        try {
            Set-AutoReplyDisabled
        }
        catch {
            Show-ErrorDialog "Set Disabled Error" $_.Exception.Message
            Update-StatusBar "Set disabled failed"
        }
    })

$btnStateScheduled.Add_Click({
        try {
            Set-AutoReplyScheduled
        }
        catch {
            Show-ErrorDialog "Set Scheduled Error" $_.Exception.Message
            Update-StatusBar "Set scheduled failed"
        }
    })

$Window.Add_Closed({
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    })

$Window.ShowDialog() | Out-Null
