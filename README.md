# Daily OOF — Out of Office Automation for Exchange Online

A PowerShell WPF GUI application that automates Exchange Online Out of Office (OOF) message management. Built for Azure Support Engineers but usable by anyone with an Exchange Online mailbox.

---

## Screenshots

### Quick Actions
<!-- TODO: Add screenshot of the Quick Actions tab -->
![Quick Actions Tab](screenshots/quick-actions.png)

### Configuration
<!-- TODO: Add screenshot of the Configuration tab -->
![Configuration Tab](screenshots/configuration.png)

### Message Templates — Edit
<!-- TODO: Add screenshot of the Message Templates tab (Edit sub-tab) -->
![Message Templates — Edit](screenshots/message-templates-edit.png)

### Message Templates — Preview
<!-- TODO: Add screenshot of the Message Templates tab (Preview sub-tab) -->
![Message Templates — Preview](screenshots/message-templates-preview.png)

### Current OOF
<!-- TODO: Add screenshot of the Current OOF tab -->
![Current OOF Tab](screenshots/current-oof.png)

---

## Features

### GUI Mode
Launch the app with no arguments to open the full graphical interface:

```powershell
.\AAOOF-GUI.ps1
```

The GUI has four tabs:

---

#### Quick Actions
| Action | Description |
|---|---|
| **Connect / Disconnect** | Authenticate to Exchange Online using your alias. On connect, the app fetches your current OOF status. |
| **Enable Scheduled Auto Reply** | Sets OOF to *Scheduled* mode using your configured shift hours and work days. The start/end times are calculated automatically based on the current day. |
| **Set Vacation OOF** | Pick a return date and the app sets a *Scheduled* OOF that disables automatically when you return. |
| **Cancel Vacation OOF** | Immediately disables the vacation/extended OOF. |
| **Refresh Status** | Shows the current Auto Reply state, start time, and end time from Exchange Online. |
| **View Current Message** | Fetches the current OOF message from Exchange Online (auto-connects if needed) and displays it on the **Current OOF** tab. |

---

#### Configuration
| Setting | Description |
|---|---|
| **Full Name** | Your display name for the auto-generated signature. Changes are saved and applied to templates immediately. |
| **Role** | Your job title, inserted into templates via the `[ROLE]` placeholder. If left blank, templates use "member of my team" instead. |
| **Email Suffix** | Your email domain suffix (default `@microsoft.com`). Combined with your Windows username to form the mailbox identity. |
| **Override Account** | Manually set your full email address instead of using auto-detection. |
| **Office Hours** | Start and end times for your shift (supports non-hour-boundary times like 8:30). Used for OOF scheduling and the signature's office-hours line. |
| **Work Days** | Select your working days (Sunday–Saturday). Preset buttons available: Mon–Fri, Sun–Wed (4×10), Wed–Sat (4×10). Days appear in the signature sorted by your work week (e.g., Wed–Sat starts with Wednesday). |
| **Auto Reply State** | Manually set OOF to Enabled, Disabled, or Scheduled. |
| **Scheduled Task** | Create a Windows Task Scheduler job to run the script daily in CLI mode, 15 minutes after your shift start. **Note:** Run the script as Administrator before clicking this button. |

---

#### Message Templates
Choose from four built-in HTML templates or load your own:

| Template | Description |
|---|---|
| **Normal OOF** | Standard out-of-office message (blue header) |
| **Vacation OOF** | Vacation message with `[RETURN DATE]` placeholder (green header) |
| **Sick OOF** | Unexpected absence / illness message (red header) |
| **Holiday OOF** | Company holiday message with `[RETURN DATE]` placeholder (amber header) |

Templates auto-load when you change the dropdown selection. Your current message is backed up to `message.html.bak` before being replaced.

##### Template Options
A **Template Options** panel lets you toggle which dynamic content is injected into templates when loaded:

| Option | Placeholder | Effect when unchecked |
|---|---|---|
| **Include Signature** | `[SIGNATURE]` | Removes the auto-generated signature block (name, details, email) |
| **Include Office Hours** | — | Removes your office hours from the signature block |
| **Include Work Days** | — | Removes your work days from the signature block |
| **Include Timezone** | — | Removes the timezone from the signature block |

> Toggling any option or changing the return date immediately re-renders the template in the editor and preview.

> **Signature note:** The signature is auto-generated from your Full Name field (or derived from your Windows username if blank). Because the auto-derived name is a best guess, **double-check that your name appears correctly** in the preview before applying, or enter your name in the Full Name field on the Configuration tab.

##### Edit & Preview
- **Edit tab** — View and hand-edit the raw HTML source
- **Preview tab** — Rendered HTML preview using a built-in browser control

##### Apply & Save
| Button | Description |
|---|---|
| **Apply as Internal Message** | Sets the message as the internal OOF reply |
| **Apply as External Message** | Sets the message as the external OOF reply |
| **Apply as Both** | Sets the message for both internal and external |
| **Save to Template File** | Overwrites the selected template file with the current editor content |
| **Save Online Msg to File** | Downloads the current live OOF message from Exchange Online and saves it locally |

---

#### Current OOF
View your live auto-reply message as it appears to senders, rendered in an embedded browser control. The message is **automatically loaded** (with auto-connect if needed) the first time you switch to this tab.

| Action | Description |
|---|---|
| **Refresh** | Re-fetches the current OOF message from Exchange Online and renders it. Auto-reconnects if the session has expired. |
| **Status indicator** | Shows the current auto-reply state (Enabled / Disabled / Scheduled) and last-refreshed timestamp. |

---

### CLI Mode
For automation and scheduled tasks:

```powershell
# Daily scheduled auto-reply (uses saved config)
# Skips if a vacation/extended OOF is active.
.\AAOOF-GUI.ps1 1

# Vacation mode until a specific date
.\AAOOF-GUI.ps1 '2026/04/14'
```

---

## Getting Started

### Prerequisites
- **PowerShell 5.1+** (Windows PowerShell) or **PowerShell 7+**
- **Exchange Online Management** module (prompted to install on first connect if missing)
- An Exchange Online mailbox

### Installation

```powershell
mkdir c:\tools
cd c:\tools
git clone https://github.com/MicrosoftAzureAaron/DailyOOF
cd DailyOOF
.\AAOOF-GUI.ps1
```

On first run the app downloads any missing config files (XAML, templates) from the repository automatically.

### Configuration Files

All configuration is stored in the `config/` folder:

| File | Purpose |
|---|---|
| `config.json` | User settings: shift times, work days, alias, name, role, suffix (gitignored) |
| `AAOOF-GUI.xaml` | WPF UI layout (auto-downloaded if missing) |
| `normal_oof.html` | Normal OOF template |
| `vacation_oof.html` | Vacation OOF template |
| `sick_oof.html` | Sick OOF template |
| `holiday_oof.html` | Holiday OOF template |
| `message.html` | Last-applied message (gitignored) |
| `message.html.bak` | Auto-backup of previous message (gitignored) |
| `AutoReplyConfig.json` | Cached Exchange auto-reply config (gitignored) |

### Custom Templates

You can edit the HTML template files directly or create your own. Supported placeholders:

| Placeholder | Replaced with |
|---|---|
| `[ROLE]` | Your configured role / job title |
| `[RETURN DATE]` | The return date selected in the Vacation date picker |
| `[SIGNATURE]` | Auto-generated signature block: greeting, display name, office hours/timezone/work days, and email link |
