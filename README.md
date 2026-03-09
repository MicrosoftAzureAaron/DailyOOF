# Daily OOF — Out of Office Automation for Exchange Online

A PowerShell WPF GUI application that automates Exchange Online Out of Office (OOF) message management. Built for Azure Support Engineers but usable by anyone with an Exchange Online mailbox.

## Features

### GUI Mode
Launch the app with no arguments to open the full graphical interface:

```powershell
.\AAOOF-GUI.ps1
```

The GUI has three tabs:

---

#### Quick Actions
| Action | Description |
|---|---|
| **Connect / Disconnect** | Authenticate to Exchange Online using your alias. On connect, the app fetches your current OOF status. |
| **Enable Scheduled Auto Reply** | Sets OOF to *Scheduled* mode using your configured shift hours and work days. The start/end times are calculated automatically based on the current day. |
| **Set Vacation OOF** | Pick a return date and the app sets a *Scheduled* OOF that disables automatically when you return. |
| **Refresh Status** | Shows the current Auto Reply state, start time, and end time from Exchange Online. |
| **View Current Message** | Fetches the current OOF message from Exchange Online and loads it into the message editor. |

---

#### Configuration
| Setting | Description |
|---|---|
| **Email Suffix** | Your email domain suffix (default `@microsoft.com`). Combined with your Windows username to form the mailbox identity. |
| **Office Hours** | Start and end times for your shift. Used for OOF scheduling and the `[OFFICE HOURS]` template placeholder. |
| **Work Days** | Select your working days (Sunday–Saturday). Preset buttons available: Mon–Fri, Sun–Wed (4×10), Wed–Sat (4×10). |
| **Auto Reply State** | Manually set OOF to Enabled, Disabled, or Scheduled. |
| **Scheduled Task** | Create a Windows Task Scheduler job to run the script daily in CLI mode, 15 minutes after your shift start. If not running as admin, the app will prompt for elevation via UAC. |

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
| **Include Office Hours** | `[OFFICE HOURS]` | Removes your office hours from the signature block |
| **Include Work Days** | `[WORK DAYS]` | Removes your work days from the signature block |
| **Include Timezone** | `[TIMEZONE]` | Removes the timezone from the signature block |

> Toggling any option immediately re-renders the template in the editor and preview.

> **Signature note:** The signature is auto-generated from your Windows username and email alias. The display name is derived by splitting your alias (e.g. `aarosanders` → `Aaro Sanders`). Because the split is a best guess, **double-check that your name appears correctly** in the preview before applying. If it's wrong, you can edit it directly in the Edit tab.

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

### CLI Mode
For automation and scheduled tasks:

```powershell
# Daily scheduled auto-reply (uses saved config)# Skips if a vacation/extended OOF is active.\AAOOF-GUI.ps1 1

# Vacation mode until a specific date
.\AAOOF-GUI.ps1 '2026/04/14'
```

---

## Getting Started

### Prerequisites
- **PowerShell 5.1+** (Windows PowerShell)
- **Exchange Online Management** module (auto-installed on first connect if missing)
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
| `config.json` | User settings: shift times, work days, alias, suffix (gitignored) |
| `AAOOF-GUI.xaml` | WPF UI layout |
| `normal_oof.html` | Normal OOF template |
| `vacation_oof.html` | Vacation OOF template |
| `sick_oof.html` | Sick OOF template |
| `holiday_oof.html` | Holiday OOF template |
| `message.html` | Last-applied message (gitignored) |
| `message.html.bak` | Auto-backup of previous message (gitignored) |

### Custom Templates

You can edit the HTML template files directly or create your own. Supported placeholders:

| Placeholder | Replaced with |
|---|---|
| `[OFFICE HOURS]` | Your configured shift start – end times (e.g. `8:00 AM - 5:00 PM`) |
| `[WORK DAYS]` | Your configured work days (e.g. `Monday, Tuesday, Wednesday, Thursday, Friday`) |
| `[TIMEZONE]` | Your local timezone display name |
| `[RETURN DATE]` | The return date selected in the Vacation date picker |
| `[SIGNATURE]` | Auto-generated signature block: display name, office hours/timezone/work days, and email address |
