# DailyOOF Roadmap

This document tracks the planned feature roadmap for DailyOOF and the versioning model used for live rollout.

## Versioning

- Main script track (`AAOOF-GUI.ps1`) uses minor+patch feature slices (for example `1.7.0` -> `1.7.1`).
- Portable script track (`AAOOF-Portable.ps1`) is versioned independently starting at `1.0.0`.
- Main and portable releases can move at different speeds, but each release should still be implemented, validated, and pushed independently.

## Release Plan

## Current State (After 1.9.29)

Completed and validated in main track:

| Version | Summary |
|---------|---------|
| `1.8.0` | Backup engineer, team alias, support link placeholders |
| `1.8.1` | Quick Actions UX polish, connection button state cues |
| `1.8.3` | Strict version comparison — no false update prompts when local is newer |
| `1.8.4` | Placeholder validation hardening with per-field guidance |
| `1.8.5` | Quick Actions button width parity |
| `1.8.6` | Update panel messaging: newer / up-to-date / local-newer |
| `1.8.7` | Button shape parity; XAML version token + auto-refresh on mismatch |
| `1.8.8` | Config-field emptiness checks in `Get-TemplateWarnings` |
| `1.8.9` | Fix first-run `$script:EnableTemplateAutoDownload` initialization |
| `1.9.0` | Auto-populate Full Name from EXO `DisplayName` post-connect |
| `1.9.1` | Prompt user for name when EXO lookup fails and field is blank |
| `1.9.2` | First-run auto-connect via `ContentRendered` hook |
| `1.9.3` | Fix EXO name lookup — remove invalid `FirstName`/`LastName` properties |
| `1.9.4` | Live elapsed-time connecting window during EXO auth (separate STA runspace) |
| `1.9.5` | Diagnostics export + EXO profile enrichment (role from recipient title) |
| `1.9.6` | Reduce false-positive template warnings when generic fallbacks are not actually used |
| `1.9.7` | Add Cancel button to connecting window with graceful post-auth abort |
| `1.9.8` | Normalize Check for Updates and Export Diagnostics button sizes |
| `1.9.9` | Fix startup null-control sizing hotfix for update/diagnostics buttons |
| `1.9.28` | Add OOF audience selection (Internal Only, External Only, Both) with persisted config and default Both |
| `1.9.29` | Harden scheduled task defaults with daily + logon triggers, start-when-available catch-up, retry behavior, and wake-to-run |

Completed and validated in portable track:

| Version | Summary |
|---------|---------|
| `1.1.3` | Add portable OOF audience selector (Internal Only, External Only, Both) and apply to Enabled/Scheduled state actions |

## Next Execution Queue

1. `1.9.5` Diagnostics Foundation
- Scope: add a one-click diagnostics export capturing connection status, task status, config source, EXO identity, and recent update-check outcome.
- Acceptance: output is readable plain text or JSON; works in both connected and disconnected states.
- Validation: run in connected/disconnected states and with task present/missing.

2. `1.9.6` False-Positive Warning Reduction
- Scope: tune `Get-TemplateWarnings` conditions based on live findings; suppress warnings for fields that have valid fallback values.
- Acceptance: warnings are actionable and map to real blocking states only.
- Validation: regression pass through Quick Actions, Automation, and Current OOF flows.

3. `1.9.7` Connecting Window — cancel support
- Scope: add a Cancel button to the connecting window that aborts the EXO auth attempt gracefully.
- Acceptance: user can cancel mid-connect without the UI freezing or leaving a dangling session.
- Validation: test cancel during MFA prompt and before auth completes.

### 1.7.x Scheduled Task Management and Validation

Goals:

- Show whether the `AAOOF` task exists.
- Show task status, next run time, last run time, last result, and target script path.
- Add buttons to refresh task status, open Task Scheduler, and run the task on demand.
- Add task configuration options in the GUI, starting with a configurable start offset in minutes.
- Add configuration validation for task creation and scheduled OOF actions.

Patch goals:

- `1.7.1`: Fix task state edge cases and improve task result messaging.
- `1.7.2`: Add enable/disable/delete controls if needed after live feedback.
- `1.7.3`: Detect scheduled task script-path mismatches and add one-click repair.

### 1.8.x Template and Contact Improvements

Goals:

- Add richer template placeholders such as backup engineer, team alias, and support link.
- Improve built-in templates for internal and external audiences.
- Add stronger placeholder warnings before applying messages.
- Consider separate template presets for vacation, sick, holiday, training, and limited availability.

Patch goals:

- `1.8.0`: Add backup engineer, team alias, and support link placeholders with config/UI support.
- `1.8.1`: Polish Quick Actions UX (connection button state cues and button sizing consistency).
- `1.8.2`: Add additional placeholder validation and preview warnings.
- `1.8.3`: Fix update detection so update prompts only appear when GitHub version is strictly newer than local.
- `1.8.4`: Add stronger unresolved-placeholder blocking/warnings and clearer per-field guidance.
- `1.8.5`: Enforce matching Quick Actions button widths for status/message actions.
- `1.8.6`: Improve update panel messaging for newer/up-to-date/local-newer scenarios.
- `1.8.7`: Normalize Quick Actions button shape (width/height/margin parity); add XAML version token and auto-refresh on mismatch.

### 1.9.x Diagnostics and Operational Safety

Goals:

- Add a diagnostics view or export option.
- Record recent task and update activity for troubleshooting.
- Add clearer startup and validation error messaging.
- Improve status indicators for Exchange connection, OOF state, and automation readiness.

Patch goals:

- `1.9.1`: Reduce false-positive warnings.
- `1.9.2`: Improve troubleshooting output based on live testing.

### Portable Track (1.x) Simplified GUI Mode

Goals:

- Ship this mode as a separate script file so the main full-featured GUI remains unchanged.
- Deliver a single-file portable script with embedded GUI resources.
- Focus only on managing auto-reply state (Enabled / Disabled / Scheduled).
- Exclude template editing and message apply flows from this portable mode.
- Exclude scheduled task and automation management from this portable mode.
- Provide clear in-app guidance that OOF message content should be managed in Outlook.

Patch goals:

- `1.0.0`: Build the portable-mode shell and minimal GUI state controls.
- `1.0.1`: Add validation, simplified status messaging, and Outlook guidance polish.
- `1.1.0`: Auto-detect mailbox, improve EXO setup/connection reuse, add connecting window with timer/cancel.
- `1.1.1`: Add work-day presets (Mon–Fri, Sun–Wed, Wed–Sat) and schedule calculation that respects the next working day instead of assuming a 7-day work week.
- `1.1.2`: Increase portable window height so the full state-management UI is visible without scrolling or manual resize.
- `1.1.3`: Add OOF audience selector and behavior parity with main GUI for state actions.

## Implementation Rules

- Every feature slice should include code comments for new logic.
- Every feature slice should include validation and at least one explicit verification step before release.
- Live updates should be pushed in small increments so the production script can be tested safely.