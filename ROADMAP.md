# DailyOOF Roadmap

This document tracks the planned feature roadmap for DailyOOF and the versioning model used for live rollout.

## Versioning

- Main script track (`AAOOF-GUI.ps1`) uses minor+patch feature slices (for example `1.7.0` -> `1.7.1`).
- Portable script track (`AAOOF-Portable.ps1`) is versioned independently starting at `1.0.0`.
- Main and portable releases can move at different speeds, but each release should still be implemented, validated, and pushed independently.

## Release Plan

## Current State (After 1.8.6)

- Completed and validated in main track: `1.8.0`, `1.8.1`, `1.8.3`, `1.8.5`, `1.8.6`.
- Confirmed behavior from live testing: when GitHub is older than local, the app no longer auto-prompts users to update.
- In progress: visual and XAML confirmation via fresh screenshots.

## Next Execution Queue

1. `1.8.4` Placeholder Validation Hardening
- Scope: strengthen template placeholder warnings before apply/save, with clear missing-field guidance.
- Acceptance: applying a template with unresolved required placeholders shows a specific warning list and does not silently proceed.
- Validation: test with empty `BackupContact`, `TeamAlias`, and `SupportLink`; confirm warnings and corrected behavior after fields are filled.

2. `1.8.5` Quick Actions Button Size Parity
- Scope: enforce identical widths for `Refresh Status` and `View Current Message` in Quick Actions.
- Acceptance: both buttons render with matching width regardless of local XAML age.
- Validation: confirm parity in GUI and screenshot output.

3. `1.8.6` Update UX Messaging Cleanup
- Scope: make update panel/status text explicitly state one of three outcomes: newer available, up to date, or local newer than GitHub.
- Acceptance: no ambiguous "Update Available" state when local is newer.
- Validation: verify UI text with three controlled version scenarios using local/remote combinations.

4. `1.9.0` Diagnostics Foundation
- Scope: add a diagnostics tab or export action that captures connection status, task status, config source, and recent update-check outcome.
- Acceptance: one-click diagnostics output available for troubleshooting without manual log gathering.
- Validation: run diagnostics in connected/disconnected states and with task present/missing.

5. `1.9.1` False-Positive Warning Reduction
- Scope: tune warning conditions based on live findings from `1.8.4`-`1.9.0`.
- Acceptance: warnings are actionable and map to real blocking states.
- Validation: regression pass through Quick Actions, Automation, and Current OOF flows.

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

- Add richer template placeholders such as backup contact, team alias, and support link.
- Improve built-in templates for internal and external audiences.
- Add stronger placeholder warnings before applying messages.
- Consider separate template presets for vacation, sick, holiday, training, and limited availability.

Patch goals:

- `1.8.0`: Add backup contact, team alias, and support link placeholders with config/UI support.
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

## Implementation Rules

- Every feature slice should include code comments for new logic.
- Every feature slice should include validation and at least one explicit verification step before release.
- Live updates should be pushed in small increments so the production script can be tested safely.