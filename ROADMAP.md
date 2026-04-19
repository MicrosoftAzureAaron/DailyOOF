# DailyOOF Roadmap

This document tracks the planned feature roadmap for DailyOOF and the versioning model used for live rollout.

## Versioning

- Main script track (`AAOOF-GUI.ps1`) uses minor+patch feature slices (for example `1.7.0` -> `1.7.1`).
- Portable script track (`AAOOF-Portable.ps1`) is versioned independently starting at `1.0.0`.
- Main and portable releases can move at different speeds, but each release should still be implemented, validated, and pushed independently.

## Release Plan

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

- `1.8.1`: Refine template defaults based on live usage.
- `1.8.2`: Add additional placeholder validation and preview warnings.

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