# DailyOOF Roadmap

This document tracks the planned feature roadmap for DailyOOF and the versioning model used for live rollout.

## Versioning

- Minor version (`1.6` -> `1.7`) is used for a user-visible feature slice.
- Patch version (`1.7.0` -> `1.7.1`) is used for bug fixes, validation hardening, and follow-up polish within the same feature slice.
- Each minor release should be implemented, validated, and pushed independently so the live script can be updated and tested incrementally.

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

## Implementation Rules

- Every feature slice should include code comments for new logic.
- Every feature slice should include validation and at least one explicit verification step before release.
- Live updates should be pushed in small increments so the production script can be tested safely.