# SPED Status Reports V1.1 Roadmap

## Implemented In This Round
- `AuditLog` sheet to capture create, update, and document-link changes with timestamps
- separate `Service Notes` and `Variance Explanation` fields
- stronger date-order validation with clearer error messages
- lightweight `Settings` sheet for warning window, timeline colors, and optional uploads folder override
- dashboard summary cards for `Total Active`, `Overdue`, `Due This Week`, `Due This Month`, and `Completed`
- admin archive flow with `ArchiveCases`
- restore flow for archived cases
- summary report sheets by case type and by evaluator
- summary report sheet by district and case type
- selected-row case opening from dashboards and the `Cases` sheet
- duplicate detection based on `Student Name + DOB + Campus` in addition to `Student ID`
- upload unlink and delete controls in the sidebar
- due-date validation harness sheet
- quick sidebar filters for overdue, due-this-week, and evaluator-specific active cases
- saved dashboard views generated from a `DashboardViews` config sheet
- archive/restore confirmation safeguards before bulk changes run

## Recommended Next
- add richer dashboard filter presets and saved views
- add more explicit upload ownership tracking before permanent delete
- add saved-view sharing guidance and a few district-specific starter templates
- add archive/restore audit summaries to the dashboard cards

## V2 Direction
- Gmail draft generation after milestone updates
- Google Calendar event creation for upcoming deadlines and ARD meetings
- opt-in reminders with district-safe message templates
