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

## Recommended Next
- add delete/unlink support for uploaded documents
- add a validation/test harness sheet for due-date calculations

## V2 Direction
- Gmail draft generation after milestone updates
- Google Calendar event creation for upcoming deadlines and ARD meetings
- opt-in reminders with district-safe message templates
