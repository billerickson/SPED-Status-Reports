# SPED Status Reports V1.1 Roadmap

## Implemented In This Round
- `AuditLog` sheet to capture create, update, and document-link changes with timestamps
- separate `Service Notes` and `Variance Explanation` fields
- stronger date-order validation with clearer error messages
- lightweight `Settings` sheet for warning window, timeline colors, and optional uploads folder override

## Recommended Next
- add dashboard summary cards for `Overdue`, `Due This Week`, `Due This Month`, and `Completed`
- add an `Archive` sheet plus admin action to move completed cases out of the active working set
- add direct dashboard links or buttons to open a case from the sidebar by `Case ID`
- add duplicate detection based on `Student Name + DOB + Campus` in addition to `Student ID`
- add delete/unlink support for uploaded documents
- add a validation/test harness sheet for due-date calculations

## V2 Direction
- Gmail draft generation after milestone updates
- Google Calendar event creation for upcoming deadlines and ARD meetings
- opt-in reminders with district-safe message templates
