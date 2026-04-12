# SPED Status Reports Excel VBA Plan

## Summary
Build a workbook-first SPED case tracker in Excel VBA with an app-like `UserForm` workflow over hidden/protected sheets. Phase 1 stays local to Excel and supports both `Initial` and `Re-evaluation` cases, milestone updates, automatic status changes, multiple document links, multi-district calendar-aware due dates, and a live dashboard. Phase 2 is documented only: optional Gmail/Google Calendar integration, which the user wants but should not be part of the first confidential-data build.

## Implementation Changes
### Workbook Interface And Structure
- Create a launcher form that asks `Initial` vs `Re-evaluation`, then `New Case` vs `Update Existing`.
- Use these primary interfaces:
  - `frmLauncher`
  - `frmCaseIntake`
  - `frmCaseUpdate`
  - `frmAdmin`
- Use hidden/protected backend sheets with Excel Tables for:
  - cases
  - case documents
  - districts
  - campuses
  - lead evaluators
  - district calendars / non-instructional days
  - dashboard output
- Open users into the launcher form, keep backend sheets `VeryHidden`, and gate admin/calendar maintenance plus due-date overrides behind an admin-only path.

### Data Model And Workflow
- Generate a unique `Case ID` for every intake; search/update by `Student ID` first, then let staff choose among matching open cases.
- Block duplicate open cases of the same type for the same student unless an admin intentionally creates one.
- Common stored fields:
  - Case ID
  - Case Type
  - Student Name
  - Student ID
  - DOB
  - Campus
  - District
  - Lead Evaluator
  - Status
  - Created/Updated timestamps
- `Initial` intake fields:
  - referral received date
  - response due date
  - projected consent date
  - actual consent date
  - projected FIIE due date
  - actual FIIE date
  - projected ARD due date
  - actual ARD date
  - needed areas/services checkboxes
  - service notes
  - multiple document links
- `Re-evaluation` intake fields:
  - user-entered reevaluation due date
  - same needed areas/services
  - same milestone update capability
  - multiple document links
- Needed areas/services are checkbox-driven with shared notes support for:
  - School Psychologist
  - Occupational Therapist
  - Physical Therapist
  - Counseling Evaluation
  - FBA
  - Speech Pathologist
  - VI
  - DHH
  - Language Dominance/Bilingual

### Timeline And Status Rules
- Support multiple district calendars from admin tables; all school-day calculations use the selected district's non-instructional dates.
- `Initial` date logic:
  - `Response Due Date` = 15 school days from referral received date using district calendar.
  - Before actual consent exists, assume provisional consent on the response due date for planning.
  - Project all later dates from that provisional consent until actual consent is entered.
  - Recalculate projected and final due dates whenever milestone data changes.
  - Calculate FIIE due from consent using the district calendar.
  - Calculate ARD due from the FIIE completion date.
- Status automation:
  - `Referral Received` from case creation until consent date is entered.
  - `In Progress` once consent exists and until ARD date is entered.
  - `Complete` when ARD date is entered.
- Allow admin-only due-date overrides with a required reason, but store only the latest values, not a separate audit log.

### VBA Deliverable Shape
- Deliver pasteable VBA text organized into standard modules plus clear setup steps for the required forms/sheets, since plain pasted code cannot recreate full `UserForm` designer layouts by itself.
- Recommended module split:
  - startup/navigation
  - case CRUD/search
  - timeline calculations
  - validation and duplicate checks
  - dashboard refresh
  - workbook protection/admin utilities

### Dashboard
- Build a dashboard sheet with filters for:
  - district
  - campus
  - lead evaluator
  - case type
  - status
- Show key dates and color-code cases as upcoming, overdue, or complete.
- Include both projected and actual milestone dates where relevant so planning dates update cleanly as milestones are entered.

### Phase 2 Note
- Document Gmail/Google Calendar integration as a later enhancement only.
- If implemented later, it should be treated as a separate security review because the user wants full case details in outbound items, which conflicts with the local-confidentiality goal.

## Public Interfaces / Types
- User-facing workbook interfaces:
  - launcher form
  - intake form
  - update/search form
  - admin form
  - dashboard sheet
- Core logical entities:
  - `Case`
  - `CaseDocument`
  - `DistrictCalendar`
  - `Campus`
  - `Evaluator`
- Core enums / controlled values:
  - `CaseType = Initial | Re-evaluation`
  - `Status = Referral Received | In Progress | Complete`

## Test Plan
- Create `Initial` case and verify response due date = 15 school days by district calendar.
- Verify projected consent/FIIE/ARD dates appear immediately at intake and recalculate after actual consent/FIIE dates are entered.
- Verify `Re-evaluation` case accepts manual due date and follows the same update/status flow.
- Verify status changes exactly at creation, consent entry, and ARD entry.
- Verify duplicate open same-type case is blocked for the same student.
- Verify student-ID search recalls existing cases correctly when a student has multiple historical cases.
- Verify campus-to-district mapping drives the correct evaluator dropdowns and calendar logic.
- Verify admin override requires a reason and changes dashboard output.
- Verify multiple document links can be added, stored, and recalled.
- Verify overdue/upcoming dashboard colors update after each milestone edit.

## Assumptions And Defaults
- Phase 1 is Excel-only and does not actually implement Gmail/Google Calendar integration.
- Hidden/protected Excel sheets are considered sufficient for v1 even though Excel is not true application-grade security.
- The build will keep only latest field values, not a separate audit history.
- The final implementation deliverable should be a pasteable VBA text file plus exact workbook setup instructions for forms and sheets.
- The VBA should be organized so an Excel user can copy modules into the workbook, create the named forms, and run a setup macro to generate the required sheets and tables.
