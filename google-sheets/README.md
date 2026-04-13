# SPED Status Reports for Google Sheets

This folder contains a Google Sheets adaptation of the SPED Status Reports workflow.

## Why changes were required
- Excel VBA and `UserForm` interfaces do not run in Google Sheets.
- Google Sheets automation is built with Google Apps Script, custom menus, and HTML sidebars/dialogs instead.
- Google Sheets supports hidden and protected sheets, but it does not have Excel's `VeryHidden` behavior.

## Deliverables
- [Code.gs](/Users/billerickson/Downloads/SPED-Excel/google-sheets/Code.gs)
- [Sidebar.html](/Users/billerickson/Downloads/SPED-Excel/google-sheets/Sidebar.html)
- [appsscript.json](/Users/billerickson/Downloads/SPED-Excel/google-sheets/appsscript.json)
- [V1_1_ROADMAP.md](/Users/billerickson/Downloads/SPED-Excel/google-sheets/V1_1_ROADMAP.md)

## Setup
1. Create or open the Google Sheet that will hold SPED Status Reports.
2. Open `Extensions -> Apps Script`.
3. Replace the default script project files with the contents of:
   - `Code.gs`
   - `Sidebar.html`
   - `appsscript.json`
4. Update the `ADMIN_ACCESS_CODE` constant in `Code.gs`.
5. Optional: set `UPLOADS_FOLDER_ID` in `Code.gs` if uploaded files should go to a specific Google Drive folder.
   If left blank, the script creates or reuses `SPED Status Report Uploads` in the spreadsheet's parent folder.
6. Add admin Google account emails to `ADMIN_EDITOR_EMAILS` in `Code.gs`.
   Use full email addresses such as `name@district.org`.
7. Run `installSpedStatusReports()` once from the Apps Script editor.
8. Refresh the spreadsheet.
9. Use the `SPED Status Reports` custom menu to open the app.

Because uploads now store files in Google Drive, the manifest also needs the Drive scope from `appsscript.json`.
V2 Gmail drafts and Calendar events also require Gmail and Calendar scopes, so Apps Script will ask for fresh authorization after you paste the updated files.

## How To Update Calendars And Dropdown Lists
Use the `SPED Status Reports` menu in the spreadsheet.

1. Open one of these admin sheets:
   - `Open Districts (Admin)`
   - `Open Campuses (Admin)`
   - `Open Evaluators (Admin)`
   - `Open Service Contacts (Admin)`
   - `Open Calendars (Admin)`
   - `Open Due Date Tests (Admin)`
   - `Open Dashboard Views (Admin)`
   - `Open Settings (Admin)`
   - `Open Audit Log (Admin)`
   - `Open Archive (Admin)`
2. Enter the admin access code.
3. Edit the rows directly in the revealed sheet.

Use these sheets for each type of maintenance:
- `Districts`
  - one row per district
  - update `ResponseSchoolDays`, `FIIESchoolDays`, and `ARDCalendarDays`
  - set `Active = Yes` for districts that should appear in dropdowns
- `Campuses`
  - one row per campus
  - map each campus to its district
  - set `Active = Yes` for campuses that should appear in dropdowns
- `Evaluators`
  - one row per evaluator
  - set `Active = Yes` for evaluators that should appear in dropdowns
- `ServiceContacts`
  - one row per `District + Service` combination
  - use `EmailTo` for milestone update recipients for that district/service
  - use `EmailCc` for optional copied recipients
  - recipient lists can be separated by commas, semicolons, or line breaks
  - set `Active = Yes` for district/service groups that should receive automatic milestone updates
- `DistrictCalendars`
- `Settings`
  - v2 communication settings also live here:
    - `AutoSendNewReferralAssignments`
    - `NewReferralAssignmentTo`
    - `NewReferralAssignmentCc`
    - `AutoSendMilestoneUpdates`
    - `IncludeLeadEvaluatorOnNewAssignments`
    - `IncludeLeadEvaluatorOnMilestoneUpdates`
    - `NotificationEmailTo`
    - `NotificationEmailCc`
    - `NotificationEmailBcc`
    - `NotificationReplyTo`
    - `NotificationFromAlias`
    - `NotificationSenderName`
    - `NotificationCalendarId`
    - `CalendarPopupReminderMinutes`
    - `DeadlineReminderHour`
    - `DeadlineReminderDurationMinutes`
    - `ARDEventHour`
    - `ARDEventDurationMinutes`
- `AuditLog`
- `ArchiveCases`
- `DueDateTests`
- `DashboardViews`
- `SummaryByCaseType`
- `SummaryByEvaluator`
- `SummaryByDistrictCaseType`
  - the sheet is preloaded with dates from January 1, 2026 through June 1, 2027
  - weekends default to `No`
  - weekdays default to `Yes`
  - each district gets its own column
  - change any district/date cell to `No` when that district is not instructional that day
  - these `Yes` / `No` values drive the 15-school-day and FIIE timeline calculations

### Building 5 District Calendars In One Sheet
Use one shared `DistrictCalendars` sheet, not five different tabs.

Structure it like this:

| Date | Weekday | District A | District B | District C | District D | District E |
| --- | --- | --- | --- | --- | --- | --- |
| 08/12/2026 | Wednesday | No | Yes | Yes | Yes | Yes |
| 08/13/2026 | Thursday | Yes | Yes | Yes | Yes | Yes |
| 09/07/2026 | Monday | No | No | No | No | No |
| 11/23/2026 | Monday | No | Yes | Yes | Yes | Yes |

How it works:
1. Each case already stores its district.
2. When the app calculates dates, it checks the selected district's column for each date.
3. `Yes` means instructional and counts toward due dates.
4. `No` means non-instructional and does not count.

Recommended setup for 5 districts:
- Keep the district names exactly the same in the `Districts` sheet and the `DistrictCalendars` column headers.
- If all districts are off on the same day, set that row to `No` across all district columns.
- If only one district is off, change only that district's cell to `No`.
- You do not need to enter weekends manually; the calendar grid preloads them as `No`.
- If you add or remove districts later, run `SPED Status Reports -> Sync Calendar Grid`.

After making admin changes:
1. Run `SPED Status Reports -> Refresh Dashboard`.
2. Reopen the sidebar if you want the latest dropdown values to reload immediately.

## Where The Database Lives
The data is stored inside this same Google Sheet on visible protected tabs:
- `Cases`
- `CaseDocuments`
- `Districts`
- `Campuses`
- `Evaluators`
- `ServiceContacts`
- `DistrictCalendars`
- `Settings`
- `AuditLog`
- `DueDateTests`
- `DashboardViews`

They are visible to everyone who can open the spreadsheet.
Direct sheet edits are limited by Google Sheets protection to the admin accounts listed in `ADMIN_EDITOR_EMAILS`.

To manage the protections:
1. Open the sidebar.
2. Open the `Admin` section.
3. Enter the admin code.
4. Use `Reapply Protection` if you changed admin accounts or copied the sheet.

For milestone updates:
1. Choose `Update Existing`.
2. Search by `Student ID`.
3. Select the existing case.
4. Then edit milestone fields and save.

To open a case directly from a dashboard or the `Cases` sheet:
1. Click the row for that case in the spreadsheet.
2. Use `SPED Status Reports -> Open Selected Case`, or click `Open Selected Row` in the sidebar.

To restore an archived case:
1. Open `ArchiveCases`.
2. Click the archived case row you want to restore.
3. Use `SPED Status Reports -> Restore Selected Archived Case (Admin)` or the sidebar admin button.
4. Enter the admin code when prompted.

To run the due-date validation harness:
1. Open `DueDateTests`.
2. Enter or update the expected due dates for any scenarios you want to verify.
3. Run `SPED Status Reports -> Refresh Due Date Tests`.
4. Review the `Result` column:
   - `PASS` means the calculated dates matched the expected dates
   - `FAIL` means one or more expected dates did not match
   - `CHECK` means expected dates have not been filled in yet

To use the quick lists in the sidebar:
1. Click `Overdue` to load all open overdue cases.
2. Click `Due This Week` to load cases due in the next 7 days.
3. Choose an evaluator and click `My Evaluator Cases` to load that evaluator's active cases.

To remove uploads:
1. Open the case in the sidebar.
2. In the uploads list, use `Unlink` to remove the document from the case only.
3. Use `Delete` to remove the document from the case and send the linked Google Drive file to trash when the script has permission.

## V2 Gmail And Calendar
The manual v2 communication tools are intentionally conservative:
- Gmail actions create drafts only
- Calendar actions create events only
- nothing runs on a timer yet

Automatic email notifications are a separate v2 workflow:
- new referral assignment emails can auto-send on case creation
- milestone update emails can auto-send on case update
- these use Settings plus the district-aware `ServiceContacts` sheet

Before using v2:
1. Open `Settings`.
2. Set `AutoSendNewReferralAssignments = Yes` if new case creation should email the assignment group automatically.
3. Fill in `NewReferralAssignmentTo` and optional `NewReferralAssignmentCc`.
4. Set `AutoSendMilestoneUpdates = Yes` if case milestone updates should email the checked service contacts automatically.
5. Leave `IncludeLeadEvaluatorOnNewAssignments` and `IncludeLeadEvaluatorOnMilestoneUpdates` set to `Yes` if the lead evaluator should always be included automatically.
6. Make sure the `Evaluators` sheet has the correct email for each lead evaluator.
7. Open `ServiceContacts` and fill in the email groups for each `District + Service` row.
8. Fill in `NotificationEmailTo` with the default recipients for manual draft creation.
9. Optional: fill in CC, BCC, reply-to, sender name, and Gmail alias settings.
10. Optional: fill in `NotificationCalendarId` if events should go to a dedicated SPED calendar.
   If left blank, events go to the authorized user's default calendar.
11. Save the settings and reauthorize the script if prompted.

If you already had an older `ServiceContacts` sheet, run `Install / Repair Workbook` once after pasting this update so the `District` column is added and preserved correctly.

Available v2 case actions:
- `Send Test Notification For Selected Case`
- `Due Soon Draft`
- `Overdue Draft`
- `Create ARD Event`
- `Deadline Reminder`

How they work:
- `Send Test Notification For Selected Case` resolves the same service-team and lead-evaluator recipients used by milestone updates and attempts a live email send for the selected case row
  - invalid recipient values are ignored and listed in `AuditLog` instead of blocking the whole email
- `Due Soon Draft` creates a Gmail draft using the case's current primary deadline.
- `Overdue Draft` creates a Gmail draft using the same case details but with overdue language.
- `Create ARD Event` uses `ARD Scheduled Date` and the configured ARD hour/duration settings.
- `Deadline Reminder` uses the case `PrimaryDeadline` and the configured reminder hour/duration settings.

Each v2 action also writes an entry to `AuditLog`.

Automatic email behavior:
- a new `Initial` case save can automatically send a referral assignment email to the configured assignment group
- an update save can automatically send an email for milestone changes and any case status change to all active `ServiceContacts` rows that match both the case district and the checked services on the case
- the lead evaluator can also be included automatically on both email types using the `Evaluators.Email` value for the selected lead evaluator
- case saves still complete even if an automatic email cannot be sent; the app shows a warning and logs the failure in `AuditLog`

To manage saved dashboard views:
1. Open `DashboardViews`.
2. Add or edit rows using these columns:
   - `ViewName`
   - `District`
   - `Evaluator`
   - `CaseType`
   - `Status`
   - `DeadlineBucket`
   - `Active`
   - `Description`
3. Use `DeadlineBucket` values like `Overdue`, `DueThisWeek`, or `DueThisMonth`.
4. Run `SPED Status Reports -> Refresh Saved Views`.
5. The app will create or update matching sheets named `View - <ViewName>`.

Bulk admin safeguards now in place:
- archiving prompts with the number of completed cases before it runs
- archived-case restore prompts with the selected `Case ID`
- document delete still requires confirmation in the sidebar before sending a Drive file to trash

## Behavior changes from Excel
- The app opens from a custom Google Sheets menu instead of `Workbook_Open`.
- Forms are rendered in a sidebar using HTML Service instead of VBA `UserForm`s.
- Database sheets are visible, but direct edits are restricted with Google Sheets sheet protection.
- `New Case -> Initial` now shows only intake fields; milestone dates and needed services are reserved for `Update Existing`.
- Evaluators are now maintained as one shared active list and are no longer tied to district.
- The same case model is preserved:
  - `Initial` and `Re-evaluation`
  - milestone updates
  - duplicate open-case blocking by `Student ID`, plus same `Student Name + DOB + Campus`
  - grade level and expanded milestone tracking
  - multiple document links plus file uploads
  - separate `Service Notes` and `Variance Explanation`
  - audit logging for creates, updates, and document changes
  - stronger date-order validation
  - configurable warning colors and deadline window from `Settings`
  - district-aware instructional-day calculations
  - master dashboard plus one dashboard sheet per district
  - dashboard summary cards for active work, overdue cases, due this week, due this month, and completed cases
  - summary report sheets by case type and by evaluator
  - summary report sheet by district and case type
  - admin archive flow for completed cases
  - restore flow for archived cases
  - selected-row case opening from dashboards and the `Cases` sheet
  - quick sidebar lists for overdue, due-this-week, and evaluator-specific active cases
  - upload unlink/delete controls from the sidebar
  - due-date test harness sheet for validating district timeline calculations
  - saved dashboard-view sheets generated from the `DashboardViews` config tab
  - archive and restore confirmations before bulk admin changes run
  - Gmail draft creation for due-soon and overdue case communication
  - Calendar event creation for ARD scheduling and deadline reminders
  - automatic new referral assignment emails to a configured group
  - automatic milestone update emails to the checked service groups on the case
  - status flow:
    - `Referral Received`
    - `Response Sent`
    - `Consent Received`
    - `Evaluation in Progress`
    - `Evaluation Complete`
    - `ARD Scheduled`
    - `Completed`
  - variance explanation required when an actual milestone date is later than its due date
  - dashboard status colors plus pink/red deadline highlighting

## Security note
- Confidentiality in Google Sheets depends primarily on who has access to the spreadsheet file.
- Hidden sheets and Apps Script protections reduce accidental exposure, but they are not equivalent to a dedicated secure application.

## Google references used for this adaptation
- Google Sheets UI extensions and custom menus:
  [Ui class](https://developers.google.com/apps-script/reference/base/ui)
- Sidebar/dialog UI:
  [Dialogs and sidebars](https://developers.google.com/apps-script/guides/dialogs)
- HTML-based Apps Script UI:
  [HTML Service](https://developers.google.com/apps-script/guides/html)
- Bound script model for Sheets:
  [Container-bound scripts](https://developers.google.com/apps-script/guides/bound)
- Sheet/range protection:
  [Protection class](https://developers.google.com/apps-script/reference/spreadsheet/protection)
- Migration note:
  [Use macros and add-ons](https://support.google.com/docs/answer/9331168?hl=en&ref_topic=9296611)
