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

## Setup
1. Create or open the Google Sheet that will hold SPED Status Reports.
2. Open `Extensions -> Apps Script`.
3. Replace the default script project files with the contents of:
   - `Code.gs`
   - `Sidebar.html`
   - `appsscript.json`
4. Update the `ADMIN_ACCESS_CODE` constant in `Code.gs`.
5. Add admin Google account emails to `ADMIN_EDITOR_EMAILS` in `Code.gs`.
6. Run `installSpedStatusReports()` once from the Apps Script editor.
7. Refresh the spreadsheet.
8. Use the `SPED Status Reports` custom menu to open the app.

## How To Update Calendars And Dropdown Lists
Use the `SPED Status Reports` menu in the spreadsheet.

1. Open one of these admin sheets:
   - `Open Districts (Admin)`
   - `Open Campuses (Admin)`
   - `Open Evaluators (Admin)`
   - `Open Calendars (Admin)`
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
- `DistrictCalendars`
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
- `DistrictCalendars`

They are visible to everyone who can open the spreadsheet.
Direct sheet edits are limited by Google Sheets protection to the admin accounts listed in `ADMIN_EDITOR_EMAILS` plus the installing/admin user.

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

## Behavior changes from Excel
- The app opens from a custom Google Sheets menu instead of `Workbook_Open`.
- Forms are rendered in a sidebar using HTML Service instead of VBA `UserForm`s.
- Database sheets are visible, but direct edits are restricted with Google Sheets sheet protection.
- `New Case -> Initial` now shows only intake fields; milestone dates and needed services are reserved for `Update Existing`.
- Evaluators are now maintained as one shared active list and are no longer tied to district.
- The same case model is preserved:
  - `Initial` and `Re-evaluation`
  - milestone updates
  - duplicate open-case blocking
  - multiple document links
  - district-aware instructional-day calculations
  - auto-status and dashboard refresh

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
