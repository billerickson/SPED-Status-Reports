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
5. Run `installSpedStatusReports()` once from the Apps Script editor.
6. Refresh the spreadsheet.
7. Use the `SPED Status Reports` custom menu to open the app.

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
  - add one row per non-instructional date
  - include the district name and closure date
  - these rows drive the 15-school-day and FIIE timeline calculations

### Building 5 District Calendars In One Sheet
Use one shared `DistrictCalendars` sheet, not five different tabs.

Structure it like this:

| District | NonInstructionalDate | Note |
| --- | --- | --- |
| District A | 08/12/2026 | Staff development |
| District A | 11/23/2026 | Thanksgiving break |
| District B | 08/19/2026 | Staff development |
| District B | 10/12/2026 | Fall holiday |
| District C | 09/07/2026 | Labor Day |
| District D | 11/25/2026 | Thanksgiving break |
| District E | 12/21/2026 | Winter break |

How it works:
1. Each case already stores its district.
2. When the app calculates dates, it looks only at `DistrictCalendars` rows matching that district.
3. That means one tab can hold all 5 calendars safely, as long as every closure date is tagged with the correct district name.

Recommended setup for 5 districts:
- Keep the `District` name exactly the same in `Districts`, `Campuses`, and `DistrictCalendars`.
- Add every non-instructional date for all 5 districts into `DistrictCalendars`.
- If two districts share the same holiday, enter two rows, one for each district.

After making admin changes:
1. Run `SPED Status Reports -> Refresh Dashboard`.
2. Reopen the sidebar if you want the latest dropdown values to reload immediately.

## Behavior changes from Excel
- The app opens from a custom Google Sheets menu instead of `Workbook_Open`.
- Forms are rendered in a sidebar using HTML Service instead of VBA `UserForm`s.
- Backend sheets are hidden and can be protected, but file sharing permissions remain the main security boundary.
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
