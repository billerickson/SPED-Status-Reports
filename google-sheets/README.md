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

## Behavior changes from Excel
- The app opens from a custom Google Sheets menu instead of `Workbook_Open`.
- Forms are rendered in a sidebar using HTML Service instead of VBA `UserForm`s.
- Backend sheets are hidden and can be protected, but file sharing permissions remain the main security boundary.
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
