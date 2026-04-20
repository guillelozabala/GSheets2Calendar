# GSheets2Calendar

A Google Apps Script to sync a Google Spreadsheet with a Google Calendar.

It is based on [this tutorial](https://workspace.google.com/blog/productivity-collaboration/g-suite-pro-tip-how-to-automatically-add-a-schedule-from-google-sheets-into-calendar).

The original code duplicated events when the sync was run more than once. This version is designed to be idempotent: each spreadsheet row gets a hidden sync ID and stores the matching Google Calendar event ID, so later syncs update or delete the same event instead of creating another copy.

Five visible columns are included in the schedule the code was made for: Presenters, Starts, Ends, Location, and Notes. The script reads rows dynamically from row 3 downward, so adding events below the original range no longer requires a code change.

The script will add and hide two metadata columns:

- `GSheets2Calendar Sync ID`
- `GSheets2Calendar Event ID`

Do not edit those hidden columns manually. They are what make repeat syncs safe.

The instructions are the same:

> Create a Google Calendar (e.g., "X")
> 
> Obtain the ID from *My calendars>Options for X>Settings and sharing>Integrate calendar>Calendar ID*
> 
> Create your schedule in a Google Spreadsheet
> 
> Go to *Extensions>Apps Script* and paste the code at `calendar_app.js`
> 
> In Apps Script, go to *Project Settings > Script properties* and add `ADVISING_CALENDAR_ID` with your Google Calendar ID as the value
>
> If needed, adjust `ADVISING_SYNC_CONFIG.HEADER_ROW`, `FIRST_DATA_ROW`, or `COLUMNS` to match your sheet layout

Note that dates (columns 2 and 3 in this code) must be formatted as "dd/MM/yyyy HH:mm:ss".

Moreover, if the event is named "Up For Grabs" or "-" (column 1), no event will be created.

The first sync also migrates old events created by the previous script version when it can match them by title and time. Managed events are tagged in the calendar description with `[GSheets2Calendar]` and a source ID. The cleanup pass removes stale managed events and duplicates inside the configured cleanup window.

The first function will sync your spreadsheet to your calendar. The second will create the "Sync to Calendar" button in your spreadsheet.

**Any changes to the spreadsheet won’t automatically reflect in the calendar. You must click the "Sync to Calendar" button to update it manually.**

After each run, the spreadsheet shows a toast summary such as created, updated, deleted, skipped, and warnings. If warnings appear, open the Apps Script execution log for row-level details.

The calendar ID is intentionally stored in Apps Script's Script Properties instead of in this repository. Do not paste private calendar IDs, API keys, OAuth secrets, or service account credentials directly into the source code.
