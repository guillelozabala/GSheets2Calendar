# sheets_to_calendar

A Google Apps Script to sync a Google Spreadsheet with a Google Calendar.

It is based on [this tutorial](https://workspace.google.com/blog/productivity-collaboration/g-suite-pro-tip-how-to-automatically-add-a-schedule-from-google-sheets-into-calendar). 

The original code duplicates events. This one fixes the issue and prevents empty slots in the calendar. I'm sure there are smarter ways to do it, but this version has proved relatively fast and reliable.

Five columns are included in the schedule the code was made for: Presenters, Starts, Ends, Location, and	Notes. Any changes to the number or position of columns require corresponding changes in the code.

The instructions are the same:

> Create a Google Calendar (e.g., "X")
> 
> Obtain the ID from *My calendars>Options for X>Settings and sharing>Integrate calendar>Calendar ID*
> 
> Create your schedule in a Google Spreadsheet
> 
> Go to *Extensions>Apps Script* and paste the code at `calendar_app.js`
> 
> Set `calendarID = "[ID]"` in the code (replace `[ID]` with your Calendar ID) (line 5)
>
> Adjust `var signups` to match your schedule's cell range (line 15)

Note that dates (columns 2 and 3 in this code) must be formatted as "dd/MM/yyyy HH:mm:ss".

Moreover, if the event is named "Up For Grabs" or "-" (column 1), no event will be created.

The first function will sync your spreadsheet to your calendar. The second will create the "Sync to Calendar" button in your spreadsheet. 

**Any changes to the spreadsheet won’t automatically reflect in the calendar. You must click the "Sync to Calendar" button to update it manually.**

