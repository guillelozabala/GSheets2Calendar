function advisingSchedule() {
  var sheet = SpreadsheetApp.getActiveSheet();

  var calendarId = "[ID]";
  var cal = CalendarApp.getCalendarById(calendarId);

  // Only events containing this tag will be managed (deleted/recreated).
  var TAG = "Advising Group";

  // Read signups (no headers)
  var signups = sheet.getRange("A3:E20").getValues();

  // Collect valid rows and the set of days they touch
  var rows = [];
  var dayKeys = {}; // map dayKey -> true

  for (var i = 0; i < signups.length; i++) {
    var presenter = signups[i][0];
    var startTime = signups[i][1];
    var endTime = signups[i][2];
    var location = signups[i][3];
    var description = signups[i][4];

    // Skip placeholders / empty
    if (!presenter || presenter === "Up For Grabs" || presenter === "-") continue;

    // Ensure start/end are Dates (Sheets sometimes gives strings if formatting is off)
    if (!(startTime instanceof Date) || isNaN(startTime.getTime())) continue;
    if (!(endTime instanceof Date) || isNaN(endTime.getTime())) continue;

    rows.push({
      presenter: String(presenter),
      startTime: startTime,
      endTime: endTime,
      location: location ? String(location) : "",
      description: description ? String(description) : ""
    });

    var dayKey = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy-MM-dd");
    dayKeys[dayKey] = true;
  }

  // 1) DELETE: For each day present in the sheet, delete previously synced events on that day.
  var tz = Session.getScriptTimeZone();
  var dayKeyList = Object.keys(dayKeys);

  for (var d = 0; d < dayKeyList.length; d++) {
    var key = dayKeyList[d]; // "yyyy-MM-dd"

    // Build day start/end in script timezone
    var parts = key.split("-");
    var dayStart = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 0, 0, 0);
    var dayEnd = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 23, 59, 59);

    var events = cal.getEvents(dayStart, dayEnd);
    for (var e = 0; e < events.length; e++) {
      var ev = events[e];
      var desc = ev.getDescription() || "";
      // Only delete events that this script created earlier
      if (desc.indexOf(TAG) !== -1) {
        ev.deleteEvent();
      }
    }
  }

  // 2) CREATE: Recreate events from the sheet, tagged so we can manage them next sync.
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];

    var fullDescription = row.description;
    if (fullDescription) fullDescription += "\n\n";
    fullDescription += TAG;

    cal.createEvent(row.presenter, row.startTime, row.endTime, {
      location: row.location,
      description: fullDescription
    });
  }
}

// Menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sync to Calendar')
    .addItem('Update presentations', 'advisingSchedule')
    .addToUi();
}

