function advisingSchedule() {

  // Open the calendar
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var calendarId = ""; // Here goes your calendar ID
  var eventCal = CalendarApp.getCalendarById(calendarId);

  // Get all existing events in the range
  var startRange = new Date("01/01/2024 15:00:00");
  var endRange = new Date("01/01/2026 15:00:00");

  var existingEvents = eventCal.getEvents(startRange, endRange);

  // Read the schedule (change accordingly, don't include headers)
  var signups = spreadsheet.getRange("A3:E20").getValues();
  
  // First loop: Create new events if no matching event exists
  for (var x = 0; x < signups.length; x++) {

    var presentation = signups[x];

    var presenter = presentation[0];
    var startTime = presentation[1];
    var endTime = presentation[2];
    var location = presentation[3];
    var description = presentation[4];

    // Flag to check if the event already exists and matches
    var eventExists = false;

    // Loop through existing events to check for a matching one
    for (var i = 0; i < existingEvents.length; i++) {
      var event = existingEvents[i];

      // Check if event matches
      if (event.getTitle() === presenter &&
          event.getStartTime().getTime() === startTime.getTime() &&
          event.getEndTime().getTime() === endTime.getTime() &&
          event.getLocation() === location &&
          event.getDescription() === description) {
        eventExists = true;
        break;
      }
    }

    // If no matching event is found, create a new one
    if (!eventExists) {
      eventCal.createEvent(presenter, startTime, endTime, {
        location: location,
        description: description
      });
    }
  }

  // Second loop: Delete old events that no longer have a corresponding signup
  for (var i = 0; i < existingEvents.length; i++) {
    var event = existingEvents[i];
    var eventHasMatch = false;

    // Loop through signups to see if the existing event has a match
    for (var x = 0; x < signups.length; x++) {

      var presentation = signups[x];
      
      var presenter = presentation[0];
      var startTime = presentation[1];
      var endTime = presentation[2];
      var location = presentation[3];
      var description = presentation[4];

      // Check if the event matches the signup
      if (event.getTitle() === presenter &&
          event.getStartTime().getTime() === startTime.getTime() &&
          event.getEndTime().getTime() === endTime.getTime() &&
          event.getLocation() === location &&
          event.getDescription() === description) {
        eventHasMatch = true;
        break;
      }
    }

    // If no match is found for the existing event, delete it
    if (!eventHasMatch) {
      event.deleteEvent();
    }
  }
}

// Update plug-in
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Sync to Calendar')
    .addItem('Update presentations', 'advisingSchedule')
    .addToUi();
}
