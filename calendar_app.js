var ADVISING_SYNC_CONFIG = {
  CALENDAR_ID: "REDACTED_CALENDAR_ID",

  // Leave blank to use the active sheet. Set to "Hoja 1" or another tab name
  // if the spreadsheet eventually contains more than one sheet.
  SHEET_NAME: "",

  HEADER_ROW: 2,
  FIRST_DATA_ROW: 3,
  DATA_COLUMN_COUNT: 5,

  COLUMNS: {
    TITLE: 1,
    START: 2,
    END: 3,
    LOCATION: 4,
    DESCRIPTION: 5
  },

  METADATA_HEADERS: {
    SOURCE_ID: "GSheets2Calendar Sync ID",
    EVENT_ID: "GSheets2Calendar Event ID"
  },

  MANAGED_MARKER: "GSheets2Calendar",
  LEGACY_TAG: "Advising Group",
  PLACEHOLDER_TITLES: ["Up For Grabs", "-"],

  // The cleanup window is intentionally bounded so a manual calendar with years
  // of history is not scanned on every run. Desired sheet dates outside this
  // window are still included automatically.
  CLEANUP_LOOKBACK_DAYS: 370,
  CLEANUP_LOOKAHEAD_DAYS: 730,
  CALENDAR_SCAN_PADDING_DAYS: 2,

  // Set to false once every old event has the GSheets2Calendar marker.
  CLEANUP_LEGACY_EVENTS: true,

  MAX_RETRIES: 3
};

function advisingSchedule() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Another calendar sync is already running. Try again in a minute.");
  }

  var summary = newSyncSummary_();

  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getScheduleSheet_(spreadsheet);
    var calendar = CalendarApp.getCalendarById(ADVISING_SYNC_CONFIG.CALENDAR_ID);

    if (!calendar) {
      throw new Error("Calendar not found. Check ADVISING_SYNC_CONFIG.CALENDAR_ID.");
    }

    var metadata = ensureMetadataColumns_(sheet);
    var syncData = readScheduleRows_(sheet, metadata, summary);
    var calendarState = loadManagedCalendarEvents_(calendar, syncData.desiredRows);
    var claimedEventIds = {};

    syncDesiredRows_(calendar, calendarState, syncData.desiredRows, claimedEventIds, summary);
    removeUndesiredRowEvents_(calendar, calendarState, syncData.rows, claimedEventIds, summary);
    cleanupOrphanedManagedEvents_(calendarState, syncData.desiredLegacyKeys, claimedEventIds, summary);
    writeMetadata_(sheet, metadata, syncData.rows);
    reportSummary_(summary);

    return summary;
  } finally {
    lock.releaseLock();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Sync to Calendar")
    .addItem("Update presentations", "advisingSchedule")
    .addToUi();
}

function getScheduleSheet_(spreadsheet) {
  if (ADVISING_SYNC_CONFIG.SHEET_NAME) {
    var sheet = spreadsheet.getSheetByName(ADVISING_SYNC_CONFIG.SHEET_NAME);
    if (!sheet) {
      throw new Error("Sheet not found: " + ADVISING_SYNC_CONFIG.SHEET_NAME);
    }
    return sheet;
  }

  return spreadsheet.getActiveSheet();
}

function ensureMetadataColumns_(sheet) {
  return {
    sourceIdColumn: ensureHeaderColumn_(sheet, ADVISING_SYNC_CONFIG.METADATA_HEADERS.SOURCE_ID),
    eventIdColumn: ensureHeaderColumn_(sheet, ADVISING_SYNC_CONFIG.METADATA_HEADERS.EVENT_ID)
  };
}

function ensureHeaderColumn_(sheet, headerName) {
  var headerRow = ADVISING_SYNC_CONFIG.HEADER_ROW;
  var lastColumn = Math.max(sheet.getLastColumn(), ADVISING_SYNC_CONFIG.DATA_COLUMN_COUNT);
  var headers = sheet.getRange(headerRow, 1, 1, lastColumn).getDisplayValues()[0];

  for (var i = 0; i < headers.length; i++) {
    if (cleanString_(headers[i]) === headerName) {
      hideColumnIfPossible_(sheet, i + 1);
      return i + 1;
    }
  }

  var newColumn = lastColumn + 1;
  sheet.getRange(headerRow, newColumn).setValue(headerName);
  hideColumnIfPossible_(sheet, newColumn);
  return newColumn;
}

function hideColumnIfPossible_(sheet, column) {
  try {
    sheet.hideColumns(column);
  } catch (error) {
    Logger.log("Could not hide metadata column " + column + ": " + error.message);
  }
}

function readScheduleRows_(sheet, metadata, summary) {
  var firstRow = ADVISING_SYNC_CONFIG.FIRST_DATA_ROW;
  var lastRow = Math.max(sheet.getLastRow(), firstRow - 1);
  var rowCount = lastRow - firstRow + 1;
  var rows = [];
  var desiredRows = [];
  var desiredSourceIds = {};
  var desiredLegacyKeys = {};
  var seenSourceIds = {};

  if (rowCount <= 0) {
    return {
      rows: rows,
      desiredRows: desiredRows,
      desiredSourceIds: desiredSourceIds,
      desiredLegacyKeys: desiredLegacyKeys
    };
  }

  var values = sheet.getRange(firstRow, 1, rowCount, ADVISING_SYNC_CONFIG.DATA_COLUMN_COUNT).getValues();
  var sourceIds = sheet.getRange(firstRow, metadata.sourceIdColumn, rowCount, 1).getValues();
  var eventIds = sheet.getRange(firstRow, metadata.eventIdColumn, rowCount, 1).getValues();

  for (var i = 0; i < rowCount; i++) {
    var row = {
      rowNumber: firstRow + i,
      values: values[i],
      sourceId: cleanString_(sourceIds[i][0]),
      eventId: cleanString_(eventIds[i][0]),
      hasAnyScheduleData: hasAnyScheduleData_(values[i]),
      desired: null
    };

    if (row.hasAnyScheduleData || row.eventId) {
      if (!row.sourceId) {
        row.sourceId = Utilities.getUuid();
      } else if (seenSourceIds[row.sourceId]) {
        summary.warnings.push("Row " + row.rowNumber + " had a duplicate hidden sync ID. A new ID was assigned.");
        row.sourceId = Utilities.getUuid();
        row.eventId = "";
      }
      seenSourceIds[row.sourceId] = true;
    }

    row.desired = buildDesiredEventFromRow_(row, summary);
    rows.push(row);

    if (row.desired) {
      desiredRows.push(row);
      desiredSourceIds[row.sourceId] = true;
      desiredLegacyKeys[row.desired.legacyKey] = true;
    }
  }

  return {
    rows: rows,
    desiredRows: desiredRows,
    desiredSourceIds: desiredSourceIds,
    desiredLegacyKeys: desiredLegacyKeys
  };
}

function buildDesiredEventFromRow_(row, summary) {
  if (!row.hasAnyScheduleData) {
    return null;
  }

  var columns = ADVISING_SYNC_CONFIG.COLUMNS;
  var title = cleanString_(row.values[columns.TITLE - 1]);

  if (isPlaceholderTitle_(title)) {
    summary.skipped++;
    return null;
  }

  var startTime = coerceDate_(row.values[columns.START - 1]);
  var endTime = coerceDate_(row.values[columns.END - 1]);

  if (!startTime || !endTime) {
    summary.skipped++;
    summary.warnings.push("Row " + row.rowNumber + " was skipped because its start or end time is not a valid date.");
    return null;
  }

  if (endTime.getTime() <= startTime.getTime()) {
    summary.skipped++;
    summary.warnings.push("Row " + row.rowNumber + " was skipped because the end time is not after the start time.");
    return null;
  }

  var location = cleanString_(row.values[columns.LOCATION - 1]);
  var notes = stripSyncFooter_(cleanString_(row.values[columns.DESCRIPTION - 1]));
  var description = buildManagedDescription_(notes, row.sourceId);

  return {
    sourceId: row.sourceId,
    title: title,
    startTime: startTime,
    endTime: endTime,
    location: location,
    description: description,
    legacyKey: buildLegacyKey_(title, startTime, endTime)
  };
}

function loadManagedCalendarEvents_(calendar, desiredRows) {
  var window = buildCalendarScanWindow_(desiredRows);
  var events = withRetry_("read calendar events", function () {
    return calendar.getEvents(window.start, window.end);
  });
  var state = {
    window: window,
    eventsById: {},
    eventsBySourceId: {},
    legacyEventsByKey: {},
    managedEvents: [],
    deletedEventIds: {}
  };

  for (var i = 0; i < events.length; i++) {
    var info = buildEventInfo_(events[i]);
    if (info && info.isManaged) {
      indexEventInfo_(state, info);
    }
  }

  return state;
}

function buildCalendarScanWindow_(desiredRows) {
  var now = new Date();
  var start = addDays_(startOfDay_(now), -ADVISING_SYNC_CONFIG.CLEANUP_LOOKBACK_DAYS);
  var end = addDays_(startOfDay_(now), ADVISING_SYNC_CONFIG.CLEANUP_LOOKAHEAD_DAYS + 1);
  var padding = ADVISING_SYNC_CONFIG.CALENDAR_SCAN_PADDING_DAYS;

  for (var i = 0; i < desiredRows.length; i++) {
    var desired = desiredRows[i].desired;
    var desiredStart = addDays_(startOfDay_(desired.startTime), -padding);
    var desiredEnd = addDays_(startOfDay_(desired.endTime), padding + 1);

    if (desiredStart.getTime() < start.getTime()) {
      start = desiredStart;
    }
    if (desiredEnd.getTime() > end.getTime()) {
      end = desiredEnd;
    }
  }

  return {
    start: start,
    end: end
  };
}

function syncDesiredRows_(calendar, state, desiredRows, claimedEventIds, summary) {
  for (var i = 0; i < desiredRows.length; i++) {
    var row = desiredRows[i];
    var eventInfo = findReusableEvent_(calendar, state, row, claimedEventIds, summary);

    if (!eventInfo) {
      eventInfo = createManagedEvent_(calendar, state, row, summary);
    }

    if (!eventInfo) {
      row.eventId = "";
      continue;
    }

    var eventId = eventInfo.id || safeEventId_(eventInfo.event);
    if (eventId) {
      claimedEventIds[eventId] = true;
    }

    if (updateManagedEvent_(eventInfo.event, row.desired, row.rowNumber, summary)) {
      eventInfo = buildEventInfo_(eventInfo.event) || eventInfo;
      if (eventInfo && eventInfo.isManaged) {
        indexEventInfo_(state, eventInfo);
      }
    }

    row.eventId = safeEventId_(eventInfo.event);
    if (row.eventId) {
      claimedEventIds[row.eventId] = true;
    }
  }
}

function findReusableEvent_(calendar, state, row, claimedEventIds, summary) {
  var desired = row.desired;
  var candidate = null;

  if (row.eventId) {
    candidate = getManagedEventById_(calendar, state, row.eventId, summary);
    if (candidate && isEventClaimedOrDeleted_(state, candidate, claimedEventIds)) {
      candidate = null;
    }
  }

  if (!candidate) {
    candidate = firstReusableEvent_(state.eventsBySourceId[desired.sourceId], state, claimedEventIds);
  }

  if (!candidate) {
    candidate = firstReusableEvent_(state.legacyEventsByKey[desired.legacyKey], state, claimedEventIds);
  }

  if (candidate) {
    deleteExtraMatches_(state, state.eventsBySourceId[desired.sourceId], candidate, claimedEventIds, summary);
    deleteExtraMatches_(state, state.legacyEventsByKey[desired.legacyKey], candidate, claimedEventIds, summary);
  }

  return candidate;
}

function getManagedEventById_(calendar, state, eventId, summary) {
  if (!eventId) {
    return null;
  }

  if (state.eventsById[eventId]) {
    return state.eventsById[eventId];
  }

  var event = null;
  try {
    event = withRetry_("read calendar event by ID", function () {
      return calendar.getEventById(eventId);
    });
  } catch (error) {
    summary.warnings.push("Could not read calendar event ID " + eventId + ": " + error.message);
    return null;
  }

  if (!event) {
    return null;
  }

  var info = buildEventInfo_(event);
  if (!info || !info.isManaged) {
    summary.warnings.push("Hidden event ID " + eventId + " points to an unmanaged calendar event, so it was ignored.");
    return null;
  }

  indexEventInfo_(state, info);
  return info;
}

function createManagedEvent_(calendar, state, row, summary) {
  var desired = row.desired;

  try {
    var event = calendar.createEvent(desired.title, desired.startTime, desired.endTime, {
      location: desired.location,
      description: desired.description
    });
    var info = buildEventInfo_(event);
    if (info && info.isManaged) {
      indexEventInfo_(state, info);
    }
    summary.created++;
    return info;
  } catch (error) {
    var recovered = findEventBySourceIdAroundDesiredTime_(calendar, desired);
    if (recovered) {
      summary.warnings.push("Row " + row.rowNumber + " create returned an error, but a matching event was found and reused.");
      indexEventInfo_(state, recovered);
      return recovered;
    }

    summary.failed++;
    summary.warnings.push("Row " + row.rowNumber + " could not be created: " + error.message);
    return null;
  }
}

function updateManagedEvent_(event, desired, rowNumber, summary) {
  var changed = false;

  try {
    if (event.getTitle() !== desired.title) {
      withRetry_("update event title", function () {
        event.setTitle(desired.title);
      });
      changed = true;
    }

    if (event.getStartTime().getTime() !== desired.startTime.getTime() ||
        event.getEndTime().getTime() !== desired.endTime.getTime()) {
      withRetry_("update event time", function () {
        event.setTime(desired.startTime, desired.endTime);
      });
      changed = true;
    }

    if (cleanString_(event.getLocation()) !== desired.location) {
      withRetry_("update event location", function () {
        event.setLocation(desired.location);
      });
      changed = true;
    }

    if (cleanString_(event.getDescription()) !== desired.description) {
      withRetry_("update event description", function () {
        event.setDescription(desired.description);
      });
      changed = true;
    }
  } catch (error) {
    summary.failed++;
    summary.warnings.push("Row " + rowNumber + " could not be fully updated: " + error.message);
    return false;
  }

  if (changed) {
    summary.updated++;
  } else {
    summary.unchanged++;
  }

  return changed;
}

function removeUndesiredRowEvents_(calendar, state, rows, claimedEventIds, summary) {
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row.desired) {
      continue;
    }

    if (row.eventId) {
      var eventInfo = getManagedEventById_(calendar, state, row.eventId, summary);
      if (eventInfo && !isEventClaimedOrDeleted_(state, eventInfo, claimedEventIds)) {
        deleteEventInfo_(state, eventInfo, summary);
      }
      row.eventId = "";
    }

    if (row.sourceId) {
      var matches = state.eventsBySourceId[row.sourceId] || [];
      for (var j = 0; j < matches.length; j++) {
        if (!isEventClaimedOrDeleted_(state, matches[j], claimedEventIds)) {
          deleteEventInfo_(state, matches[j], summary);
        }
      }
    }

    if (!row.hasAnyScheduleData) {
      row.sourceId = "";
    }
  }
}

function cleanupOrphanedManagedEvents_(state, desiredLegacyKeys, claimedEventIds, summary) {
  for (var i = 0; i < state.managedEvents.length; i++) {
    var info = state.managedEvents[i];

    if (isEventClaimedOrDeleted_(state, info, claimedEventIds)) {
      continue;
    }

    if (info.sourceId) {
      deleteEventInfo_(state, info, summary);
      continue;
    }

    if (info.hasLegacyTag && !desiredLegacyKeys[info.legacyKey]) {
      deleteEventInfo_(state, info, summary);
    }
  }
}

function deleteExtraMatches_(state, matches, keepInfo, claimedEventIds, summary) {
  if (!matches) {
    return;
  }

  for (var i = 0; i < matches.length; i++) {
    var info = matches[i];
    if (sameEvent_(info, keepInfo) || isEventClaimedOrDeleted_(state, info, claimedEventIds)) {
      continue;
    }
    deleteEventInfo_(state, info, summary);
  }
}

function deleteEventInfo_(state, info, summary) {
  if (!info || !info.event) {
    return false;
  }

  var eventId = info.id || safeEventId_(info.event);
  if (eventId && state.deletedEventIds[eventId]) {
    return false;
  }

  try {
    withRetry_("delete calendar event", function () {
      info.event.deleteEvent();
    });
    if (eventId) {
      state.deletedEventIds[eventId] = true;
    }
    summary.deleted++;
    return true;
  } catch (error) {
    summary.failed++;
    summary.warnings.push("Could not delete calendar event " + (eventId || info.title) + ": " + error.message);
    return false;
  }
}

function firstReusableEvent_(matches, state, claimedEventIds) {
  if (!matches) {
    return null;
  }

  for (var i = 0; i < matches.length; i++) {
    if (!isEventClaimedOrDeleted_(state, matches[i], claimedEventIds)) {
      return matches[i];
    }
  }

  return null;
}

function isEventClaimedOrDeleted_(state, info, claimedEventIds) {
  var eventId = info && (info.id || safeEventId_(info.event));
  return !!(eventId && (claimedEventIds[eventId] || state.deletedEventIds[eventId]));
}

function sameEvent_(left, right) {
  if (!left || !right) {
    return false;
  }

  var leftId = left.id || safeEventId_(left.event);
  var rightId = right.id || safeEventId_(right.event);
  return !!(leftId && rightId && leftId === rightId);
}

function findEventBySourceIdAroundDesiredTime_(calendar, desired) {
  var start = addDays_(startOfDay_(desired.startTime), -1);
  var end = addDays_(startOfDay_(desired.endTime), 2);
  var events = [];

  try {
    Utilities.sleep(1000);
    events = calendar.getEvents(start, end);
  } catch (error) {
    return null;
  }

  for (var i = 0; i < events.length; i++) {
    var info = buildEventInfo_(events[i]);
    if (info && info.sourceId === desired.sourceId) {
      return info;
    }
  }

  return null;
}

function indexEventInfo_(state, info) {
  if (!info || !info.isManaged) {
    return;
  }

  var eventId = info.id || safeEventId_(info.event);
  if (eventId) {
    state.eventsById[eventId] = info;
  }

  if (info.sourceId) {
    if (!state.eventsBySourceId[info.sourceId]) {
      state.eventsBySourceId[info.sourceId] = [];
    }
    if (!containsEventInfo_(state.eventsBySourceId[info.sourceId], info)) {
      state.eventsBySourceId[info.sourceId].push(info);
    }
  }

  if (info.hasLegacyTag) {
    if (!state.legacyEventsByKey[info.legacyKey]) {
      state.legacyEventsByKey[info.legacyKey] = [];
    }
    if (!containsEventInfo_(state.legacyEventsByKey[info.legacyKey], info)) {
      state.legacyEventsByKey[info.legacyKey].push(info);
    }
  }

  if (!containsEventInfo_(state.managedEvents, info)) {
    state.managedEvents.push(info);
  }
}

function containsEventInfo_(list, info) {
  var eventId = info.id || safeEventId_(info.event);
  for (var i = 0; i < list.length; i++) {
    var listId = list[i].id || safeEventId_(list[i].event);
    if (eventId && listId && eventId === listId) {
      return true;
    }
  }
  return false;
}

function buildEventInfo_(event) {
  if (!event) {
    return null;
  }

  var description = cleanString_(event.getDescription());
  var sourceId = extractSourceId_(description);
  var hasManagedMarker = description.indexOf("[" + ADVISING_SYNC_CONFIG.MANAGED_MARKER + "]") !== -1;
  var hasLegacyTag = ADVISING_SYNC_CONFIG.CLEANUP_LEGACY_EVENTS &&
    description.indexOf(ADVISING_SYNC_CONFIG.LEGACY_TAG) !== -1;
  var isManaged = !!(sourceId || hasManagedMarker || hasLegacyTag);

  return {
    event: event,
    id: safeEventId_(event),
    title: cleanString_(event.getTitle()),
    startTime: event.getStartTime(),
    endTime: event.getEndTime(),
    description: description,
    sourceId: sourceId,
    hasManagedMarker: hasManagedMarker,
    hasLegacyTag: hasLegacyTag,
    isManaged: isManaged,
    legacyKey: buildLegacyKey_(event.getTitle(), event.getStartTime(), event.getEndTime())
  };
}

function writeMetadata_(sheet, metadata, rows) {
  if (!rows.length) {
    return;
  }

  var sourceIdValues = [];
  var eventIdValues = [];

  for (var i = 0; i < rows.length; i++) {
    sourceIdValues.push([rows[i].sourceId || ""]);
    eventIdValues.push([rows[i].eventId || ""]);
  }

  sheet.getRange(rows[0].rowNumber, metadata.sourceIdColumn, rows.length, 1).setValues(sourceIdValues);
  sheet.getRange(rows[0].rowNumber, metadata.eventIdColumn, rows.length, 1).setValues(eventIdValues);
}

function buildManagedDescription_(notes, sourceId) {
  var footer = "[" + ADVISING_SYNC_CONFIG.MANAGED_MARKER + "]\n" +
    "Source ID: " + sourceId + "\n" +
    ADVISING_SYNC_CONFIG.LEGACY_TAG;

  return notes ? notes + "\n\n" + footer : footer;
}

function stripSyncFooter_(value) {
  return cleanString_(value).replace(/\n*\[GSheets2Calendar\][\s\S]*$/m, "").trim();
}

function extractSourceId_(description) {
  var match = cleanString_(description).match(/\[GSheets2Calendar\][\s\S]*?Source ID:\s*([A-Za-z0-9_-]+)/);
  return match ? match[1] : "";
}

function buildLegacyKey_(title, startTime, endTime) {
  return cleanString_(title) + "|" + startTime.getTime() + "|" + endTime.getTime();
}

function coerceDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getTime());
  }

  if (typeof value === "number" && isFinite(value)) {
    return serialDateToLocalDate_(value);
  }

  var text = cleanString_(value);
  if (!text) {
    return null;
  }

  var isoMatch = text.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (isoMatch) {
    return buildValidatedDate_(
      Number(isoMatch[1]),
      Number(isoMatch[2]),
      Number(isoMatch[3]),
      Number(isoMatch[4] || 0),
      Number(isoMatch[5] || 0),
      Number(isoMatch[6] || 0)
    );
  }

  var dayFirstMatch = text.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (dayFirstMatch) {
    var year = Number(dayFirstMatch[3]);
    if (year < 100) {
      year += 2000;
    }
    return buildValidatedDate_(
      year,
      Number(dayFirstMatch[2]),
      Number(dayFirstMatch[1]),
      Number(dayFirstMatch[4] || 0),
      Number(dayFirstMatch[5] || 0),
      Number(dayFirstMatch[6] || 0)
    );
  }

  var parsed = new Date(text);
  return parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : null;
}

function serialDateToLocalDate_(serial) {
  var wholeDays = Math.floor(serial);
  var dayFraction = serial - wholeDays;
  var seconds = Math.round(dayFraction * 24 * 60 * 60);
  var base = new Date(Date.UTC(1899, 11, 30));
  var utcDate = new Date(base.getTime() + wholeDays * 24 * 60 * 60 * 1000);

  return new Date(
    utcDate.getUTCFullYear(),
    utcDate.getUTCMonth(),
    utcDate.getUTCDate(),
    Math.floor(seconds / 3600),
    Math.floor((seconds % 3600) / 60),
    seconds % 60
  );
}

function buildValidatedDate_(year, month, day, hour, minute, second) {
  var date = new Date(year, month - 1, day, hour, minute, second || 0);

  if (date.getFullYear() !== year ||
      date.getMonth() !== month - 1 ||
      date.getDate() !== day ||
      date.getHours() !== hour ||
      date.getMinutes() !== minute ||
      date.getSeconds() !== (second || 0)) {
    return null;
  }

  return date;
}

function hasAnyScheduleData_(rowValues) {
  for (var i = 0; i < ADVISING_SYNC_CONFIG.DATA_COLUMN_COUNT; i++) {
    if (rowValues[i] instanceof Date && !isNaN(rowValues[i].getTime())) {
      return true;
    }
    if (cleanString_(rowValues[i])) {
      return true;
    }
  }
  return false;
}

function isPlaceholderTitle_(title) {
  var normalized = cleanString_(title).toLowerCase();
  if (!normalized) {
    return true;
  }

  for (var i = 0; i < ADVISING_SYNC_CONFIG.PLACEHOLDER_TITLES.length; i++) {
    if (normalized === ADVISING_SYNC_CONFIG.PLACEHOLDER_TITLES[i].toLowerCase()) {
      return true;
    }
  }

  return false;
}

function cleanString_(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

function safeEventId_(event) {
  try {
    return event.getId();
  } catch (error) {
    return "";
  }
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function addDays_(date, days) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
}

function withRetry_(label, callback) {
  var lastError = null;

  for (var attempt = 1; attempt <= ADVISING_SYNC_CONFIG.MAX_RETRIES; attempt++) {
    try {
      return callback();
    } catch (error) {
      lastError = error;
      if (attempt < ADVISING_SYNC_CONFIG.MAX_RETRIES) {
        Utilities.sleep(500 * Math.pow(2, attempt - 1));
      }
    }
  }

  throw new Error(label + " failed after " + ADVISING_SYNC_CONFIG.MAX_RETRIES + " attempts: " + lastError.message);
}

function newSyncSummary_() {
  return {
    created: 0,
    updated: 0,
    unchanged: 0,
    deleted: 0,
    skipped: 0,
    failed: 0,
    warnings: []
  };
}

function reportSummary_(summary) {
  var message = "Calendar sync: " +
    summary.created + " created, " +
    summary.updated + " updated, " +
    summary.unchanged + " unchanged, " +
    summary.deleted + " deleted";

  if (summary.skipped) {
    message += ", " + summary.skipped + " skipped";
  }
  if (summary.failed) {
    message += ", " + summary.failed + " failed";
  }
  if (summary.warnings.length) {
    message += ", " + summary.warnings.length + " warnings";
  }

  Logger.log(message);
  for (var i = 0; i < summary.warnings.length; i++) {
    Logger.log(summary.warnings[i]);
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Sync to Calendar", 8);
  } catch (error) {
    Logger.log("Could not show sync toast: " + error.message);
  }
}
