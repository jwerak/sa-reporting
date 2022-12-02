const MILLIS_PER_HOUR = 1000 * 60 * 60;
const MILLIS_PER_WORK_WEEK = MILLIS_PER_HOUR * 24 * 5;
const DATE_FORMATS = {
  "Date DD/MM/YYYY": "d/M/y",
  "Date DD.MM.YYYY": "d.M.y",
  "Date MM/DD/YYYY": "M/d/y",
};

function importCalendars() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SpreadsheetApp.getActiveSheet();
  const cal = CalendarApp.getCalendarById(sheet.getRange("C1").getValue());
  const currentTimezone = spreadsheet.getSpreadsheetTimeZone();
  const date_format = spreadsheet.getRange("D1").getValue();
  const currentFormat = DATE_FORMATS[date_format];

  // Calculate start & end dates
  const start_date = sheet.getRange("A1").getValue();
  const end_date = new Date(start_date.getTime() + MILLIS_PER_WORK_WEEK);

  cleanUpSheet(sheet);

  // Get and filter events
  const events = cal.getEvents(start_date, end_date).filter(function (e) {
    return [
      CalendarApp.GuestStatus.OWNER,
      CalendarApp.GuestStatus.YES,
    ].includes(e.getMyStatus());
  });

  const lastWeekEvents: string[][] = [];
  events.forEach((event: GoogleAppsScript.Calendar.CalendarEvent) => {
    const date = event.getStartTime();
    let hours =
      (event.getEndTime().getTime() - event.getStartTime().getTime()) /
      MILLIS_PER_HOUR;

    // Normalize full day events
    if (hours % 24 == 0) {
      hours = hours / 3;
    }

    const formatted_date = Utilities.formatDate(
      date,
      currentTimezone,
      currentFormat
    );

    lastWeekEvents.push([
      event.getTitle(),
      formatted_date,
      String(hours.toFixed(2)),
    ]);
  });

  sheet.getRange(2, 3, lastWeekEvents.length, 3).setValues(lastWeekEvents);
}

function cleanUpSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  sheet.getRange("A2:Z").clearContent();
  sheet.getRange("A2:Z").setVerticalAlignment("top");
  sheet.getRange("A2:Z").setHorizontalAlignment("left");
  sheet.getRange("G2:H").setHorizontalAlignment("right");
  sheet.getRange("D2:E").setHorizontalAlignment("right");
}

function onEdit(e) {
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (row === 1 && col === 1) {
    //Set the sheet name
    updateSheetName(sheet);
    return;
  }

  if (col == 6 && row > 1) {
    const lookup_str = range.getValue();
    if (lookup_str == "") {
      // Cleanup line
      sheet.getRange(row, 9, 1, 6).setValues([["", "", "", "", "", ""]]);
      return;
    }

    const [oppname, account, product] = lookup_str.split(" | ");
    sheet.getRange(row, 9).setValue("Opportunity");
    sheet.getRange(row, 11).setValue(product);
    sheet.getRange(row, 13).setValue(account);
    sheet.getRange(row, 14).setValue(oppname);
  }
}

function updateSheetName(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const newSheetName = Utilities.formatDate(
    sheet.getRange(1, 1).getValue(),
    "GMT+11",
    "yyyy-MM-dd"
  );

  sheet.setName(newSheetName);
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    { name: "Load Calendar Events", functionName: "importCalendars" },
  ];
  ss.addMenu("SA Reporting", menuEntries);
}
