/***************************************************************
 * DAILY SCRUM AUTOMATION - BLANK TOP SPACER VERSION
 *
 * Layout:
 * Row 1       : Headers
 * Row 2       : Completely blank spacer (A:G)
 * Rows 3-12   : Team members (detected dynamically from Column B)
 * Row 13      : Existing blank spacer between date groups
 * Row 14+     : Previous Scrum groups
 *
 * Important:
 * - The top spacer is NOT part of the date merge.
 * - The date is merged only across team-member rows.
 * - No extra team-member row is added.
 * - The existing top spacer moves down and becomes the separator
 *   between the newly created block and the previous block.
 ***************************************************************/


// =============================================================
// 1. CONFIGURATION
// =============================================================

const CONFIG = {

  // false = actual current date in India
  // true  = use DEMO_DATE
  DEMO_MODE: false,
  DEMO_DATE: "2026-11-11",

  TIMEZONE: "Asia/Kolkata",

  HEADER_ROW: 1,
  FIRST_DATA_ROW: 2,

  DATE_COLUMN: 1,       // A
  TEAM_COLUMN: 2,       // B
  YESTERDAY_COLUMN: 3,  // C
  TODAY_COLUMN: 4,      // D
  TOTAL_COLUMNS: 7,     // A:G

  // One completely blank spacer above the newest block.
  TOP_SPACER_ROWS: 1,
  SPACER_HEIGHT: 18,

  TRIGGER_HOUR: 2,
  MAX_CATCH_UP_DAYS: 370,

  // Holiday dates: YYYY-MM-DD
  HOLIDAYS: [

    // ================= 2026 =================
    "2026-01-01", // New Year's Day
    "2026-01-26", // Republic Day
    "2026-03-04", // Holi
    "2026-04-03", // Good Friday
    "2026-06-15", // Raja Sankranti
    "2026-06-26", // Muharram
    "2026-09-14", // Ganesh Puja
    "2026-10-02", // Gandhi Jayanti
    "2026-10-20", // Dussehra
    "2026-12-25", // Christmas

    // ================= 2027 =================
    "2027-01-01",
    "2027-01-26",
    "2027-03-23",
    "2027-03-26",
    "2027-06-15",
    "2027-06-16",
    "2027-09-04",
    "2027-10-02",
    "2027-10-09",
    "2027-12-25",

    // ================= 2028 =================
    "2028-01-01",
    "2028-01-26",
    "2028-03-11",
    "2028-04-14",
    "2028-06-15",
    "2028-06-04",
    "2028-08-23",
    "2028-10-02",
    "2028-09-27",
    "2028-12-25",

    // ================= 2029 =================
    "2029-01-01",
    "2029-01-26",
    "2029-03-01",
    "2029-03-30",
    "2029-06-15",
    "2029-05-25",
    "2029-09-11",
    "2029-10-02",
    "2029-10-16",
    "2029-12-25",

    // ================= 2030 =================
    "2030-01-01",
    "2030-01-26",
    "2030-03-20",
    "2030-04-19",
    "2030-06-15",
    "2030-05-14",
    "2030-09-01",
    "2030-10-02",
    "2030-10-06",
    "2030-12-25"
  ]
};


// =============================================================
// 2. ON OPEN
// =============================================================

function onOpen(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) runScrumSafely_(ss);
}


// =============================================================
// 3. DAILY TRIGGER / MANUAL ENTRY POINT
// =============================================================

function createDailyScrumBlocks() {
  try {
    const spreadsheetId = PropertiesService
      .getScriptProperties()
      .getProperty("SCRUM_SPREADSHEET_ID");

    const ss = spreadsheetId
      ? SpreadsheetApp.openById(spreadsheetId)
      : SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) throw new Error("Spreadsheet could not be opened.");

    runScrumSafely_(ss);

  } catch (error) {
    Logger.log("SPREADSHEET ERROR: " + error.message);
  }
}


// =============================================================
// 4. SAFE EXECUTION
// =============================================================

function runScrumSafely_(ss) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log("Another Daily Scrum execution is already running.");
    return;
  }

  try {
    processDailyScrum_(ss);
  } catch (error) {
    Logger.log("DAILY SCRUM ERROR: " + error.message);
    if (error.stack) Logger.log(error.stack);

    try {
      ss.toast(error.message, "Daily Scrum Automation", 10);
    } catch (toastError) {
      Logger.log("Could not display spreadsheet notification.");
    }
  } finally {
    lock.releaseLock();
  }
}


// =============================================================
// 5. LEFTMOST VISIBLE TAB
// =============================================================

function getLatestSheet_(ss) {
  const visibleSheets = ss.getSheets().filter(
    sheet => !sheet.isSheetHidden()
  );

  if (visibleSheets.length === 0) {
    throw new Error("No visible sheet tabs were found.");
  }

  const sheet = visibleSheets[0];
  Logger.log("Selected latest tab: " + sheet.getName());
  return sheet;
}


// =============================================================
// 6. MAIN PROCESS
// =============================================================

function processDailyScrum_(ss) {
  const sheet = getLatestSheet_(ss);

  if (sheet.getMaxColumns() < CONFIG.TOTAL_COLUMNS) {
    throw new Error("The selected tab must contain columns A:G.");
  }

  const topBlock = getTopBlock_(sheet);
  const latestKey = topBlock.dateKey;

  const targetKey = CONFIG.DEMO_MODE
    ? validateDateKey_(CONFIG.DEMO_DATE)
    : Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd");

  Logger.log("Detected team rows: " + topBlock.teamCount);
  Logger.log("Detected top spacer rows: " + CONFIG.TOP_SPACER_ROWS);
  Logger.log("Latest date in sheet: " + latestKey);
  Logger.log("Target date in India: " + targetKey);
  Logger.log("Demo mode: " + CONFIG.DEMO_MODE);

  if (targetKey <= latestKey) {
    Logger.log("No new Scrum section required.");
    return;
  }

  const latestParts = latestKey.split("-").map(Number);
  const targetParts = targetKey.split("-").map(Number);

  let current = new Date(Date.UTC(
    latestParts[0], latestParts[1] - 1, latestParts[2]
  ));

  const targetUTC = Date.UTC(
    targetParts[0], targetParts[1] - 1, targetParts[2]
  );

  const daysDifference = Math.floor(
    (targetUTC - current.getTime()) / 86400000
  );

  if (daysDifference > CONFIG.MAX_CATCH_UP_DAYS) {
    throw new Error(
      "Target date is more than " + CONFIG.MAX_CATCH_UP_DAYS +
      " days ahead of the latest sheet date. Check the date and configuration."
    );
  }

  const holidays = new Set(CONFIG.HOLIDAYS);
  let createdCount = 0;

  current.setUTCDate(current.getUTCDate() + 1);

  while (current.getTime() <= targetUTC) {
    const dateKey = Utilities.formatDate(current, "UTC", "yyyy-MM-dd");
    const weekday = current.getUTCDay();
    const isWeekend = weekday === 0 || weekday === 6;
    const isHoliday = holidays.has(dateKey);

    if (!isWeekend && !isHoliday) {
      insertScrumBlock_(sheet, topBlock.teamCount, dateKey);
      createdCount++;
      Logger.log("Created Scrum block: " + dateKey);
    } else {
      Logger.log("Skipped non-working day: " + dateKey);
    }

    current.setUTCDate(current.getUTCDate() + 1);
  }

  Logger.log("Scrum automation completed. Blocks created: " + createdCount);

  try {
    ss.toast(
      createdCount + " new Scrum section(s) created.",
      "Daily Scrum Automation",
      5
    );
  } catch (error) {
    Logger.log("Completed; notification was unavailable.");
  }
}


// =============================================================
// 7. DETECT TOP BLOCK AND NORMALIZE ITS TOP SPACER
// =============================================================

function getTopBlock_(sheet) {
  const firstRow = CONFIG.FIRST_DATA_ROW;
  const lastRow = sheet.getLastRow();

  if (lastRow < firstRow) {
    throw new Error("No Scrum data was found below the header.");
  }

  const firstTwo = sheet.getRange(
    firstRow,
    CONFIG.TEAM_COLUMN,
    Math.min(2, lastRow - firstRow + 1),
    1
  ).getDisplayValues();

  const b2 = String((firstTwo[0] || [""])[0]).trim();
  const b3 = String((firstTwo[1] || [""])[0]).trim();

  let memberStart;
  let needsTopSpacer = false;

  if (b2 === "" && b3 !== "") {
    // Top spacer already exists.
    memberStart = firstRow + CONFIG.TOP_SPACER_ROWS;
  } else if (b2 !== "") {
    // Older layout has no top spacer. We will add one after
    // reading the current date and team rows.
    memberStart = firstRow;
    needsTopSpacer = true;
  } else {
    throw new Error(
      "Could not find the first team row. Check Column B rows 2 and 3."
    );
  }

  const availableRows = lastRow - memberStart + 1;
  const memberValues = sheet.getRange(
    memberStart,
    CONFIG.TEAM_COLUMN,
    availableRows,
    1
  ).getDisplayValues();

  let teamCount = 0;
  for (const row of memberValues) {
    if (String(row[0]).trim() === "") break;
    teamCount++;
  }

  if (teamCount === 0) {
    throw new Error("No consecutive team-member rows were found in Column B.");
  }

  // Ensure the first blank after the team list is a separator.
  const firstBlankRow = memberStart + teamCount;
  const rowAfterBlank = firstBlankRow + 1;

  if (rowAfterBlank <= lastRow) {
    const nextName = sheet.getRange(
      rowAfterBlank,
      CONFIG.TEAM_COLUMN
    ).getDisplayValue().trim();

    if (nextName !== "") {
      const nextDate = parseDateCell_(
        sheet.getRange(rowAfterBlank, CONFIG.DATE_COLUMN)
      );

      if (!nextDate) {
        throw new Error(
          "A blank row appears inside the team list near row " +
          firstBlankRow + ". Check the team rows before running."
        );
      }
    }
  }

  // Read date before changing any merged cells. The older layout may
  // have the date merged from A2 through the team rows.
  const dateSearchStart = firstRow;
  const dateSearchRows = memberStart + teamCount - firstRow;
  const dateRangeToRead = sheet.getRange(
    dateSearchStart,
    CONFIG.DATE_COLUMN,
    dateSearchRows,
    1
  );

  const values = dateRangeToRead.getValues();
  const displays = dateRangeToRead.getDisplayValues();
  const dates = [];

  for (let i = 0; i < dateSearchRows; i++) {
    const parsed = parseDateValue_(values[i][0], displays[i][0]);
    if (parsed) dates.push(parsed);
  }

  const uniqueKeys = [...new Set(dates.map(item => item.key))];

  if (uniqueKeys.length === 0) {
    throw new Error(
      "No date found in the top Scrum block in Column A. " +
      "Please restore the latest date; no new rows were inserted."
    );
  }

  if (uniqueKeys.length > 1) {
    throw new Error(
      "Multiple dates found in the top Scrum block. Check Column A."
    );
  }

  const dateKey = uniqueKeys[0];

  // If no top spacer existed, insert a separate blank spacer row.
  // The date value was already captured before changing the layout.
  if (needsTopSpacer) {
    sheet.insertRowsBefore(firstRow, CONFIG.TOP_SPACER_ROWS);
    memberStart += CONFIG.TOP_SPACER_ROWS;
  }

  // Make the top spacer completely blank and keep the date merge
  // ONLY across team-member rows, not across the spacer.
  normalizeTopSpacerAndDate_(sheet, memberStart, teamCount, dateKey);

  return {
    teamCount: teamCount,
    memberStart: memberStart,
    dateKey: dateKey
  };
}


// =============================================================
// 8. NORMALIZE TOP SPACER + DATE MERGE
// =============================================================

function normalizeTopSpacerAndDate_(sheet, memberStart, teamCount, dateKey) {
  const firstRow = CONFIG.FIRST_DATA_ROW;
  const totalColumns = CONFIG.TOTAL_COLUMNS;
  const spacerRow = firstRow;
  const bottomOfTeam = memberStart + teamCount - 1;
  const oldBottomSpacerRow = bottomOfTeam + 1;

  // Use the existing blank spacer after the current team block as
  // the formatting source for the new top spacer, when available.
  if (oldBottomSpacerRow <= sheet.getMaxRows() &&
      isCompletelyBlankRow_(sheet, oldBottomSpacerRow)) {
    sheet.getRange(oldBottomSpacerRow, 1, 1, totalColumns).copyTo(
      sheet.getRange(spacerRow, 1, 1, totalColumns),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
  }

  // Break any old date merge that included the spacer row.
  sheet.getRange(
    firstRow,
    CONFIG.DATE_COLUMN,
    memberStart + teamCount - firstRow,
    1
  ).breakApart();

  // The spacer must be blank across A:G.
  sheet.getRange(
    spacerRow,
    1,
    1,
    totalColumns
  ).clearContent();

  sheet.setRowHeight(spacerRow, CONFIG.SPACER_HEIGHT);

  // Merge the date ONLY over the team-member rows.
  const dateRange = sheet.getRange(
    memberStart,
    CONFIG.DATE_COLUMN,
    teamCount,
    1
  );

  dateRange.merge();
  writeDateFormula_(sheet, memberStart, dateKey, teamCount);
}


// =============================================================
// 9. CHECK WHETHER A ROW IS COMPLETELY BLANK
// =============================================================

function isCompletelyBlankRow_(sheet, row) {
  const range = sheet.getRange(
    row,
    1,
    1,
    CONFIG.TOTAL_COLUMNS
  );

  const values = range.getDisplayValues()[0];
  const formulas = range.getFormulas()[0];

  return values.every(value => String(value).trim() === "") &&
    formulas.every(formula => String(formula).trim() === "");
}


// =============================================================
// 10. PARSE DATE
// =============================================================

function parseDateCell_(cell) {
  return parseDateValue_(cell.getValue(), cell.getDisplayValue());
}

function parseDateValue_(value, displayValue) {
  const text = String(displayValue || "").trim();

  if (text !== "") {
    const formats = [
      "d MMM yyyy",
      "dd MMM yyyy",
      "d MMMM yyyy",
      "dd MMMM yyyy",
      "yyyy-MM-dd",
      "MMM d, yyyy",
      "MMMM d, yyyy"
    ];

    for (const format of formats) {
      try {
        const parsed = Utilities.parseDate(text, CONFIG.TIMEZONE, format);
        if (parsed instanceof Date && !isNaN(parsed.getTime())) {
          return {
            date: parsed,
            key: Utilities.formatDate(parsed, CONFIG.TIMEZONE, "yyyy-MM-dd")
          };
        }
      } catch (error) {
        // Try next supported format.
      }
    }
  }

  if (value instanceof Date && !isNaN(value.getTime())) {
    return {
      date: value,
      key: Utilities.formatDate(value, CONFIG.TIMEZONE, "yyyy-MM-dd")
    };
  }

  return null;
}


// =============================================================
// 11. VALIDATE DATE KEY
// =============================================================

function validateDateKey_(dateKey) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dateKey));
  if (!match) throw new Error("DEMO_DATE must use YYYY-MM-DD format.");

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));

  if (date.getUTCFullYear() !== year ||
      date.getUTCMonth() !== month - 1 ||
      date.getUTCDate() !== day) {
    throw new Error("DEMO_DATE is not a valid calendar date.");
  }

  return String(dateKey);
}


// =============================================================
// 12. WRITE DATE FORMULA
// =============================================================


function writeDateFormula_(sheet, firstRow, dateKey, sectionRows) {

  const [year, month, day] = dateKey.split("-").map(Number);

  const dateRange = sheet.getRange(
    firstRow,
    CONFIG.DATE_COLUMN,
    sectionRows,
    1
  );

  const dateCell = sheet.getRange(
    firstRow,
    CONFIG.DATE_COLUMN
  );

  // Store an actual date value, not a formula.
  const dateValue = new Date(
    Date.UTC(year, month - 1, day, 12, 0, 0)
  );

  dateCell.setValue(dateValue);

  // Display format
  dateRange.setNumberFormat("d MMM yyyy");

  // Enable calendar date picker
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();

  dateCell.setDataValidation(dateRule);

  dateRange
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}



// =============================================================
// 13. INSERT NEW SCRUM BLOCK WITH A BLANK SPACER ABOVE
// =============================================================

function insertScrumBlock_(sheet, teamCount, dateKey) {
  const firstRow = CONFIG.FIRST_DATA_ROW;
  const totalColumns = CONFIG.TOTAL_COLUMNS;
  const spacerRows = CONFIG.TOP_SPACER_ROWS;

  // Current top block is normalized before this function runs:
  // row 2 is a blank spacer; team members begin on row 3.
  const currentMemberStart = firstRow + spacerRows;

  // Read names BEFORE inserting rows.
  const names = sheet.getRange(
    currentMemberStart,
    CONFIG.TEAM_COLUMN,
    teamCount,
    1
  ).getValues();

  if (names.some(row => String(row[0] || "").trim() === "")) {
    throw new Error(
      "A blank team-member cell was detected in the existing team block. " +
      "No rows were inserted."
    );
  }

  // Insert only the top spacer + member rows.
  // The old top spacer shifts down and becomes the separator between blocks.
  const rowsToInsert = spacerRows + teamCount;
  const oldSpacerRowAfterInsert = firstRow + rowsToInsert;
  const oldMemberStart = currentMemberStart + rowsToInsert;
  let rowsInserted = false;

  try {
    sheet.insertRowsBefore(firstRow, rowsToInsert);
    rowsInserted = true;

    const newSpacer = sheet.getRange(
      firstRow,
      1,
      spacerRows,
      totalColumns
    );

    const newMembers = sheet.getRange(
      currentMemberStart,
      1,
      teamCount,
      totalColumns
    );

    const oldMembers = sheet.getRange(
      oldMemberStart,
      1,
      teamCount,
      totalColumns
    );

    // Copy the existing spacer formatting to the new top spacer.
    sheet.getRange(
      oldSpacerRowAfterInsert,
      1,
      1,
      totalColumns
    ).copyTo(
      newSpacer,
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

    newSpacer.breakApart();
    newSpacer.clearContent();
    sheet.setRowHeight(firstRow, CONFIG.SPACER_HEIGHT);

    // Copy formatting from previous member block.
    oldMembers.copyTo(
      newMembers,
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

    newMembers.breakApart();

    for (let i = 0; i < teamCount; i++) {
      sheet.setRowHeight(
        currentMemberStart + i,
        sheet.getRowHeight(oldMemberStart + i)
      );
    }

    // Clear new member content and restore the names.
    newMembers.clearContent();

    sheet.getRange(
      currentMemberStart,
      CONFIG.TEAM_COLUMN,
      teamCount,
      1
    ).setValues(names);

    // Previous date's Column D -> new date's Column C.
    // Copy each member row so rich-text hyperlinks remain clickable.
    for (let i = 0; i < teamCount; i++) {
      const oldRow = oldMemberStart + i;
      const newRow = currentMemberStart + i;

      const oldToday = sheet.getRange(oldRow, CONFIG.TODAY_COLUMN);
      const newYesterday = sheet.getRange(newRow, CONFIG.YESTERDAY_COLUMN);
      const oldYesterdayFormat = sheet.getRange(oldRow, CONFIG.YESTERDAY_COLUMN);

      oldToday.copyTo(
        newYesterday,
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

      oldYesterdayFormat.copyTo(
        newYesterday,
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false
      );
    }

    // New Today column is blank.
    sheet.getRange(
      currentMemberStart,
      CONFIG.TODAY_COLUMN,
      teamCount,
      1
    ).clearContent();

    // Date merge covers ONLY team members, never the spacer row.
    const newDateRange = sheet.getRange(
      currentMemberStart,
      CONFIG.DATE_COLUMN,
      teamCount,
      1
    );

    newDateRange.breakApart();
    newDateRange.merge();
    writeDateFormula_(sheet, currentMemberStart, dateKey, teamCount);

    // The shifted old spacer row is now the blank separator after this block.
    // No additional bottom spacer row is inserted, avoiding double gaps.

  } catch (error) {
    if (rowsInserted) {
      try {
        sheet.deleteRows(firstRow, rowsToInsert);
        Logger.log("Incomplete insertion was rolled back.");
      } catch (rollbackError) {
        Logger.log("ROLLBACK ERROR: " + rollbackError.message);
      }
    }

    throw error;
  }
}


// =============================================================
// 14. INSTALL DAILY TRIGGER
// =============================================================

function installDailyTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("Open the target Google Sheet before installing the trigger.");
  }

  PropertiesService.getScriptProperties().setProperty(
    "SCRUM_SPREADSHEET_ID",
    ss.getId()
  );

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "createDailyScrumBlocks") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("createDailyScrumBlocks")
    .timeBased()
    .atHour(CONFIG.TRIGGER_HOUR)
    .everyDays(1)
    .inTimezone(CONFIG.TIMEZONE)
    .create();

  ss.toast(
    "Daily Scrum trigger installed.",
    "Daily Scrum Automation",
    5
  );

  Logger.log(
    "Daily trigger installed for " + CONFIG.TIMEZONE +
    " at approximately hour " + CONFIG.TRIGGER_HOUR + "."
  );
}


// =============================================================
// 15. MANUAL RUN
// =============================================================

function runScrumManually() {
  createDailyScrumBlocks();
}
