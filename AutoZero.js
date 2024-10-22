function updatePreviousDayRow() {
  const FIRST_ROW = 6; // The first row with data
  const SPREADSHEET_ID = "1op3uW3K-6i5ANWouEFrl9-HnOZBuLO7opq124-nRfhs";
  const LOG_SHEET_NAME = "Anki Log";
  const DATE_COLUMN = "C";
  const COLUMNS_TO_SET = ["G", "I", "K", "M", "W", "X"]; // Columns you want to set to 0
  
  var logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
  
  var today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize to midnight
  
  // Get yesterday's date
  var yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1); // Go back one day

  var lastRow = logSheet.getLastRow();
  var rowUpdated = false; // Flag to track if a row was updated

  for (var i = FIRST_ROW; i <= lastRow; i++) {
    var cellDate = new Date(logSheet.getRange(DATE_COLUMN + i).getValue());
    cellDate.setHours(0, 0, 0, 0);  // Normalize the date for comparison

    if (cellDate.getTime() === yesterday.getTime()) {
      // Edge case: If you're on the first row (6), don't update the previous row
      if (i === FIRST_ROW) {
        Logger.log("First row, no previous row to update.");
        break;
      }

      // Set the corresponding cells in the specified columns to 0 if they are empty
      setValuesIfEmpty(logSheet, COLUMNS_TO_SET, i, 0);

      rowUpdated = true; // Set the flag to true
      break;  // Stop after setting the matching row for yesterday
    }
  }

  // Log the result
  if (rowUpdated) {
    Logger.log("Columns updated for yesterday's date and column R updated.");
  } else {
    Logger.log("No matching date found for yesterday.");
  }
}

function setValuesIfEmpty(sheet, columns, rowIndex, value) {
  // Set the corresponding cells in the specified columns to the given value if they are empty
  columns.forEach(function(column) {
    var cell = sheet.getRange(column + rowIndex);
    if (cell.getValue() === "") { // Check if the cell is empty
      cell.setValue(value);
    }
  });
}