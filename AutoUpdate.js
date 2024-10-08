function updatePreviousDayRow() {
  const FIRST_ROW = 6; // The first row with data
  const SPREADSHEET_ID = "1op3uW3K-6i5ANWouEFrl9-HnOZBuLO7opq124-nRfhs";
  var logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Learning Log");
  var graphDataSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Graph Data");
  var dateColumn = "C";
  var columnsToSet = ["E", "G", "I", "K", "T", "U"]; // Columns you want to set to 0
  var rateColumn = "P"; // Column where the recommended rate is recorded
  var today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize to midnight
  
  // Get yesterday's date
  var yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1); // Go back one day

  // Get the value of the named range "recommended_rate" from the Graph Data sheet
  var recommendedRate = graphDataSheet.getRange("recommended_rate").getValue();

  var lastRow = logSheet.getLastRow();
  var rowUpdated = false; // Flag to track if a row was updated

  for (var i = FIRST_ROW; i <= lastRow; i++) {
    var cellDate = new Date(logSheet.getRange(dateColumn + i).getValue());
    cellDate.setHours(0, 0, 0, 0);  // Normalize the date for comparison

    if (cellDate.getTime() === yesterday.getTime()) {
      // Edge case: If you're on the first row (6), don't update the previous row
      if (i === FIRST_ROW) {
        Logger.log("First row, no previous row to update.");
        break;
      }

      // Set the corresponding cells in the specified columns to 0 if they are empty
      setValuesIfEmpty(logSheet, columnsToSet, i, 0);

      // Set the corresponding cell in the recommended rate column to the recommended rate if it is empty
      setValuesIfEmpty(logSheet, [rateColumn], i, recommendedRate);

      // Update column R: R = Q - recommended_rate if column R is empty
      setCalculatedValueIfEmpty(logSheet, "R", "Q", i, recommendedRate);


      rowUpdated = true; // Set the flag to true
      break;  // Stop after setting the matching row for yesterday
    }
  }

  // Log the result
  if (rowUpdated) {
    Logger.log("Columns updated for yesterday's date and column O updated.");
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

function setCalculatedValueIfEmpty(sheet, targetColumn, sourceColumn, rowIndex, rate) {
  var targetCell = sheet.getRange(targetColumn + rowIndex);
  if (targetCell.getValue() === "") { // Check if the target cell is empty
    var valueInSource = sheet.getRange(sourceColumn + rowIndex).getValue(); // Get the value from the source column
    var newValue = valueInSource - rate; // Calculate the new value
    targetCell.setValue(newValue); // Set the new value in the target cell
  }
}



function updateBackgroundColor() {
  var spreadsheetId = "1op3uW3K-6i5ANWouEFrl9-HnOZBuLO7opq124-nRfhs"; // Your spreadsheet ID
  var sheetName = "Learning Log"; // Your sheet name
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var dateColumn = "C"; // Column where dates are stored
  var today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize to midnight

  var yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1); // Get yesterday's date

  var lastRow = sheet.getLastRow();
  var goldenYellow = '#FFD700'; // Golden yellow color code

  // Store yesterday's row index
  var yesterdayRow = -1;

  for (var i = FIRST_ROW; i <= lastRow; i++) {
    var cellDate = new Date(sheet.getRange(dateColumn + i).getValue());
    cellDate.setHours(0, 0, 0, 0); // Normalize the date for comparison

    if (cellDate.getTime() === yesterday.getTime()) {
      // Store the row index for yesterday to reset its background color later
      yesterdayRow = i;
    } else if (cellDate.getTime() === today.getTime()) {
      // Update background color for today's row
      sheet.getRange("A" + i + ":C" + i).setBackground(goldenYellow);
    }
  }

  // Reset the background color for yesterday's row if it was found
  if (yesterdayRow !== -1) {
    sheet.getRange("A" + yesterdayRow + ":C" + yesterdayRow).setBackground(null); // Reset to default
  }
}
