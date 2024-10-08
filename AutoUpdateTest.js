function updatePreviousDayRowTest() {
    const FIRST_ROW = 6; // The first row with data
    const SPREADSHEET_ID = "1op3uW3K-6i5ANWouEFrl9-HnOZBuLO7opq124-nRfhs";
    const LOG_SHEET_NAME = "Learning Log";
    const GRAPH_DATA_SHEET_NAME = "Graph Data";
    const DATE_COLUMN = "C";
    const COLUMNS_TO_SET = ["E", "G", "I", "K", "U", "V"]; // Columns you want to set to 0
    const RATE_COLUMN = "P"; // Column where the recommended rate is recorded
    const TARGET_COLUMN = "R"; // Target column for calculated value
    const SOURCE_COLUMN = "Q"; // Source column for calculation
    
    var logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
    var graphDataSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(GRAPH_DATA_SHEET_NAME);
    
    var today = new Date();
    today.setHours(0, 0, 0, 0); // Normalize to midnight
    
    // Get the value of the named range "recommended_rate" from the Graph Data sheet
    var recommendedRate = graphDataSheet.getRange("recommended_rate").getValue();
  
    var lastRow = logSheet.getLastRow();
    var rowUpdated = false; // Flag to track if a row was updated
  
    // Check for today's date in the log sheet
    for (var i = FIRST_ROW; i <= lastRow; i++) {
      var cellDate = new Date(logSheet.getRange(DATE_COLUMN + i).getValue());
      cellDate.setHours(0, 0, 0, 0);  // Normalize the date for comparison
  
      if (cellDate.getTime() === today.getTime()) {
        // Update today's row
        // Set the corresponding cells in the specified columns to 0 if they are empty
        setValuesIfEmpty(logSheet, COLUMNS_TO_SET, i, 0);
  
        // Set the corresponding cell in the recommended rate column to the recommended rate if it is empty
        setValuesIfEmpty(logSheet, [RATE_COLUMN], i, recommendedRate);
  
        rowUpdated = true; // Set the flag to true
        break;  // Stop after setting the matching row for today
      }
    }
  
    // Log the result
    if (rowUpdated) {
      Logger.log("Today's row updated with values.");
    } else {
      Logger.log("No matching date found for today.");
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
  
  function updateBackgroundColorTest() {
    const FIRST_ROW = 6; // The first row with data
    const SPREADSHEET_ID = "1op3uW3K-6i5ANWouEFrl9-HnOZBuLO7opq124-nRfhs"; // Your spreadsheet ID
    const SHEET_NAME = "Learning Log"; // Your sheet name
    const DATE_COLUMN = "C"; // Column where dates are stored
    const GOLDEN_YELLOW = '#FFD700'; // Golden yellow color code
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var today = new Date();
    today.setHours(0, 0, 0, 0); // Normalize to midnight
  
    var yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1); // Get yesterday's date
  
    var lastRow = sheet.getLastRow();
    
    // Store yesterday's row index
    var yesterdayRow = -1;
  
    for (var i = FIRST_ROW; i <= lastRow; i++) {
      var cellDate = new Date(sheet.getRange(DATE_COLUMN + i).getValue());
      cellDate.setHours(0, 0, 0, 0); // Normalize the date for comparison
  
      if (cellDate.getTime() === yesterday.getTime()) {
        // Store the row index for yesterday to reset its background color later
        yesterdayRow = i;
      } else if (cellDate.getTime() === today.getTime()) {
        // Update background color for today's row
        sheet.getRange("A" + i + ":C" + i).setBackground(GOLDEN_YELLOW);
      }
    }
  
    // Reset the background color for yesterday's row if it was found
    if (yesterdayRow !== -1) {
      sheet.getRange("A" + yesterdayRow + ":C" + yesterdayRow).setBackground(null); // Reset to default
    }
  }