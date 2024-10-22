// Helper function for modularity's sake
function updateBackgroundColorForSheet(sheetName, firstRow, dateColumn, color) {
  const SPREADSHEET_ID = "1op3uW3K-6i5ANWouEFrl9-HnOZBuLO7opq124-nRfhs"; 
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize to midnight

  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1); // Get yesterday's date

  const lastRow = sheet.getLastRow();
  
  // Store yesterday's row index
  let yesterdayRow = -1;

  for (let i = firstRow; i <= lastRow; i++) {
    const cellDate = new Date(sheet.getRange(dateColumn + i).getValue());
    cellDate.setHours(0, 0, 0, 0); // Normalize the date for comparison

    if (cellDate.getTime() === yesterday.getTime()) {
      // Store the row index for yesterday to reset its background color later
      yesterdayRow = i;
    } else if (cellDate.getTime() === today.getTime()) {
      // Update background color for today's row
      sheet.getRange("A" + i + ":C" + i).setBackground(color);
    }
  }

  // Reset the background color for yesterday's row if it was found
  if (yesterdayRow !== -1) {
    sheet.getRange("A" + yesterdayRow + ":C" + yesterdayRow).setBackground(null); // Reset to default
  }
}

function updateBackgroundColor() {
  const GOLDEN_YELLOW = '#FFD700';
  
  // Update "Anki Log" sheet (dates start at row 6)
  updateBackgroundColorForSheet("Anki Log", 6, "C", GOLDEN_YELLOW);
  
  // Update "Immersion" sheet (dates start at row 5)
  updateBackgroundColorForSheet("Immersion", 5, "C", GOLDEN_YELLOW);

  // Update "Other Log" sheet (dates start at row 5)
  updateBackgroundColorForSheet("Other Log", 4, "C", GOLDEN_YELLOW);
}
