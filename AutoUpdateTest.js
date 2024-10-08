// function testZeroColumnsAndUpdateO() {
//   var spreadsheetId = "1op3uW3K-6i5ANWouEFrl9-HnOZBuLO7opq124-nRfhs";
//   var sheetName = "Learning Log";
//   var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
//   var dateColumn = "C";  // Column where dates are stored
//   var columnsToSet = ["D", "F", "J", "L", "Q", "R"]; // Columns you want to set to 0
//   var today = new Date();
//   today.setHours(0, 0, 0, 0); // Normalize to midnight
  
//   // Get the value of the named range "recommended_rate"
//   var recommendedRateStr = sheet.getRange("recommended_rate").getValue();
  
//   // Extract the numeric value from the string
//   var recommendedRate = parseInt(recommendedRateStr.split(" ")[0], 10); // Get the first part and convert to an integer

//   var lastRow = sheet.getLastRow();
//   var rowUpdated = false; // Flag to track if a row was updated

//   for (var i = 15; i <= lastRow; i++) {
//     var cellDate = new Date(sheet.getRange(dateColumn + i).getValue());
//     cellDate.setHours(0, 0, 0, 0);  // Normalize the date for comparison

//     if (cellDate.getTime() === today.getTime()) {
//       // Set the corresponding cells in the specified columns to 0 if they are empty
//       columnsToSet.forEach(function(column) {
//         var cell = sheet.getRange(column + i);
//         if (cell.getValue() === "") { // Check if the cell is empty
//           cell.setValue(0);
//         }
//       });
      
//       // Update column O: O = N - recommended_rate if column O is empty
//       var cellO = sheet.getRange("O" + i);
//       if (cellO.getValue() === "") { // Check if column O is empty
//         var valueInN = sheet.getRange("N" + i).getValue(); // Get the value from column N
//         var newValueInO = valueInN - recommendedRate; // Calculate the new value for column O
//         cellO.setValue(newValueInO); // Set the new value in column O
//       }

//       rowUpdated = true; // Set the flag to true
//       break;  // Stop after setting the matching row for today
//     }
//   }

//   // Log the result
//   if (rowUpdated) {
//     Logger.log("Columns updated for today's date and column O updated.");
//   } else {
//     Logger.log("No matching date found for today.");
//   }
// }


// function testUpdateBackgroundColor() {
//   var spreadsheetId = "1op3uW3K-6i5ANWouEFrl9-HnOZBuLO7opq124-nRfhs"; // Your spreadsheet ID
//   var sheetName = "Learning Log"; // Your sheet name
//   var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
//   var dateColumn = "C"; // Column where dates are stored
//   var today = new Date();
//   today.setHours(0, 0, 0, 0); // Normalize to midnight
  
//   var lastRow = sheet.getLastRow();
//   var goldenYellow = '#FFD700'; // Golden yellow color code

//   for (var i = 15; i <= lastRow; i++) {
//     var cellDate = new Date(sheet.getRange(dateColumn + i).getValue());
//     cellDate.setHours(0, 0, 0, 0); // Normalize the date for comparison
    
//     if (cellDate.getTime() === today.getTime()) {
//       // Update background color for columns A, B, and C for the matching row
//       sheet.getRange("A" + i + ":C" + i).setBackground(goldenYellow);
//       break; // Stop after updating the matching row for today
//     }
//   }
// }

