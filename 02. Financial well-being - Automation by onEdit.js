// Global variable declaration and store information to reuse
// const spreadSheetID = "1BYDLENNlvp6yarpq0UUPPwHYj15yRSK-qBOaS80Vs1w"; // Controller
// const sheetName_1 = "Inc lv1 job tracking"; // Controller
// const sheetId_1 = "281717863"; // Controller
// const ss_1 = SpreadsheetApp.openById(spreadSheetID).getSheetByName(sheetName_1); // Fixed
// Function to get sheet id
// function checkSheetId() {
//   try {
//     if (ss_1) {
//       console.log(`The actual ID for ${sheetName_1} is: ${ss_1.getSheetId()}`);
//     } else {
//       console.log(`The actual ID for ${sheetName_1} is not found`);
//     }
//   } catch (error) {
//     console.error("Error message:", error.message);
//   }
// }

// Function callback onEdit
function onEdit(e) {
  try {
    if (e) {
      addTimeStamp(e);
      adjustTimeStamp(e);
    } else {
      console.error("onEdit triggered without event object 'e'.");
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}

// Function Add time stamp
function addTimeStamp(e) {
  const sheetId_1 = 281717863; // Controller
  try {
    if (
      e &&
      e.range.getRow() > 1 &&
      e.range.getColumn() >=
        SpreadsheetApp.getActiveSpreadsheet()
          .getActiveSheet()
          .getRange("C1") // Controller - Chanage the column name
          .getColumn() &&
      e.range.getColumn() <=
        SpreadsheetApp.getActiveSpreadsheet()
          .getActiveSheet()
          .getRange("K1") // Controller - Chanage the column name
          .getColumn() &&
      e.source.getSheetId() === sheetId_1 &&
      e.source.getActiveSheet().getRange(e.range.getRow(), 1).getValue() === ""
    ) {
      e.source
        .getActiveSheet()
        .getRange(e.range.getRow(), 1) // Change the column write out position
        .setValue(
          Utilities.formatDate(
            new Date(),
            "Asia/Ho_Chi_Minh",
            "yyyy-MM-dd HH:mm:ss"
          )
        );
    } else {
      console.error("addTimeStamp error");
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}

// Function Adjust time stamp
function adjustTimeStamp(e) {
  const sheetId_1 = 281717863; // Controller
  try {
    if (
      e &&
      e.range.getRow() > 1 &&
      e.range.getColumn() >=
        SpreadsheetApp.getActiveSpreadsheet()
          .getActiveSheet()
          .getRange("C1") // Controller - Chanage the column name
          .getColumn() &&
      e.range.getColumn() <=
        SpreadsheetApp.getActiveSpreadsheet()
          .getActiveSheet()
          .getRange("K1") // Controller - Chanage the column name
          .getColumn() &&
      e.source.getSheetId() === sheetId_1
    ) {
      e.source
        .getActiveSheet()
        .getRange(e.range.getRow(), 2) // Change the column write out position
        .setValue(
          Utilities.formatDate(
            new Date(),
            "Asia/Ho_Chi_Minh",
            "yyyy-MM-dd HH:mm:ss"
          )
        );
    } else {
      console.error("addTimeStamp error");
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}
