// Global variable declaration and store information to reuse
// const spreadSheetID = "1YRbL3n2iNv4hYz202pllVOzfO2Pg8V94AM5GoEX1HKY"; // Controller
// const sheetName_1 = "Link builder"; // Controller
// const sheetId_1 = 462100576; // Controller

// Function to get sheet id
// function checkSheetId() {
//   const ss_1 =
//     SpreadsheetApp.openById(spreadSheetID).getSheetByName(sheetName_1); // Fixed
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
      linkBuilderIdGenerate(e);
    } else {
      console.error("onEdit triggered without event object 'e'.");
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}

// Function add id
function linkBuilderIdGenerate(e) {
  const sheetId_1 = 462100576; // Controller
  if (
    e &&
    e.range.getRow() > 1 &&
    e.range.getColumn() === 2 &&
    e.source.getSheetId() === sheetId_1 &&
    e.source.getActiveSheet().getRange(e.range.getRow(), 1).getValue() === ""
  ) {
    e.source.getActiveSheet().getRange(e.range.getRow(), 1).setValue("Test");
  }
}
