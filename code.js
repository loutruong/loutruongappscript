// Global variable declaration
const spreadSheetID = "1BYDLENNlvp6yarpq0UUPPwHYj15yRSK-qBOaS80Vs1w"; // Controller
const sheetName_1 = "Inc lv1 job tracking"; // Controller
const sheetId_1 = "281717863"; // Controller

// Function to get sheet id
function checkSheetId() {
  let ss = SpreadsheetApp.openById(spreadSheetID).getSheetByName(sheetName_1);
  try {
    if (ss) {
      console.log(`The actual ID for ${sheetName_1} is: ${ss.getSheetId()}`);
    } else {
      console.log(`The actual ID for ${sheetName_1} is not found`);
    }
  } catch (error) {
    console.error("Error message:", error.message);
  }
}

// Function callback onEdit
function onEdit(e) {
  timeStamp(e);
}

// Function handle timeStamp feature

function timeStamp(e) {
  if (e) {
    e.source.getAcitveSheet().getRange("A11:A12").setValue("Testing");
  }
}
