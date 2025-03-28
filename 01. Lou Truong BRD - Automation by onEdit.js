// Global variable declaration and store information to reuse
const spreadSheetID = "1YRbL3n2iNv4hYz202pllVOzfO2Pg8V94AM5GoEX1HKY"; // Controller
const sheetName_1 = "Link builder"; // Controller
const sheetId_1 = 462100576; // Controller

// Function to get sheet id
function checkSheetId() {
  const ss_1 =
    SpreadsheetApp.openById(spreadSheetID).getSheetByName(sheetName_1); // Fixed
  try {
    if (ss_1) {
      console.log(`The actual ID for ${sheetName_1} is: ${ss_1.getSheetId()}`);
    } else {
      console.log(`The actual ID for ${sheetName_1} is not found`);
    }
  } catch (error) {
    console.error("Error message:", error.message);
  }
}
