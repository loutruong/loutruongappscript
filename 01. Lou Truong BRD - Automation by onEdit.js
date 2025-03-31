// Global variable declaration and store information to reuse
// const spreadSheetID = "1YRbL3n2iNv4hYz202pllVOzfO2Pg8V94AM5GoEX1HKY"; // Controller
// const sheetName_1 = "Link builder"; // Controller
// const sheetId_1 = 462100576; // Controller
// const ss_1 = SpreadsheetApp.openById(spreadSheetID).getSheetByName(sheetName_1); // Fixed

// Feature to get sheet id
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

// Feature callback onEdit
function onEdit(e) {
  try {
    if (e) {
      addTimeStamp(e);
      adjustTimeStamp(e);
      linkBuilderIdGenerate(e);
      linkBuilderGenerate(e);
    } else {
      console.error("onEdit triggered without event object 'e'.");
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}

// Feature add timestamp
function addTimeStamp(e) {
  const sheetId_1 = 462100576; // Controller
  try {
    if (
      e &&
      e.range.getRow() > 1 &&
      e.range.getColumn() === 4 &&
      e.source.getSheetId() === sheetId_1 &&
      e.source.getActiveSheet().getRange(e.range.getRow(), 1).getValue() === ""
    ) {
      e.source
        .getActiveSheet()
        .getRange(e.range.getRow(), 2) // Controller
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

// Feature adjust timestamp
function adjustTimeStamp(e) {
  const sheetId_1 = 462100576; // Controller
  try {
    if (
      e &&
      e.range.getRow() > 1 &&
      e.range.getColumn() === 4 &&
      e.source.getSheetId() === sheetId_1
    ) {
      e.source
        .getActiveSheet()
        .getRange(e.range.getRow(), 3) // Controller
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

// Feature add id
function linkBuilderIdGenerate(e) {
  const sheetId_1 = 462100576; // Controller
  const idPrefix = "link"; // Controller the prefix for the ID
  try {
    if (
      e &&
      e.range.getRow() > 1 &&
      e.range.getColumn() === 4 &&
      e.source.getSheetId() === sheetId_1 &&
      e.source.getActiveSheet().getRange(e.range.getRow(), 1).getValue() === ""
    ) {
      let maxIdNumber = 0;
      const valueId = e.range
        .getSheet()
        .getRange(2, 1, e.range.getSheet().getMaxRows() - 1)
        .getValues();
      for (let i = 0; i < valueId.length; i++) {
        const value = valueId[i][0]; // Get the value from the cell [row][col=0]
        if (typeof value === "string" && value.startsWith(idPrefix)) {
          const curIdNumber = parseInt(value.substring(idPrefix.length), 10);
          if (!isNaN(curIdNumber)) {
            maxIdNumber = Math.max(maxIdNumber, curIdNumber); // Update the maximum number found
          }
        }
      }
      const nextIdNumber = maxIdNumber + 1;
      const newId = idPrefix + nextIdNumber;
      e.range.getSheet().getRange(e.range.getRow(), 1).setValue(newId);
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}

// Feature link builder
function linkBuilderGenerate(e) {
  // --- Start Configuration ---
  const targetSheetId = 462100576; // Controller
  const firstDataRow = 2; // Row number where your data starts (assuming Row 1 has headers)

  // Define column numbers based on structure:
  // A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12
  const urlCol = 4; // Column D: link_original (Base URL)
  const outputCol = 5; // Column E: link_full (Generated URL Output)
  const sourceCol = 7; // Column G: utm_source
  const mediumCol = 8; // Column H: utm_medium
  const campaignCol = 9; // Column I: utm_campaign
  const utmIdCol = 10; // Column J: utm_id
  const termCol = 11; // Column K: utm_term
  const contentCol = 12; // Column L: utm_content

  // Define which columns trigger the script when edited: We trigger on changes to the base URL or any UTM parameter
  const triggerColumns = [
    urlCol,
    sourceCol,
    mediumCol,
    campaignCol,
    utmIdCol,
    termCol,
    contentCol,
  ];
  // --- End Configuration ---
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const editedRow = range.getRow();
  const editedCol = range.getColumn();

  // Exit if the edit is not on the specified sheet, is in the header row, or is not in one of the trigger columns.
  if (
    sheet.getSheetId() != targetSheetId ||
    editedRow < firstDataRow ||
    !triggerColumns.includes(editedCol)
  ) {
    return;
  }

  // Get the base URL ('link_original') from the edited row
  let baseUrl = sheet.getRange(editedRow, urlCol).getValue().toString().trim();
  // Get the cell where the output ('link_full') will be written
  const outputCell = sheet.getRange(editedRow, outputCol);

  // If the base URL ('link_original') is empty, clear the output cell and stop
  if (!baseUrl) {
    outputCell.setValue("");
    return;
  }

  // Define the parameters and their corresponding columns
  const paramsConfig = [
    { name: "utm_source", col: sourceCol },
    { name: "utm_medium", col: mediumCol },
    { name: "utm_campaign", col: campaignCol },
    { name: "utm_id", col: utmIdCol }, // Note: 'utm_id' is the standard parameter name
    { name: "utm_term", col: termCol },
    { name: "utm_content", col: contentCol },
  ];

  let queryStringParts = []; // Array to hold "key=value" strings

  // Loop through each parameter configuration
  paramsConfig.forEach((paramInfo) => {
    const value = sheet
      .getRange(editedRow, paramInfo.col)
      .getValue()
      .toString()
      .trim(); // Only add the parameter if its value is not empty

    if (value) {
      queryStringParts.push(`${paramInfo.name}=${encodeURIComponent(value)}`); // Encode the value to make it URL-safe (handles spaces, special characters, etc.)
    }
  });

  // Assemble the final URL
  let finalUrl = baseUrl;
  if (queryStringParts.length > 0) {
    // Check if the base URL already contains a query string (a '?')
    if (baseUrl.includes("?")) {
      // If yes, append parameters with '&'
      finalUrl += "&" + queryStringParts.join("&");
    } else {
      // If no, start the query string with '?'
      finalUrl += "?" + queryStringParts.join("&");
    }
  }

  // Write the generated URL ('link_full') to the output column (Column C)
  outputCell.setValue(finalUrl);
}
